import { normalizePresentation, validatePresentation, type PresentationDocument } from "@pptkit/core";
import { resolveNormalizedLayout } from "@pptkit/layout";
import type { ImportedPresentation } from "@pptkit/pptx-parser";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

export const maxFileBytes = 32 * 1024 * 1024;
export const pptxMimeType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export function download(content: BlobPart, filename: string, type: string) {
  const url = URL.createObjectURL(new Blob([content], { type }));
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export async function preview(document: PresentationDocument) {
  const svg = await renderPresentationToSvg(document);
  return { document, svg };
}

export function inspectDocument(document: PresentationDocument) {
  const ir = normalizePresentation(document);
  const layout = resolveNormalizedLayout(ir);
  return {
    validation: validatePresentation(document),
    irVersion: ir.irVersion,
    size: layout.size,
    slides: layout.slides.map(({ id, hidden, elements }, index) => ({
      id, hidden, notes: ir.slides[index]?.notes,
      elements: elements.map(({ id, type, box }) => ({ id, type, box })),
    })),
  };
}

async function withImport<T>(bytes: Uint8Array, use: (session: ImportedPresentation) => Promise<T>) {
  const { importPptx } = await import("@pptkit/pptx-parser");
  const session = await importPptx(bytes, {
    projection: "content",
    diagramPreview: "cache",
    editSaveProfile: "text-runs-v3",
    zipLimits: { maxInputBytes: maxFileBytes },
  });
  try {
    return await use(session);
  } finally {
    session.dispose();
  }
}

function findEditableRun({ document, report }: ImportedPresentation) {
  for (const binding of report.textRuns ?? []) {
    if (binding.status !== "bound" || binding.owner !== "text-box") continue;
    const slide = document.slides.find(({ id }) => id === binding.slideId);
    const element = slide?.elements.find(({ id }) => id === binding.elementId);
    if (!slide || element?.type !== "text" || typeof element.content === "string") continue;
    const run = element.content[binding.paragraphIndex]?.runs[binding.runIndex];
    if (run?.text) return { slide, element, binding, text: run.text };
  }
}

export function previewImport(bytes: Uint8Array) {
  return withImport(bytes, async (session) => {
    const { inspection, diagnostics, findings, diagramPreviews } = session.report;
    return {
      ...(await preview(session.document)),
      bytes,
      editableText: findEditableRun(session)?.text,
      inspection: {
        dialect: inspection.dialect,
        slides: inspection.slides.length,
        parts: inspection.package.parts.length,
        diagnostics, findings, diagramPreviews,
      },
    };
  });
}

export function saveImport(bytes: Uint8Array, replacement?: string) {
  return withImport(bytes, async (session) => {
    if (replacement !== undefined) {
      const target = findEditableRun(session);
      if (!target) throw new Error("This file has no supported text run to edit.");
      const edited = structuredClone(target.element);
      if (typeof edited.content === "string") throw new Error("The source run is unavailable.");
      edited.content[target.binding.paragraphIndex]!.runs[target.binding.runIndex]!.text = replacement;
      target.slide.replaceElement(edited.id, edited);
    }
    const prepared = await session.prepareSave({ mode: replacement === undefined ? "no-op" : "edit" });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(({ message }) => message).join("\n") || "PPTKit could not prepare a preservation save.");
    }
    try {
      return await session.save(prepared.plan);
    } finally {
      prepared.plan.dispose();
    }
  });
}
