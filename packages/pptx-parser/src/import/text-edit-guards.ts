import type { structureContext } from "../presentation/structure-context.js";
import type { PptxInspection } from "../types/structure.js";
import { PptxImportError } from "../types/import.js";
import type { TextEditGuard } from "./analyze-text-edits.js";
import type { TextRunBinding } from "./text-run-bindings.js";
import { paragraphStableOwners, stableParagraphTarget, textTargetId } from "./paragraph-targets.js";

/** Scan original branches, so inactive range references cannot disappear. */
export async function textEditGuards(
  context: Awaited<ReturnType<typeof structureContext>>, inspection: PptxInspection,
  maxGuards: number, checkpoint: () => Promise<void>,
  includeNotes = false,
  paragraphBindings?: readonly TextRunBinding[],
): Promise<TextEditGuard[]> {
  const guards: TextEditGuard[] = [];
  const add = (guard: TextEditGuard) => {
    if (guards.length >= maxGuards) throw new PptxImportError("import-limit-exceeded", "Text-edit source guards exceed maxReportFindings.");
    guards.push(guard);
  };
  if (inspection.kind !== "presentation" || inspection.macroEnabled || inspection.package.parts.some(part => /digital-signature|vbaProject/i.test(part.contentType) || /^\/_xmlsignatures(?:\/|$)/i.test(part.name))) {
    add({ diagnostic: { code: "save-text-package-policy", severity: "error", source: { part: inspection.mainPart }, message: "Text-run preservation saves require an ordinary non-macro presentation without recognized package signatures or VBA parts." } });
  }
  for (const part of inspection.package.parts) {
    if (!part.relationshipPart) continue;
    const root = context.rawXml.get(part.name)!;
    const stack = [root]; let index = 0;
    while (stack.length) {
      await checkpoint(); const node = stack.pop()!; index++;
      const type = node.attributes.find(a => !a.namespaceUri && a.local === "Type")?.value;
      if (node.namespaceUri === "http://schemas.openxmlformats.org/package/2006/relationships" && node.local === "Relationship" && type && (/\/digital-signature\//.test(type) || /\/(?:vbaProject|vbaProjectSignature|vbaData)$/.test(type))) {
        add({ diagnostic: { code: "save-text-package-policy", severity: "error", source: { part: part.name, elementIndex: index }, message: "A raw source relationship, including an inactive branch, declares signature/VBA semantics outside the text-run preservation profiles." } });
      }
      for (let i = node.children.length - 1; i >= 0; i--) {
        const child = node.children[i]!; if (child.kind === "element") stack.push(child);
      }
    }
  }
  const targetParts = new Set(inspection.slides.flatMap(slide => [slide.part, ...(includeNotes && slide.notesPart ? [slide.notesPart] : [])]));
  const paragraphOwners = paragraphBindings ? await paragraphStableOwners(context, paragraphBindings, maxGuards, checkpoint) : undefined;
  for (const part of targetParts) {
    const root = context.rawXml.get(part)!;
    const stack = [root]; let index = 0;
    while (stack.length) {
      await checkpoint();
      const node = stack.pop()!; index++;
      if (node.namespaceUri === context.p && node.local === "spTgt") {
        const id = node.attributes.find(a => !a.namespaceUri && a.local === "spid")?.value;
        const ownerId = textTargetId(id), validId = ownerId !== undefined;
        const opaque = node.attributes.some(a => a.namespaceUri !== "http://www.w3.org/2000/xmlns/" && (a.namespaceUri || a.local !== "spid"));
        const paragraphCount = ownerId === undefined ? undefined : paragraphOwners?.get(part)?.get(ownerId);
        const stable = paragraphCount !== undefined && stableParagraphTarget(node, context.p, paragraphCount);
        if (validId && !stable && (opaque || node.children.some(child => child.kind === "element"))) {
          add({ part, sourceObjectId: ownerId, diagnostic: {
            code: "save-text-range-reference", severity: "error", source: { part, elementIndex: index, sourceObjectId: id! },
            message: "A raw source shape target includes text/sub-object or opaque target semantics; its range remapping is outside the text-run preservation profiles.",
          } });
        } else if (!validId) {
          add({ part, diagnostic: { code: "save-text-range-reference", severity: "error", source: { part, elementIndex: index }, message: "A raw source shape target lacks a resolvable owner identity." } });
        }
      }
      for (let i = node.children.length - 1; i >= 0; i--) {
        const child = node.children[i]!; if (child.kind === "element") stack.push(child);
      }
    }
  }
  return guards;
}
