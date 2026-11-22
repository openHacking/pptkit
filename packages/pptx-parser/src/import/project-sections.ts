import type { ImportDiagnostic, ImportFinding, PptxSourceLocation } from "../types/import.js";
import type { PptxInspection } from "../types/structure.js";

type Finding = (feature: string, operation: ImportFinding["operation"], tier: ImportFinding["tier"], source: PptxSourceLocation, reason: string, slideId?: string) => void;

/** Core slide labels do not stand in for a native section collection. */
export async function projectSections(inspection: PptxInspection, finding: Finding, checkpoint: () => Promise<void>, collectionProjection = false) {
  const labels = new Map<number, string>(), diagnostics: ImportDiagnostic[] = [];
  const collection = inspection.sectionList;
  if (!collection) return { labels, diagnostics };
  let exportable = true, position = 0;
  if (collectionProjection) {
    const positions = new Map(inspection.slides.map((slide, index) => [slide.id, index]));
    for (const section of collection.sections) {
      const members = section.slides.map(slide => positions.get(slide.id)!).sort((a, b) => a - b);
      for (const index of members) { await checkpoint(); if (index !== position++) exportable = false; }
    }
    if (position !== inspection.slides.length) exportable = false;
  }
  for (const [index, section] of collection.sections.entries()) {
    await checkpoint();
    const source: PptxSourceLocation = {
      part: collection.part,
      path: `presentation/extLst/ext[@uri="${collection.extensionUri}"]/sectionLst/section[${index + 1}]`,
      ...(section.id !== undefined ? { sourceObjectId: section.id } : {}),
    };
    for (const operation of ["create", "import", "edit", "preview", "export", "no-op-save", "edit-save"] as const) {
      const tier = operation === "no-op-save" ? "preserved" : collectionProjection && (["create", "import", "edit"].includes(operation) || (operation === "export" && exportable)) ? "editable" : "unsupported";
      finding("section.native-collection", operation, tier, source,
        collectionProjection ? operation === "export" && !exportable ? "Fresh native export requires complete contiguous section membership in presentation order." : "Ordered section names and members map to independent Core identities; source GUIDs and nested extensions remain source-only, and source-preserving section edits are unsupported." : "Native section identity, ordered collection, empty entries and nested extensions remain in the source report/original bytes. Core slide section strings represent labels only.");
    }
    for (const [memberIndex, member] of section.slides.entries()) {
      await checkpoint();
      const slideId = `source-slide-${member.id}`;
      const at = { ...source, path: `${source.path}/sldIdLst/sldId[${memberIndex + 1}]`, relationshipId: member.relationshipId };
      if (section.name === undefined) {
        finding("section.label", "import", "unsupported", at, "The source section has no name; no label is inferred from its identity, order or neighboring sections.", slideId);
        continue;
      }
      labels.set(member.id, section.name);
      for (const operation of ["create", "import", "edit", "preview", "export", "no-op-save", "edit-save"] as const) {
        const tier = operation === "create" || operation === "import" ? "editable"
          : operation === "no-op-save" ? "preserved"
          : operation === "export" ? "fallback" : "unsupported";
        finding("section.label", operation, tier, at,
          operation === "export"
            ? "Fresh export writes the Core section string as PPTKit slide-extension metadata, not as a native PowerPoint section collection."
            : "The source member's decoded section name maps to its Core slide label by numeric slide identity. Duplicate/empty names do not merge source sections; no native collection editability is implied.",
          slideId);
      }
    }
  }
  diagnostics.push({
    ...(collectionProjection ? {
      code: "section-collection-projection",
      severity: "warning" as const,
      source: { part: collection.part },
      message: "Native sections map to independent Core identities and members in presentation order; original member order, source GUIDs and nested extensions remain source-only. Fresh export derives GUIDs from Core IDs and requires complete contiguous membership; source-preserving section edits are unsupported.",
    } : {
    code: "section-label-projection-only",
    severity: "warning" as const,
    source: { part: collection.part },
    message: "Native section member names map to Core slide labels only. Empty sections, GUID identity, collection order, nested extensions and native section editing/export remain source-only; slide movement retains labels without reconciling the native collection.",
    }),
  });
  return { labels, diagnostics };
}
