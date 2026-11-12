import { projectCustomProperty, sourceDateToUtc } from "./project-custom-properties.js";
import { createPresentation, type PresentationMetadataInput } from "@pptkit/core";
import type { PptxInspection } from "../types/structure.js";
import { PptxImportError, type ImportFinding, type ImportReport, type ImportSourceIdentity, type ImportDiagnostic, type PptxSourceLocation } from "../types/import.js";
import { freezeReport } from "./state.js";
import { projectSections } from "./project-sections.js";
import type { structureContext } from "../presentation/structure-context.js";
import { contentReader } from "./content-reader.js";

export interface MasterOwner { kind: "master" | "layout" | "slide"; id: string; name: string; part: string; masterId?: string; layoutPart?: string; showMasterShapes?: boolean }
export async function projectStructure(inspection: PptxInspection, maxFindings: number, checkpoint: () => Promise<void>, context?: Awaited<ReturnType<typeof structureContext>>, backgrounds = false) {
  const source = { part: inspection.mainPart, relationshipId: inspection.mainRelationshipId };
  if (inspection.kind !== "presentation" || inspection.macroEnabled) throw new PptxImportError("import-unsupported", "This import profile accepts non-macro presentation documents only, not templates or slideshows.", source);
  for (const part of inspection.package.parts) {
    await checkpoint();
    const type = part.contentType.toLowerCase();
    if (type.startsWith("application/vnd.openxmlformats-package.digital-signature-") || type === "application/vnd.ms-office.vbaproject" || type.includes("vbaprojectsignature")) throw new PptxImportError("import-unsupported", "Signed packages and VBA project parts require a separate import policy.", { part: part.name });
  }
  for (const r of inspection.package.relationships) {
    await checkpoint();
    if (r.type.startsWith("http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/") || /\/(?:vbaProject|vbaProjectSignature|vbaData)$/.test(r.type)) throw new PptxImportError("import-unsupported", "Signature/VBA relationships require a separate import policy.", { part: r.source, relationshipId: r.id });
  }
  if (!inspection.slideSize) throw new PptxImportError("import-unsupported", "A source slide size is required; structural import does not invent document dimensions.", source);
  const findings: ImportFinding[] = [], identities: ImportSourceIdentity[] = [], diagnostics: ImportDiagnostic[] = [
    { code: "structural-projection-only", severity: "warning", source, message: "Only source dimensions, eligible starting slide number, selected metadata, slide order/hidden state and named native section member labels are mapped to Core. Drawing content, themes, notes and native section collections/editing are not imported; preview/fresh export of this document is incomplete. Use the session no-op save to retain source content." },
  ];
  const finding = (feature: string, operation: ImportFinding["operation"], tier: ImportFinding["tier"], at: PptxSourceLocation, reason: string, slideId?: string) => {
    if (findings.length >= maxFindings) throw new PptxImportError("import-limit-exceeded", "Import capability report exceeds maxReportFindings.", at);
    findings.push({ id: `finding-${findings.length + 1}`, feature, operation, tier, source: at, reason, ...(slideId ? { slideId } : {}) });
  };
  // The exact-byte baseline retains even unreachable and unknown parts. This is
  // not a claim of schema validity, semantic mapping, or edit-save closure safety.
  for (const part of inspection.package.parts) {
    await checkpoint();
    const at = { part: part.name };
    finding("source-part", "no-op-save", "preserved", at, "The complete original package bytes, including this part, are retained for unchanged saves.");
    for (const operation of ["create", "import", "edit", "preview", "export", "edit-save"] as const) finding("source-part-content", operation, "unsupported", at, "Part content has no complete semantic mapping in the structural projection; separately reported structure fields are the only mapped subset.");
  }
  finding("source-package-envelope", "no-op-save", "preserved", { part: "/" }, "Original ZIP bytes retain content-type declarations, directory entries and envelope metadata without reserialization.");
  const metadata: PresentationMetadataInput = {};
  for (const [field, coreName] of [["title", "title"], ["author", "creator"], ["subject", "subject"], ["description", "description"], ["language", "language"]] as const) {
    const property = inspection.metadata.core?.properties.find((p) => p.name === coreName);
    if (property) {
      metadata[field] = property.value;
      finding(`metadata.${field}`, "import", "editable", { part: inspection.metadata.core!.part, path: `coreProperties/${coreName}` }, "Decoded source string is mapped to Core metadata without date/number conversion.");
    }
  }
  const keywords = inspection.metadata.core?.properties.find((p) => p.name === "keywords");
  if (keywords) {
    metadata.keywords = [keywords.value];
    finding("metadata.keywords", "import", "editable", { part: inspection.metadata.core!.part, path: "coreProperties/keywords" }, "The aggregate keyword string is preserved as one Core keyword entry; segment boundaries and per-segment language annotations are not represented.");
    if (keywords.language !== undefined || (keywords.keywordSegments?.length ?? 0) > 1 || keywords.keywordSegments?.some((segment) => segment.kind === "value" || segment.language !== undefined)) {
      finding("metadata.keyword-segments", "import", "unsupported", { part: inspection.metadata.core!.part, path: "coreProperties/keywords" }, "Core metadata stores keywords as plain strings and cannot preserve mixed-content boundaries or keyword and segment language annotations.");
    }
  }
  const revision = inspection.metadata.core?.properties.find((p) => p.name === "revision");
  if (revision) {
    const value = Number(revision.value);
    if (/^(?:0|[1-9]\d*)$/.test(revision.value) && Number.isSafeInteger(value)) {
      metadata.revision = value;
      finding("metadata.revision", "import", "editable", { part: inspection.metadata.core!.part, path: "coreProperties/revision" }, "A nonnegative revision within JavaScript's safe integer range maps to Core metadata.");
    } else {
      finding("metadata.revision", "import", "unsupported", { part: inspection.metadata.core!.part, path: "coreProperties/revision" }, "Revision is not a canonical nonnegative decimal integer within JavaScript's safe range; its XML-normalized source text remains available in the inspection report.");
    }
  }
  for (const [field, name] of [["createdAt", "created"], ["modifiedAt", "modified"]] as const) {
    const property = inspection.metadata.core?.properties.find(item => item.name === name);
    if (!property) continue;
    const canonical = sourceDateToUtc(property.value);
    if (canonical !== undefined) {
      metadata[field] = canonical;
      finding(`metadata.${field}`, "import", "editable", { part: inspection.metadata.core!.part, path: `coreProperties/${name}` }, "Explicitly zoned date with exact millisecond precision maps to canonical UTC; source lexical form remains available.");
    } else finding(`metadata.core.${name}`, "import", "unsupported", { part: inspection.metadata.core!.part, path: `coreProperties/${name}` }, "Date is unzoned, invalid, outside the supported year range or requires sub-millisecond precision; its source value remains inventoried.");
  }
  const company = inspection.metadata.extended?.properties.find((p) => p.name === "Company");
  if (company?.value.status === "scalar") {
    metadata.company = company.value.value;
    finding("metadata.company", "import", "editable", { part: inspection.metadata.extended!.part, path: "Properties/Company" }, "Decoded source company string is mapped to Core metadata.");
  }
  // Inspection is a source inventory, not an assertion that every property is
  // represented by Core. Report each unmapped property without flattening it.
  const mappedCore = new Set(["title", "creator", "subject", "description", "language", "keywords", "revision", "created", "modified"]);
  for (const property of inspection.metadata.core?.properties ?? []) {
    await checkpoint();
    if (!mappedCore.has(property.name)) finding(`metadata.core.${property.name}`, "import", "unsupported", { part: inspection.metadata.core!.part, path: `coreProperties/${property.name}` }, "This core property is retained in the inspection inventory but is not represented in the Core metadata projection.");
  }
  for (const property of inspection.metadata.extended?.properties ?? []) {
    await checkpoint();
    if (property.name !== "Company") finding(`metadata.extended.${property.name}`, "import", "unsupported", { part: inspection.metadata.extended!.part, path: `Properties/${property.name}` }, "This extended property is retained in the inspection inventory but is not represented in the Core metadata projection.");
  }
  if (inspection.metadata.custom) {
    const mapped: NonNullable<PresentationMetadataInput["customProperties"]>[number][] = [];
    const projected = [], nameCounts = new Map<string, number>();
    for (const property of inspection.metadata.custom.properties) {
      await checkpoint();
      const value = projectCustomProperty(property);
      projected.push(value);
      if (value) nameCounts.set(value.name, (nameCounts.get(value.name) ?? 0) + 1);
    }
    for (const [index, property] of inspection.metadata.custom.properties.entries()) {
      await checkpoint();
      const value = projected[index], unique = value && nameCounts.get(value.name) === 1;
      if (unique) mapped.push(value);
      finding("metadata.custom-property", "import", unique ? "editable" : "unsupported", { part: inspection.metadata.custom.part, path: "Properties/property", sourceObjectId: `${property.formatId}:${property.id}` }, unique ? "The typed source value maps to Core; source IDs and lexical form remain in inspection. This does not enable source-preserving edit save." : "This custom property remains in inspection and is not represented in the Core metadata projection; its type, value or decoded name may be unsupported or ambiguous.");
    }
    metadata.customProperties = mapped;
  }
  const declaredFirstSlideNumber = inspection.firstSlideNumberDeclared && inspection.firstSlideNumber >= 0 ? inspection.firstSlideNumber : undefined;
  if (inspection.firstSlideNumberDeclared) finding("first-slide-number", "import", declaredFirstSlideNumber === undefined ? "unsupported" : "editable", { ...source, path: "presentation/@firstSlideNum" }, declaredFirstSlideNumber === undefined ? "Negative source numbering remains source-only; Core accepts nonnegative numbers." : "The declared nonnegative starting number maps to Core; PPTX fresh export additionally enforces its PowerPoint-compatible range.");
  if (inspection.firstSlideNumberDeclared && (inspection.firstSlideNumber < 0 || inspection.firstSlideNumber > 9999)) finding("first-slide-number", "export", "unsupported", { ...source, path: "presentation/@firstSlideNum" }, "The source value is outside the PowerPoint-compatible fresh PPTX export range. It remains in source-preserving saves; choose an explicit supported value for fresh output.");
  const document = createPresentation({ id: "imported-presentation", metadata, ...(declaredFirstSlideNumber === undefined ? {} : { firstSlideNumber: declaredFirstSlideNumber }), size: { width: inspection.slideSize.widthEmu / 12700, height: inspection.slideSize.heightEmu / 12700 } });
  const masterOwners: MasterOwner[] | undefined = context ? [] : undefined;
  const layoutIds = new Map<string, string>();
  const layoutOwners = new Map<string, MasterOwner>();
  if (context && masterOwners) {
    const layoutInfo = new Map(inspection.layouts.map(layout => [layout.part, layout]));
    const layoutCounts = new Map<string, number>();
    for (const master of inspection.slideMasters) for (const layout of master.layouts) { await checkpoint(); layoutCounts.set(layout.part, (layoutCounts.get(layout.part) ?? 0) + 1); }
    for (const master of inspection.slideMasters) {
      await checkpoint();
      const id = `source-master-${master.id}`, name = `Source master ${master.id}`;
      document.defineSlideMaster({ id, name });
      masterOwners.push({ kind: "master", id, name, part: master.part });
      identities.push({ kind: "master", id, source: { part: master.part, sourceObjectId: String(master.id) } });
      for (const layout of master.layouts) {
        await checkpoint();
        if (layoutInfo.get(layout.part)?.masterPart !== master.part || layoutCounts.get(layout.part) !== 1) {
          finding("master-artwork.layout-ownership", "import", "unsupported", { part: layout.part }, "Layout ownership is absent, inconsistent or ambiguous; no Core inheritance path is fabricated.");
          continue;
        }
        const root = await context.root(layout.part, "sldLayout"), reader = contentReader(layout.part, context.p, context.a);
        const raw = reader.attr(root, "showMasterSp"), showMasterShapes = raw === undefined ? undefined : reader.bool(root, "showMasterSp");
        const layoutId = `source-layout-${master.id}-${layout.id}`, layoutName = layoutInfo.get(layout.part)?.name || `Source layout ${layout.id}`;
        document.defineSlideLayout({ id: layoutId, name: layoutName, masterId: id, ...(showMasterShapes !== undefined ? { showMasterShapes } : {}) });
        const owner: MasterOwner = { kind: "layout", id: layoutId, name: layoutName, part: layout.part, masterId: id, ...(showMasterShapes !== undefined ? { showMasterShapes } : {}) };
        masterOwners.push(owner); layoutOwners.set(layoutId, owner);
        layoutIds.set(layout.part, layoutId);
        identities.push({ kind: "layout", id: layoutId, source: { part: layout.part, sourceObjectId: String(layout.id) } });
      }
    }
  }
  identities.push({ kind: "presentation", id: document.id, source });
  finding("presentation-dimensions", "import", "editable", { ...source, path: "presentation/sldSz" }, "Exact source EMUs are converted to Core points using 12700 EMU per point.");
  const sections = await projectSections(inspection, finding, checkpoint);
  diagnostics.push(...sections.diagnostics);
  for (const slide of inspection.slides) {
    await checkpoint();
    const id = `source-slide-${slide.id}`, at = { part: slide.part, sourceObjectId: String(slide.id), relationshipId: slide.relationshipId };
    const section = sections.labels.get(slide.id);
    const layoutId = slide.layoutPart ? layoutIds.get(slide.layoutPart) : undefined;
    if (context && !layoutId) finding("master-artwork.slide-ownership", "import", "unsupported", at, "The slide has no verified listed layout/master path; Core's default layout is a projection fallback.", id);
    const slideRoot = context && layoutId ? await context.root(slide.part, "sld") : undefined;
    const reader = slideRoot && context ? contentReader(slide.part, context.p, context.a) : undefined;
    const showInheritedShapes = slideRoot && reader?.attr(slideRoot, "showMasterSp") !== undefined ? reader!.bool(slideRoot, "showMasterSp") : undefined;
    document.addSlide({ id, hidden: slide.hidden, ...(section !== undefined ? { section } : {}), ...(layoutId ? { layoutId } : {}), ...(showInheritedShapes !== undefined ? { showInheritedShapes } : {}) });
    if (backgrounds && masterOwners) {
      const layout = layoutId ? layoutOwners.get(layoutId) : undefined;
      masterOwners.push({ kind: "slide", id, name: id, part: slide.part, ...(layout?.masterId ? { masterId: layout.masterId, layoutPart: layout.part } : {}) });
    }
    identities.push({ kind: "slide", id, source: at });
    if (slide.transition) {
      const t = slide.transition, mapped = t.status === "mapped";
      if (mapped) document.slides[document.slides.length - 1]!.setTransition({ ...(t.type === "fade" ? { type: "fade" as const } : { type: t.type as "push" | "wipe", direction: t.direction! }), durationMs: t.durationMs!, onClick: t.onClick, ...(t.afterMs !== undefined ? { afterMs: t.afterMs } : {}) });
      for (const operation of ["import", "edit", "export", "preview", "no-op-save", "edit-save"] as const) finding("slide.transition", operation, operation === "no-op-save" ? "preserved" : operation === "preview" ? "fallback" : operation === "edit-save" || !mapped ? "unsupported" : "editable", { part: slide.part, path: "sld/transition" }, t.reason ?? "Plain fade/push/wipe with explicit millisecond duration; static preview displays the final slide. Transition edits are not admitted by preservation profiles.", id);
    }
    finding("slide-order-and-visibility", "import", "editable", at, "Source slide order and hidden state are mapped; no drawing elements or inherited layout are synthesized.", id);
  }
  for (const item of inspection.findings) diagnostics.push({ code: item.code, severity: "warning", message: item.message, source: { part: item.part, ...(item.slideId !== undefined ? { sourceObjectId: String(item.slideId) } : {}) } });
  diagnostics.push({ code: "core-projection-defaults", severity: "info", source, message: "Unmapped metadata/theme fields use Core defaults in this projection only; defaults are not inferred source values and are never written by no-op save." });
  const report: ImportReport = freezeReport({ projection: "structure", inspection, findings, identities, diagnostics });
  return { document, report, masterOwners };
}
