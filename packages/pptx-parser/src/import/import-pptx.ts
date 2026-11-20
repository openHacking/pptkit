import { planPropertyCreation } from "../preservation/plan-property-creation.js";
import { sourceDateToUtc } from "./project-custom-properties.js";
import { customPropertySources as inspectCustomPropertySources } from "./custom-property-sources.js";
import { planCustomPropertyCreation } from "../preservation/plan-custom-property-creation.js";
import { XmlSourceEditError } from "../preservation/xml-text-source.js";
import { snapshotOptions } from "../opc/read-package.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { readStructureFromContext } from "../presentation/read-structure.js";
import { structureContext } from "../presentation/structure-context.js";
import { PackageInspectionError } from "../types/archive.js";
import { PptxImportError, type ImportedPictureSource, type ImportedSlideVisibilitySource, type ImportedCoreMetadataSource, type ImportedPresentation, type PptxImportOptions, type PptxSessionLimits } from "../types/import.js";
import { projectStructure } from "./project-structure.js";
import { decodeLiteralEscapes, projectCustomProperty } from "./project-custom-properties.js";
import { createSession } from "./session.js";
import { freezeReport } from "./state.js";
import { projectContent } from "./project-content.js";
import type { ChartValueBinding } from "./chart-value-bindings.js";
import type { TextRunBinding } from "./text-run-bindings.js";
import { textEditGuards } from "./text-edit-guards.js";
import { captureDiagramLimits, captureDiagramTextComparison } from "../presentation/diagram-limits.js";

// Each metadata profile deliberately includes every capability introduced by earlier profiles.
const METADATA_SAVE_PROFILES: readonly NonNullable<PptxImportOptions["metadataSaveProfile"]>[] = [
  "core-strings-v1",
  "core-values-v2",
  "core-keywords-v3",
  "extended-company-v4",
  "core-insert-v5",
  "core-date-values-v6",
  "core-date-insert-v7",
  "core-create-v8",
  "core-revision-insert-v9",
  "core-keywords-insert-v10",
  "extended-company-insert-v11",
  "extended-create-v12",
  "metadata-remove-v13",
  "metadata-date-remove-v14",
  "metadata-keywords-remove-v15",
  "metadata-revision-remove-v16",
  "metadata-default-remove-v17",
];

export const DEFAULT_SESSION_LIMITS: Readonly<Required<PptxSessionLimits>> = Object.freeze({ maxNotesStyleWork: 1000000, maxTableStyleProperties: 500000, maxTableStyleDependencies: 1000000, maxTableStyleWork: 2000000, maxChartBindingWork: 1000000, maxDiagramPreviewWork: 1000000, maxSaveAnalysisWork: 16 * 1024 * 1024, maxSaveChanges: 10000, maxSaveValidationBytes: 128 * 1024 * 1024, maxSaveValidationWork: 512 * 1024 * 1024, maxTextStyleProperties: 500000, maxTextStyleDependencies: 1000000, maxPlaceholderEntries: 100000, maxExpandedChartWorkbookBytes: 16 * 1024 * 1024, maxTotalExpandedChartWorkbookBytes: 64 * 1024 * 1024, maxChartDataPoints: 100000, maxImportedAssetBytes: 16 * 1024 * 1024, maxTotalImportedAssetBytes: 64 * 1024 * 1024, maxReportFindings: 150000, maxSnapshotValues: 1000000, maxPreparedOutputBytes: 64 * 1024 * 1024, maxRetainedOutputBytes: 128 * 1024 * 1024, maxLivePlans: 16 });
/** Shared invocation-time option capture for byte and Node file imports. */
export function captureImportOptions(options: PptxImportOptions) {
  if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("importPptx requires explicit projection: structure or content.");
  const mode = options.projection;
  if (mode !== "structure" && mode !== "content") throw new TypeError("importPptx requires explicit projection: structure or content.");
  const pictureSourceProfile = options.pictureSourceProfile;
  const audioSourceProfile = options.audioSourceProfile;
  const videoSourceProfile = options.videoSourceProfile;
  const videoProjection = options.videoProjection;
  const audioProjection = options.audioProjection;
  if (audioProjection !== undefined && (audioProjection !== "wave-pcm-v1" || mode !== "content" || [options.editSaveProfile, options.descriptionEditSaveProfile, options.drawingNameSaveProfile, options.decorativeSaveProfile, options.slideVisibilitySaveProfile, options.metadataSaveProfile, options.customPropertySaveProfile, options.firstSlideNumberSaveProfile, options.chartEditSaveProfile, options.pictureEditSaveProfile, options.pictureSourceProfile].some(value => value !== undefined))) throw new TypeError("Audio projection requires wave-pcm-v1 content projection without edit-save or picture source profiles.");
  if (videoProjection !== undefined && (videoProjection !== "iso-bmff-v1" || mode !== "content" || [options.editSaveProfile, options.descriptionEditSaveProfile, options.drawingNameSaveProfile, options.decorativeSaveProfile, options.slideVisibilitySaveProfile, options.metadataSaveProfile, options.customPropertySaveProfile, options.firstSlideNumberSaveProfile, options.chartEditSaveProfile, options.pictureEditSaveProfile, options.pictureSourceProfile].some(value => value !== undefined))) throw new TypeError("Video projection requires iso-bmff-v1 content projection without edit-save or picture source profiles.");
  if (audioSourceProfile !== undefined && (audioSourceProfile !== "wave-pcm-v1" || mode !== "content")) throw new TypeError("Audio source inventory requires wave-pcm-v1 with content projection.");
  if (videoSourceProfile !== undefined && (videoSourceProfile !== "iso-bmff-v1" || mode !== "content")) throw new TypeError("Video source inventory requires iso-bmff-v1 with content projection.");
  if (pictureSourceProfile !== undefined && (pictureSourceProfile !== "embedded-v1" || mode !== "content")) throw new TypeError("Picture source inventory requires embedded-v1 with content projection.");
  const pictureEditSaveProfile = options.pictureEditSaveProfile;
  const profile = options.editSaveProfile;
  const descriptionEditSaveProfile = options.descriptionEditSaveProfile;
  const slideVisibilitySaveProfile = options.slideVisibilitySaveProfile;
  const customPropertySaveProfile = options.customPropertySaveProfile;
  if (customPropertySaveProfile !== undefined && (!(["scalar-v1", "scalar-and-name-v2", "remove-v3", "append-and-remove-v4", "create-and-append-v5"] as const).includes(customPropertySaveProfile) || mode !== "content" || [options.pictureSourceProfile, options.pictureEditSaveProfile, options.masterArtwork, options.diagramPreview, options.fieldTextProfile, options.textFontProfile, options.notesTextProfile, options.textTargetSaveProfile].some(value => value !== undefined) || descriptionEditSaveProfile !== undefined && descriptionEditSaveProfile !== "mixed-v5")) throw new TypeError("Custom scalar preservation requires content projection and compatible edit profiles.");
  const customPropertySourceProfile = options.customPropertySourceProfile;
  if (customPropertySourceProfile !== undefined && (customPropertySourceProfile !== "raw-v1" || mode !== "content")) throw new TypeError("Custom property source inventory requires content projection.");
  const metadataSaveProfile = options.metadataSaveProfile;
  if (metadataSaveProfile !== undefined && (!METADATA_SAVE_PROFILES.includes(metadataSaveProfile) || mode !== "content" || [options.pictureSourceProfile, pictureEditSaveProfile, options.masterArtwork, options.diagramPreview, options.fieldTextProfile, options.textFontProfile, options.notesTextProfile, options.textTargetSaveProfile].some(value => value !== undefined) || descriptionEditSaveProfile !== undefined && descriptionEditSaveProfile !== "mixed-v5")) throw new TypeError("Core metadata saves require content projection and compatible edit profiles.");
  if (slideVisibilitySaveProfile !== undefined && (slideVisibilitySaveProfile !== "attribute-v1" || mode !== "content" || [options.pictureSourceProfile, pictureEditSaveProfile, options.masterArtwork, options.diagramPreview, options.fieldTextProfile, options.textFontProfile, options.notesTextProfile, options.textTargetSaveProfile].some(value => value !== undefined) || descriptionEditSaveProfile !== undefined && descriptionEditSaveProfile !== "mixed-v5")) throw new TypeError("Slide visibility saves require content projection and compatible edit profiles.");
  const drawingNameSaveProfile = options.drawingNameSaveProfile;
  const decorativeSaveProfile = options.decorativeSaveProfile;
  if (decorativeSaveProfile !== undefined && (decorativeSaveProfile !== "existing-v1" || mode !== "content" || (descriptionEditSaveProfile === "mixed-v5" ? false : [profile, descriptionEditSaveProfile, drawingNameSaveProfile, options.firstSlideNumberSaveProfile, options.chartEditSaveProfile, slideVisibilitySaveProfile, metadataSaveProfile, customPropertySaveProfile].some(value => value !== undefined)) || [options.pictureSourceProfile, pictureEditSaveProfile, options.fieldTextProfile, options.masterArtwork, options.notesTextProfile, options.textFontProfile, options.textTargetSaveProfile, options.diagramPreview].some(value => value !== undefined))) throw new TypeError("Decorative preservation requires content projection; only mixed-v5 permits compatible text, property, chart and numbering edits.");
  if (drawingNameSaveProfile !== undefined && (drawingNameSaveProfile !== "existing-v1" || mode !== "content" || (descriptionEditSaveProfile === "mixed-v5" ? false : [profile, descriptionEditSaveProfile, options.firstSlideNumberSaveProfile, options.chartEditSaveProfile, slideVisibilitySaveProfile, metadataSaveProfile, customPropertySaveProfile].some(value => value !== undefined)) || [options.pictureSourceProfile, pictureEditSaveProfile, options.fieldTextProfile, options.masterArtwork, options.notesTextProfile, options.textFontProfile, options.textTargetSaveProfile, options.diagramPreview].some(value => value !== undefined))) throw new TypeError("Drawing name preservation requires content projection; only mixed-v5 permits compatible text, description, chart and numbering edits.");
  if (descriptionEditSaveProfile !== undefined && (!['existing-v1', 'attribute-v2', 'batch-v3', 'mixed-v4', 'mixed-v5'].includes(descriptionEditSaveProfile) || mode !== 'content' || (descriptionEditSaveProfile === 'mixed-v4' || descriptionEditSaveProfile === 'mixed-v5' ? profile !== 'text-runs-v3' : profile !== undefined) || (descriptionEditSaveProfile === 'mixed-v5' ? options.firstSlideNumberSaveProfile !== undefined && options.firstSlideNumberSaveProfile !== 'attribute-v2' : options.firstSlideNumberSaveProfile !== undefined) || options.pictureSourceProfile !== undefined || pictureEditSaveProfile !== undefined || options.fieldTextProfile !== undefined || options.masterArtwork !== undefined || options.notesTextProfile !== undefined || options.textFontProfile !== undefined || options.textTargetSaveProfile !== undefined || (descriptionEditSaveProfile !== 'mixed-v5' && (options.chartEditSaveProfile !== undefined || slideVisibilitySaveProfile !== undefined || metadataSaveProfile !== undefined || customPropertySaveProfile !== undefined)) || options.diagramPreview !== undefined)) throw new TypeError('Description save requires a supported content profile; mixed-v4/v5 require text-runs-v3 and only mixed-v5 admits chart and attribute-v2 numbering profiles.');
  const firstSlideNumberSaveProfile = options.firstSlideNumberSaveProfile;
  if (firstSlideNumberSaveProfile !== undefined && (!(["existing-v1", "attribute-v2"] as const).includes(firstSlideNumberSaveProfile) || mode !== "content" || (profile !== undefined && firstSlideNumberSaveProfile !== "attribute-v2") || (options.chartEditSaveProfile !== undefined && firstSlideNumberSaveProfile !== "attribute-v2") || (pictureEditSaveProfile !== undefined && firstSlideNumberSaveProfile !== "attribute-v2") || (pictureSourceProfile !== undefined && firstSlideNumberSaveProfile !== "attribute-v2") || options.masterArtwork !== undefined || options.diagramPreview !== undefined || (firstSlideNumberSaveProfile !== "attribute-v2" && (options.textFontProfile !== undefined || options.notesTextProfile !== undefined || options.textTargetSaveProfile !== undefined)))) throw new TypeError("firstSlideNumberSaveProfile requires content projection; optional text, chart, picture-inventory and text-projection profiles require attribute-v2; picture edit requires attribute-v2, and artwork profiles are excluded.");
  const fieldTextProfile = options.fieldTextProfile;
  if (fieldTextProfile !== undefined && (!["slide-number-v1", "date-and-number-v1", "date-time-and-number-v1"].includes(fieldTextProfile) || mode !== "content" || pictureEditSaveProfile !== undefined || options.masterArtwork !== undefined || options.diagramPreview !== undefined)) throw new TypeError("fieldTextProfile requires a supported field profile with content projection and no picture edit or artwork/diagram profiles.");
  const masterArtwork = options.masterArtwork;
  if (masterArtwork !== undefined && (!["pictures-v1", "artwork-v1", "artwork-v2", "artwork-v3", "artwork-v4", "artwork-v5", "artwork-v6", "artwork-v7", "artwork-v8", "artwork-v9", "artwork-v10", "artwork-v11"].includes(masterArtwork) || mode !== "content" || profile !== undefined || options.chartEditSaveProfile !== undefined)) throw new TypeError("masterArtwork requires a supported artwork profile with content projection and without preservation edit profiles.");
  const notesTextProfile = options.notesTextProfile;
  if (notesTextProfile !== undefined && (notesTextProfile !== "master-style-v1" || mode !== "content")) throw new TypeError("notesTextProfile requires master-style-v1 with content projection.");
  const textFontProfile = options.textFontProfile;
  if (textFontProfile !== undefined && (!["script-families-v1", "script-families-v2", "script-families-v3"].includes(textFontProfile) || mode !== "content")) throw new TypeError("textFontProfile requires script-families-v1, script-families-v2 or script-families-v3 with content projection.");
  const textTargetSaveProfile = options.textTargetSaveProfile;
  if (textTargetSaveProfile !== undefined && (textTargetSaveProfile !== "paragraph-stable-v1" || mode !== "content" || profile === undefined)) throw new TypeError("Paragraph target preservation requires content projection and an explicit text-run profile.");
  const chartEditSaveProfile = options.chartEditSaveProfile;
  if (chartEditSaveProfile !== undefined && (!['existing-values-v1', 'existing-values-v2', 'existing-values-v3', 'existing-values-and-labels-v4', 'existing-values-and-series-names-v5', 'existing-series-order-v6'].includes(chartEditSaveProfile) || mode !== "content")) throw new TypeError("Chart saves require a supported content chart preservation profile.");
  const diagramPreview = options.diagramPreview;
  if (diagramPreview !== undefined && (diagramPreview !== "cache" || mode !== "content")) throw new TypeError("diagramPreview requires cache with content projection.");
  const mediaPreview = options.mediaPreview;
  if (mediaPreview !== undefined && (mediaPreview !== "placeholder-v1" || mode !== "content" || [profile, descriptionEditSaveProfile, drawingNameSaveProfile, decorativeSaveProfile, slideVisibilitySaveProfile, metadataSaveProfile, customPropertySaveProfile, firstSlideNumberSaveProfile, chartEditSaveProfile, pictureEditSaveProfile].some(value => value !== undefined))) throw new TypeError("mediaPreview requires content projection without edit-save profiles.");
  if (pictureEditSaveProfile !== undefined && (pictureEditSaveProfile !== "replace-v1" || mode !== "content" || profile !== undefined || chartEditSaveProfile !== undefined || masterArtwork !== undefined || diagramPreview !== undefined || textFontProfile !== undefined || notesTextProfile !== undefined)) throw new TypeError("Picture replacement save requires content projection without other optional edit or projection profiles.");
  if (profile !== undefined && (!["text-runs-v1", "text-runs-v2", "text-runs-v3"].includes(profile) || mode !== "content")) throw new TypeError("editSaveProfile requires text-runs-v1, text-runs-v2 or text-runs-v3 with content projection.");
  const captured = snapshotOptions(options);
  const diagramLimits = captureDiagramLimits(options.diagramLimits);
  const diagramTextComparison = captureDiagramTextComparison(options.diagramTextComparison);
  const sessionLimits = options.sessionLimits;
  if (sessionLimits !== undefined && (!sessionLimits || typeof sessionLimits !== "object" || Array.isArray(sessionLimits))) throw new TypeError("Invalid session limits.");
  const limits = { ...DEFAULT_SESSION_LIMITS, ...sessionLimits };
  for (const [name, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError(`Invalid session limit ${name}.`);
  return { mode, profile, customPropertySaveProfile, customPropertySourceProfile, descriptionEditSaveProfile, drawingNameSaveProfile, decorativeSaveProfile, slideVisibilitySaveProfile, metadataSaveProfile, firstSlideNumberSaveProfile, masterArtwork, textFontProfile, notesTextProfile, textTargetSaveProfile, chartEditSaveProfile, diagramPreview, mediaPreview, diagramTextComparison, diagramLimits, captured, limits, pictureSourceProfile, audioSourceProfile, audioProjection, videoSourceProfile, videoProjection, pictureEditSaveProfile, fieldTextProfile };
}
export async function importPresentation(bytes: Uint8Array, options: PptxImportOptions): Promise<ImportedPresentation> {
  const { mode, profile, customPropertySaveProfile, customPropertySourceProfile, descriptionEditSaveProfile, drawingNameSaveProfile, decorativeSaveProfile, slideVisibilitySaveProfile, metadataSaveProfile, firstSlideNumberSaveProfile, masterArtwork, textFontProfile, notesTextProfile, textTargetSaveProfile, chartEditSaveProfile, diagramPreview, mediaPreview, diagramTextComparison, diagramLimits, captured, limits, pictureSourceProfile, audioSourceProfile, audioProjection, videoSourceProfile, videoProjection, pictureEditSaveProfile, fieldTextProfile } = captureImportOptions(options), signal = captured.signal;
  const metadataLevel = metadataSaveProfile === undefined ? 0 : METADATA_SAVE_PROFILES.indexOf(metadataSaveProfile) + 1;
  if (!(bytes instanceof Uint8Array)) throw new TypeError("Import input must be Uint8Array.");
  const view = byteViewInfo(bytes);
  if (view.shared) throw new PptxImportError("import-unsupported", "Import requires an ordinary byte buffer so its invocation-time snapshot cannot race another agent.");
  const abort = () => { if (signal?.aborted) throw new PptxImportError("operation-aborted", "PPTX import was aborted."); };
  abort();
  if (view.length > captured.zipLimits.maxInputBytes) throw new PptxImportError("import-limit-exceeded", "Input exceeds maxInputBytes.");
  // The session keeps this invocation-time copy, never the caller's mutable view.
  const source = copyByteView(bytes);
  try {
    const context = await structureContext(source, { ...(diagramTextComparison ? { diagramTextComparison } : {}), diagramLimits, limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(signal ? { signal } : {}) }, mode === "content");
    const inspection = await readStructureFromContext(context);
    let work = 0;
    const checkpoint = async () => { abort(); if (++work % 64 === 0) { await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort(); } };
    const projection = await projectStructure(inspection, limits.maxReportFindings, checkpoint, masterArtwork ? context : undefined, ["artwork-v5", "artwork-v6", "artwork-v7", "artwork-v8", "artwork-v9", "artwork-v10", "artwork-v11"].includes(masterArtwork ?? ""));
    const bindings: TextRunBinding[] | undefined = profile ? [] : undefined;
    const chartBindings: ChartValueBinding[] | undefined = chartEditSaveProfile ? [] : undefined;
    const chartLabelBindings: import("./chart-label-bindings.js").ChartLabelBinding[] | undefined = (chartEditSaveProfile === "existing-values-and-labels-v4" || chartEditSaveProfile === "existing-values-and-series-names-v5" || chartEditSaveProfile === "existing-series-order-v6") ? [] : undefined;
    const chartOrderBindings: import("./chart-series-order.js").ChartSeriesOrderBinding[] | undefined = chartEditSaveProfile === "existing-series-order-v6" ? [] : undefined;
    const descriptionBindings: import('./drawing-description-binding.js').DrawingDescriptionBinding[] | undefined = descriptionEditSaveProfile || drawingNameSaveProfile || decorativeSaveProfile ? [] : undefined;
    const pictureSources: ImportedPictureSource[] | undefined = pictureSourceProfile || pictureEditSaveProfile ? [] : undefined;
    if (mode === "content") projection.report = await projectContent(projection.document, projection.report, context, limits, checkpoint, { limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(signal ? { signal } : {}) }, bindings, profile, diagramPreview, chartBindings, textFontProfile, notesTextProfile, projection.masterOwners, masterArtwork !== undefined && masterArtwork !== "pictures-v1", ["artwork-v2", "artwork-v3", "artwork-v4", "artwork-v5", "artwork-v6", "artwork-v7", "artwork-v8", "artwork-v9", "artwork-v10", "artwork-v11"].includes(masterArtwork ?? ""), ["artwork-v3", "artwork-v4", "artwork-v5", "artwork-v6", "artwork-v7", "artwork-v8", "artwork-v9", "artwork-v10", "artwork-v11"].includes(masterArtwork ?? ""), ["artwork-v4", "artwork-v5", "artwork-v6", "artwork-v7", "artwork-v8", "artwork-v9", "artwork-v10", "artwork-v11"].includes(masterArtwork ?? ""), ["artwork-v5", "artwork-v6", "artwork-v7", "artwork-v8", "artwork-v9", "artwork-v10", "artwork-v11"].includes(masterArtwork ?? ""), masterArtwork === "artwork-v6" || masterArtwork === "artwork-v7" || masterArtwork === "artwork-v8" || masterArtwork === "artwork-v9" || (masterArtwork === "artwork-v10" || masterArtwork === "artwork-v11"), masterArtwork === "artwork-v7" || masterArtwork === "artwork-v8" || masterArtwork === "artwork-v9" || (masterArtwork === "artwork-v10" || masterArtwork === "artwork-v11"), masterArtwork === "artwork-v8" || masterArtwork === "artwork-v9" || (masterArtwork === "artwork-v10" || masterArtwork === "artwork-v11"), masterArtwork === "artwork-v9" || (masterArtwork === "artwork-v10" || masterArtwork === "artwork-v11"), (masterArtwork === "artwork-v10" || masterArtwork === "artwork-v11"), masterArtwork === "artwork-v11", pictureSources, fieldTextProfile, descriptionBindings, descriptionEditSaveProfile === 'attribute-v2' || descriptionEditSaveProfile === 'batch-v3' || descriptionEditSaveProfile === 'mixed-v4' || descriptionEditSaveProfile === 'mixed-v5', drawingNameSaveProfile !== undefined && descriptionEditSaveProfile === undefined, drawingNameSaveProfile !== undefined && descriptionEditSaveProfile === 'mixed-v5', decorativeSaveProfile !== undefined, decorativeSaveProfile !== undefined && descriptionEditSaveProfile === undefined, chartEditSaveProfile === "existing-values-v2" || chartEditSaveProfile === "existing-values-v3" || chartEditSaveProfile === "existing-values-and-labels-v4" || chartEditSaveProfile === "existing-values-and-series-names-v5" || chartEditSaveProfile === "existing-series-order-v6", chartEditSaveProfile === "existing-values-v3" || chartEditSaveProfile === "existing-values-and-labels-v4" || chartEditSaveProfile === "existing-values-and-series-names-v5" || chartEditSaveProfile === "existing-series-order-v6", chartLabelBindings, chartEditSaveProfile === "existing-values-and-series-names-v5" || chartEditSaveProfile === "existing-series-order-v6", chartOrderBindings, false, mediaPreview, audioSourceProfile, audioProjection, videoSourceProfile, videoProjection);
    const metadataBindings: import("./analyze-text-edits.js").CoreMetadataBinding[] | undefined = metadataSaveProfile ? [] : undefined;
    const coreMetadataSources: ImportedCoreMetadataSource[] | undefined = metadataSaveProfile ? [] : undefined;
    let coreCreation: { part: string; create: { dialect: "strict" | "transitional"; relationshipId: string } } | undefined;
    if (metadataLevel >= 8 && !inspection.metadata.core) {
      try {
        const dialect = context.entry.dialect, plan = await planPropertyCreation(source, "core", dialect, { limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(signal ? { signal } : {}) }, limits.maxPreparedOutputBytes, limits.maxSaveChanges);
        coreCreation = { part: plan.part, create: { dialect, relationshipId: plan.relationshipId } };
      } catch (cause) { if (!(cause instanceof XmlSourceEditError) || !["xml-edit-unsupported", "xml-edit-conflict"].includes(cause.code)) throw cause; }
    }
    let extendedCreation: typeof coreCreation;
    if (metadataLevel >= 12 && !inspection.metadata.extended) {
      try {
        const dialect = context.entry.dialect, plan = await planPropertyCreation(source, "extended", dialect, { limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(signal ? { signal } : {}) }, limits.maxPreparedOutputBytes, limits.maxSaveChanges);
        extendedCreation = { part: plan.part, create: { dialect, relationshipId: plan.relationshipId } };
      } catch (cause) { if (!(cause instanceof XmlSourceEditError) || !["xml-edit-unsupported", "xml-edit-conflict"].includes(cause.code)) throw cause; }
    }
    if (metadataBindings && coreMetadataSources) {
      const core = inspection.metadata.core, root = core ? context.rawXml.get(core.part) : undefined;
      const validRoot = Boolean(root && root.namespaceUri === "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" && root.local === "coreProperties");
      const rawRoot = root && core && validRoot ? await context.sourceElement(core.part, root) : undefined;
      for (const [field, local] of [["title", "title"], ["author", "creator"], ["subject", "subject"], ["description", "description"], ["language", "language"]] as const) {
        await checkpoint();
        if (coreMetadataSources.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Core metadata sources exceed maxReportFindings.", { part: core?.part ?? "/" });
        const property = core?.properties.find(item => item.name === local);
        const nodes = root?.children.filter(item => item.kind === "element" && item.namespaceUri === "http://purl.org/dc/elements/1.1/" && item.local === local) ?? [];
        const node = nodes[0], directText = node?.kind === "element" && node.children.every(child => child.kind === "text" || child.kind === "cdata");
        const raw = core && node && node.kind === "element" ? await context.sourceElement(core.part, node) : undefined;
        const bound = Boolean(core && root && rawRoot && raw && property && nodes.length === 1 && directText && rawRoot.elementIndex === 1 && rawRoot.parent === undefined && raw.parent === root && node!.children.map(child => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("") === property.value);
        const insertable = Boolean((metadataLevel >= 5) && core && root && rawRoot && rawRoot.elementIndex === 1 && rawRoot.parent === undefined && !property && nodes.length === 0 || coreCreation && !property && nodes.length === 0);
        coreMetadataSources.push({ field, ...(core || coreCreation ? { source: { part: core?.part ?? coreCreation!.part, ...(raw ? { elementIndex: raw.elementIndex } : {}), path: `coreProperties/${local}` } } : {}), status: bound ? "bound" : insertable ? "insertable" : "unmapped", ...(!bound && !insertable ? { reason: "A direct source property or an eligible existing core-properties part is required." } : {}) });
        if (bound && core && raw && rawRoot && property && node?.kind === "element") metadataBindings.push({ field, part: core.part, elementIndex: raw.elementIndex, rootElementIndex: rawRoot.elementIndex, namespace: node.namespaceUri, local, baseline: property.value, ...((metadataLevel >= 13) && (field === "subject" || field === "description" || metadataLevel >= 17 && (field === "title" || field === "author" || field === "language")) && node.attributes.every(attribute => attribute.namespaceUri === "http://www.w3.org/2000/xmlns/") && !context.reports.some(report => report.part === core.part) ? { removable: true as const } : {}) });
        if (insertable && core && rawRoot) metadataBindings.push({ field, part: core.part, elementIndex: rawRoot.elementIndex, rootElementIndex: rawRoot.elementIndex, namespace: "http://purl.org/dc/elements/1.1/", local, baseline: field === "title" ? "PPTKit Presentation" : field === "author" ? "PPTKit" : field === "language" ? "en-US" : undefined, insert: true });
        if (insertable && coreCreation) metadataBindings.push({ field, part: coreCreation.part, elementIndex: 1, rootElementIndex: 1, namespace: "http://purl.org/dc/elements/1.1/", local, baseline: field === "title" ? "PPTKit Presentation" : field === "author" ? "PPTKit" : field === "language" ? "en-US" : undefined, insert: true, create: coreCreation.create });
      }
      if (metadataLevel >= 2) {
        const property = core?.properties.find(item => item.name === "revision"), namespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        const nodes = root?.children.filter(item => item.kind === "element" && item.namespaceUri === namespace && item.local === "revision") ?? [];
        const node = nodes[0], raw = core && node && node.kind === "element" ? await context.sourceElement(core.part, node) : undefined;
        const lexical = property?.value, numeric = lexical === undefined ? NaN : Number(lexical);
        const bound = Boolean(core && root && rawRoot && raw && node?.kind === "element" && nodes.length === 1 && rawRoot.elementIndex === 1 && rawRoot.parent === undefined && raw.parent === root && node.children.every(child => child.kind === "text" || child.kind === "cdata") && node.children.map(child => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("") === lexical && lexical !== undefined && /^(?:0|[1-9]\d*)$/.test(lexical) && Number.isSafeInteger(numeric) && projection.document.metadata.revision === numeric);
        if (coreMetadataSources.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Core metadata sources exceed maxReportFindings.", { part: core?.part ?? "/" });
        const insertable = Boolean((metadataLevel >= 9) && !property && nodes.length === 0 && (core && root && rawRoot && rawRoot.elementIndex === 1 && rawRoot.parent === undefined || coreCreation));
        coreMetadataSources.push({ field: "revision", ...(core || coreCreation ? { source: { part: core?.part ?? coreCreation!.part, ...(raw ? { elementIndex: raw.elementIndex } : {}), path: "coreProperties/revision" } } : {}), status: bound ? "bound" : insertable ? "insertable" : "unmapped", ...(!bound && !insertable ? { reason: "An existing canonical revision or eligible missing source element is required." } : {}) });
        if (bound && core && raw && rawRoot && node?.kind === "element" && lexical !== undefined) metadataBindings.push({ field: "revision", part: core.part, elementIndex: raw.elementIndex, rootElementIndex: rawRoot.elementIndex, namespace, local: "revision", baseline: lexical, ...((metadataLevel >= 16) && node.attributes.every(attribute => attribute.namespaceUri === "http://www.w3.org/2000/xmlns/") && !context.reports.some(report => report.part === core.part) ? { removable: true as const } : {}) });
        if (insertable && core && rawRoot) metadataBindings.push({ field: "revision", part: core.part, elementIndex: rawRoot.elementIndex, rootElementIndex: rawRoot.elementIndex, namespace, local: "revision", baseline: "1", insert: true });
        if (insertable && coreCreation) metadataBindings.push({ field: "revision", part: coreCreation.part, elementIndex: 1, rootElementIndex: 1, namespace, local: "revision", baseline: "1", insert: true, create: coreCreation.create });
      }
      if (metadataLevel >= 3) {
        const property = core?.properties.find(item => item.name === "keywords"), namespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        const nodes = root?.children.filter(item => item.kind === "element" && item.namespaceUri === namespace && item.local === "keywords") ?? [];
        const node = nodes[0], raw = core && node && node.kind === "element" ? await context.sourceElement(core.part, node) : undefined;
        const bound = Boolean(core && root && rawRoot && raw && node?.kind === "element" && nodes.length === 1 && rawRoot.elementIndex === 1 && rawRoot.parent === undefined && raw.parent === root && property && property.language === undefined && (property.keywordSegments?.length ?? 0) <= 1 && property.keywordSegments?.every(segment => segment.kind === "text" && segment.language === undefined) && node.children.every(child => child.kind === "text" || child.kind === "cdata") && node.children.map(child => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("") === property.value && projection.document.metadata.keywords?.length === 1 && projection.document.metadata.keywords[0] === property.value);
        if (coreMetadataSources.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Core metadata sources exceed maxReportFindings.", { part: core?.part ?? "/" });
        const insertable = Boolean((metadataLevel >= 10) && !property && nodes.length === 0 && (core && root && rawRoot && rawRoot.elementIndex === 1 && rawRoot.parent === undefined || coreCreation));
        coreMetadataSources.push({ field: "keywords", ...(core || coreCreation ? { source: { part: core?.part ?? coreCreation!.part, ...(raw ? { elementIndex: raw.elementIndex } : {}), path: "coreProperties/keywords" } } : {}), status: bound ? "bound" : insertable ? "insertable" : "unmapped", ...(!bound && !insertable ? { reason: "An unsegmented source keyword or eligible missing source element is required." } : {}) });
        if (bound && core && raw && rawRoot && node?.kind === "element" && property) metadataBindings.push({ field: "keywords", part: core.part, elementIndex: raw.elementIndex, rootElementIndex: rawRoot.elementIndex, namespace, local: "keywords", baseline: property.value, ...((metadataLevel >= 15) && node.attributes.every(attribute => attribute.namespaceUri === "http://www.w3.org/2000/xmlns/") && !context.reports.some(report => report.part === core.part) ? { removable: true as const } : {}) });
        if (insertable && core && rawRoot) metadataBindings.push({ field: "keywords", part: core.part, elementIndex: rawRoot.elementIndex, rootElementIndex: rawRoot.elementIndex, namespace, local: "keywords", baseline: undefined, insert: true });
        if (insertable && coreCreation) metadataBindings.push({ field: "keywords", part: coreCreation.part, elementIndex: 1, rootElementIndex: 1, namespace, local: "keywords", baseline: undefined, insert: true, create: coreCreation.create });
      }
      if (metadataLevel >= 4) {
        const extended = inspection.metadata.extended, namespace = context.entry.dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/extendedProperties" : "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        const rawRoot = extended ? context.rawXml.get(extended.part) : undefined, rootIdentity = extended && rawRoot ? await context.sourceElement(extended.part, rawRoot) : undefined;
        const property = extended?.properties.find(item => item.name === "Company"), nodes = rawRoot?.children.filter(item => item.kind === "element" && item.namespaceUri === namespace && item.local === "Company") ?? [];
        const node = nodes[0], raw = extended && node && node.kind === "element" ? await context.sourceElement(extended.part, node) : undefined;
        const bound = Boolean(extended && rawRoot && rootIdentity && raw && node?.kind === "element" && rawRoot.namespaceUri === namespace && rawRoot.local === "Properties" && rootIdentity.elementIndex === 1 && rootIdentity.parent === undefined && raw.parent === rawRoot && nodes.length === 1 && !context.reports.some(report => report.part === extended.part) && property?.value.status === "scalar" && node.children.every(child => child.kind === "text" || child.kind === "cdata") && node.children.map(child => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("") === property.value.value && projection.document.metadata.company === property.value.value);
        if (coreMetadataSources.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Metadata sources exceed maxReportFindings.", { part: extended?.part ?? "/" });
        const insertable = Boolean(((metadataLevel >= 11) && extended && rawRoot && rootIdentity && rawRoot.namespaceUri === namespace && rawRoot.local === "Properties" && rootIdentity.elementIndex === 1 && rootIdentity.parent === undefined && !context.reports.some(report => report.part === extended.part) || extendedCreation) && !property && nodes.length === 0);
        coreMetadataSources.push({ field: "company", ...(extended || extendedCreation ? { source: { part: extended?.part ?? extendedCreation!.part, ...(raw ? { elementIndex: raw.elementIndex } : {}), path: "Properties/Company" } } : {}), status: bound ? "bound" : insertable ? "insertable" : "unmapped", ...(!bound && !insertable ? { reason: "An eligible extended-properties part without compatibility processing is required." } : {}) });
        if (bound && extended && raw && rootIdentity && property?.value.status === "scalar") metadataBindings.push({ field: "company", part: extended.part, elementIndex: raw.elementIndex, rootElementIndex: rootIdentity.elementIndex, namespace, local: "Company", baseline: property.value.value, extendedRootNamespace: namespace, ...((metadataLevel >= 13) && node?.kind === "element" && node.attributes.every(attribute => attribute.namespaceUri === "http://www.w3.org/2000/xmlns/") ? { removable: true as const } : {}) });
        if (insertable && extended && rootIdentity) metadataBindings.push({ field: "company", part: extended.part, elementIndex: rootIdentity.elementIndex, rootElementIndex: rootIdentity.elementIndex, namespace, local: "Company", baseline: undefined, insert: true, extendedRootNamespace: namespace });
        if (insertable && extendedCreation) metadataBindings.push({ field: "company", part: extendedCreation.part, elementIndex: 1, rootElementIndex: 1, namespace, local: "Company", baseline: undefined, insert: true, extendedRootNamespace: namespace, create: extendedCreation.create });
      }
      if (metadataLevel >= 6) {
        for (const [field, local] of [["createdAt", "created"], ["modifiedAt", "modified"]] as const) {
          await checkpoint();
          if (coreMetadataSources.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Core metadata sources exceed maxReportFindings.", { part: core?.part ?? "/" });
          const property = core?.properties.find(item => item.name === local), nodes = root?.children.filter(item => item.kind === "element" && item.namespaceUri === "http://purl.org/dc/terms/" && item.local === local) ?? [];
          const node = nodes[0], raw = core && node?.kind === "element" ? await context.sourceElement(core.part, node) : undefined;
          const canonical = property ? sourceDateToUtc(property.value) : undefined;
          const bound = Boolean(core && root && rawRoot && raw && property && canonical !== undefined && nodes.length === 1 && node?.kind === "element" && node.children.every(child => child.kind === "text" || child.kind === "cdata") && rawRoot.elementIndex === 1 && rawRoot.parent === undefined && raw.parent === root && node.children.map(child => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("") === property.value);
          const insertable = Boolean((metadataLevel >= 7) && core && root && rawRoot && rawRoot.elementIndex === 1 && rawRoot.parent === undefined && !property && nodes.length === 0 || coreCreation && !property && nodes.length === 0);
          coreMetadataSources.push({ field, ...(core || coreCreation ? { source: { part: core?.part ?? coreCreation!.part, ...(raw ? { elementIndex: raw.elementIndex } : {}), path: `coreProperties/${local}` } } : {}), status: bound ? "bound" : insertable ? "insertable" : "unmapped", ...(!bound && !insertable ? { reason: "An exactly representable direct date or an eligible existing core-properties part is required." } : {}) });
          if (bound && core && raw && rawRoot && property && canonical !== undefined) metadataBindings.push({ field, part: core.part, elementIndex: raw.elementIndex, rootElementIndex: rawRoot.elementIndex, namespace: "http://purl.org/dc/terms/", local, baseline: canonical, lexical: property.value, ...((metadataLevel >= 14) && node?.kind === "element" && node.attributes.every(attribute => attribute.namespaceUri === "http://www.w3.org/2000/xmlns/" || attribute.namespaceUri === "http://www.w3.org/2001/XMLSchema-instance" && attribute.local === "type") && !context.reports.some(report => report.part === core.part) ? { removable: true as const } : {}) });
          if (insertable && core && rawRoot) metadataBindings.push({ field, part: core.part, elementIndex: rawRoot.elementIndex, rootElementIndex: rawRoot.elementIndex, namespace: "http://purl.org/dc/terms/", local, baseline: undefined, insert: true });
          if (insertable && coreCreation) metadataBindings.push({ field, part: coreCreation.part, elementIndex: 1, rootElementIndex: 1, namespace: "http://purl.org/dc/terms/", local, baseline: undefined, insert: true, create: coreCreation.create });
        }
      }
    }
    const slideVisibilityBindings: import("./analyze-text-edits.js").SlideVisibilityBinding[] | undefined = slideVisibilitySaveProfile ? [] : undefined;
    const slideVisibilitySources: ImportedSlideVisibilitySource[] | undefined = slideVisibilitySaveProfile ? [] : undefined;
    if (slideVisibilityBindings && slideVisibilitySources) for (const slide of inspection.slides) {
      await checkpoint();
      const effective = await context.root(slide.part, "sld"), raw = await context.sourceElement(slide.part, effective);
      const attribute = raw.node.attributes.find(item => item.namespaceUri === "" && item.local === "show"), normalized = attribute?.value.trim();
      const mapped = normalized === undefined || normalized === "1" || normalized === "true" ? false : normalized === "0" || normalized === "false" ? true : undefined;
      const slideId = `source-slide-${slide.id}`, bound = raw.parent === undefined && raw.elementIndex === 1 && raw.node.namespaceUri === context.p && raw.node.local === "sld" && mapped === slide.hidden;
      if (slideVisibilitySources.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Slide visibility sources exceed maxReportFindings.", { part: slide.part });
      slideVisibilitySources.push({ slideId, source: { part: slide.part, elementIndex: raw.elementIndex, path: "sld/@show", sourceObjectId: String(slide.id) }, status: bound ? "bound" : "unmapped", ...(!bound ? { reason: "Raw slide root and visibility do not match the effective slide projection." } : {}) });
      if (bound) slideVisibilityBindings.push({ slideId, part: slide.part, elementIndex: raw.elementIndex, namespace: context.p, ...(attribute ? { expected: attribute.value } : {}), baseline: slide.hidden });
    }
    const customPropertySources = customPropertySourceProfile || customPropertySaveProfile ? await inspectCustomPropertySources(context, limits, checkpoint) : undefined;
    const customPropertyBindings: import("./analyze-text-edits.js").CustomPropertyBinding[] | undefined = customPropertySaveProfile ? [] : undefined;
    let customInsertionSource: import("./analyze-text-edits.js").CustomPropertyInsertionSource | undefined;
    if (customPropertyBindings && customPropertySources && context.metadata.custom) {
      const projected = projection.document.metadata.customProperties ?? [];
      const projectedIndexes = new Map(projected.map((item, index) => [item.name, index]));
      const editableSources = new Set(projection.report.findings.filter(item => item.feature === "metadata.custom-property" && item.tier === "editable").map(item => item.source.sourceObjectId));
      const reservedNames = new Set<string>(); let unknownReservedName = false;
      for (const property of context.metadata.custom.properties) {
        await checkpoint();
        if (editableSources.has(`${property.formatId}:${property.id}`) || property.name === undefined) continue;
        const name = decodeLiteralEscapes(property.name);
        if (name === undefined) unknownReservedName = true; else reservedNames.add(name);
      }
      if (customPropertySaveProfile === "append-and-remove-v4" || customPropertySaveProfile === "create-and-append-v5") {
        const part = context.metadata.custom.part, root = context.rawXml.get(part), rawRoot = root ? await context.sourceElement(part, root) : undefined;
        if (root && rawRoot && rawRoot.elementIndex === 1 && rawRoot.parent === undefined && !unknownReservedName && !context.reports.some(report => report.part === part)) {
          customInsertionSource = { part, rootNamespace: root.namespaceUri, valueNamespace: context.entry.dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes" : "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", usedIds: new Set(context.metadata.custom.properties.map(property => property.id)), reservedNames };
        }
      }
      for (const [index, property] of context.metadata.custom.properties.entries()) {
        await checkpoint();
        const source = customPropertySources[index];
        if (!source || source.status !== "bound" || source.valueElementIndex === undefined || source.source.elementIndex === undefined || !editableSources.has(source.source.sourceObjectId) || property.value.status !== "scalar") continue;
        const mapped = projectCustomProperty(property), projectedIndex = mapped ? projectedIndexes.get(mapped.name) : undefined;
        if (!mapped || projectedIndex === undefined) continue;
        customPropertyBindings.push({ index: projectedIndex, part: context.metadata.custom.part, propertyElementIndex: source.source.elementIndex, valueElementIndex: source.valueElementIndex, rootNamespace: context.rawXml.get(context.metadata.custom.part)!.namespaceUri, valueNamespace: context.entry.dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes" : "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", local: property.valueType, baseline: mapped.value, lexical: property.value.value, ...(customPropertySaveProfile !== "scalar-v1" && !unknownReservedName ? { baselineName: mapped.name, lexicalName: property.name!, reservedNames } : {}), sourceObjectId: source.source.sourceObjectId! });
      }
    }
    if (customPropertySaveProfile === "create-and-append-v5" && !context.metadata.custom) {
      try {
        const dialect = context.entry.dialect;
        const planned = await planCustomPropertyCreation(source, dialect, { limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(signal ? { signal } : {}) }, limits.maxPreparedOutputBytes, limits.maxSaveChanges);
        const base = dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument" : "http://schemas.openxmlformats.org/officeDocument/2006";
        customInsertionSource = { part: planned.part, rootNamespace: `${base}/${dialect === "strict" ? "customProperties" : "custom-properties"}`, valueNamespace: `${base}/docPropsVTypes`, usedIds: new Set(), reservedNames: new Set(), create: { dialect, relationshipId: planned.relationshipId } };
      } catch (cause) {
        if (!(cause instanceof XmlSourceEditError) || !["xml-edit-unsupported", "xml-edit-conflict"].includes(cause.code)) throw cause;
      }
    }
    const customBindingBySource = new Map(customPropertyBindings?.map(binding => [binding.sourceObjectId, binding]) ?? []);
    const customSourceReport = customPropertySources?.map(source => {
      const binding = customBindingBySource.get(source.source.sourceObjectId ?? "");
      return binding === undefined ? source : { ...source, projectedIndex: binding.index, ...(customPropertySaveProfile !== "scalar-v1" ? { nameEditable: binding.baselineName !== undefined } : {}) };
    });
    const customPropertyInsertion = customPropertySaveProfile === "append-and-remove-v4" || customPropertySaveProfile === "create-and-append-v5" ? customInsertionSource ? { status: "available" as const, part: customInsertionSource.part } : { status: "unavailable" as const, ...(context.metadata.custom ? { part: context.metadata.custom.part } : {}), reason: "A safe custom property insertion source could not be established." } : undefined;
    const metadataSourceReport = (metadataLevel >= 13) ? coreMetadataSources?.map(source => source.status === "bound" ? { ...source, removable: metadataBindings?.some(binding => binding.field === source.field && binding.removable === true) ?? false } : source) : coreMetadataSources;
    if (slideVisibilitySources || metadataSourceReport || customSourceReport || customPropertyInsertion) projection.report = freezeReport({ ...projection.report, ...(slideVisibilitySources ? { slideVisibilitySources } : {}), ...(metadataSourceReport ? { coreMetadataSources: metadataSourceReport } : {}), ...(customSourceReport ? { customPropertySources: customSourceReport } : {}), ...(customPropertyInsertion ? { customPropertyInsertion } : {}) });
    let firstNumberBinding: import("./analyze-text-edits.js").FirstSlideNumberBinding | undefined;
    if (firstSlideNumberSaveProfile && (inspection.firstSlideNumberDeclared ? inspection.firstSlideNumber >= 0 : firstSlideNumberSaveProfile === "attribute-v2")) {
      const root = await context.root(inspection.mainPart, "presentation"), raw = await context.sourceElement(inspection.mainPart, root);
      const attr = raw.node.attributes.find(item => item.namespaceUri === "" && item.local === "firstSlideNum");
      if (Boolean(attr) === Boolean(inspection.firstSlideNumberDeclared) && raw.node.namespaceUri === context.p && raw.node.local === "presentation") firstNumberBinding = { part: inspection.mainPart, elementIndex: raw.elementIndex, namespace: context.p, ...(attr ? { expected: attr.value, baseline: inspection.firstSlideNumber } : {}), allowPresenceChange: firstSlideNumberSaveProfile === "attribute-v2" };
    }
    const editState = bindings || chartBindings || pictureEditSaveProfile || firstSlideNumberSaveProfile || descriptionBindings || slideVisibilityBindings || metadataBindings || customPropertyBindings ? {
      bindings: bindings ?? [], ...((metadataLevel >= 16) ? { allowCoreRemovals: true } : {}), ...(metadataLevel >= 17 ? { allowCoreStringRemovals: true } : {}), ...(customPropertyBindings ? { customPropertyBindings } : {}), ...(["remove-v3", "append-and-remove-v4", "create-and-append-v5"].includes(customPropertySaveProfile ?? "") ? { allowCustomRemovals: true } : {}), ...(["append-and-remove-v4", "create-and-append-v5"].includes(customPropertySaveProfile ?? "") ? { allowCustomInsertions: true } : {}), ...(customInsertionSource ? { customInsertionSource } : {}), ...(metadataBindings ? { metadataBindings } : {}), ...((metadataLevel >= 5) ? { allowCoreInsertions: true } : {}), ...(slideVisibilityBindings ? { slideVisibilityBindings } : {}), ...(chartBindings ? { chartBindings } : {}), ...(chartLabelBindings ? { chartLabelBindings } : {}), ...(chartOrderBindings ? { chartOrderBindings } : {}), ...(diagramPreview ? { diagramPreview } : {}), ...(textFontProfile ? { textFontProfile } : {}), ...(notesTextProfile ? { notesTextProfile } : {}), guards: await textEditGuards(context, inspection, limits.maxReportFindings, checkpoint, profile !== "text-runs-v1", textTargetSaveProfile ? bindings : undefined),
      ...(firstNumberBinding ? { firstNumberBinding } : {}), ...(descriptionBindings ? { descriptionBindings, descriptionBatch: drawingNameSaveProfile !== undefined || decorativeSaveProfile !== undefined || descriptionEditSaveProfile === 'batch-v3' || descriptionEditSaveProfile === 'mixed-v4' || descriptionEditSaveProfile === 'mixed-v5', descriptionMixed: descriptionEditSaveProfile === 'mixed-v4' || descriptionEditSaveProfile === 'mixed-v5' } : {}),
      ...(pictureEditSaveProfile ? { pictureSources: pictureSources! } : {}), ...(fieldTextProfile ? { fieldTextProfile } : {}),
      options: { ...(diagramTextComparison ? { diagramTextComparison } : {}), diagramLimits, limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits },
    } : undefined;
    abort();
    return createSession(projection.document, projection.report, source, limits, editState);
  } catch (cause) {
    if (cause instanceof PptxImportError && cause.code === "save-limit-exceeded") throw new PptxImportError("import-limit-exceeded", cause.message, cause.source, { cause });
    if (cause instanceof PackageInspectionError) {
      const code = cause.code === "operation-aborted" ? "operation-aborted" : cause.code.endsWith("limit-exceeded") ? "import-limit-exceeded" : cause.code.endsWith("unsupported") ? "import-unsupported" : "import-invalid-package";
      throw new PptxImportError(code, cause.message, cause.entry ? { part: cause.entry.startsWith("/") ? cause.entry : `/${cause.entry}` } : undefined, { cause });
    }
    throw cause;
  }
}
