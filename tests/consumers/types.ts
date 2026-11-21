import { inspectPptxMasterPictures, type MasterPictureInventory } from "@pptkit/pptx-parser";
import { renderLayoutToPdf, type PdfBrowser, type PdfRenderResult } from "@pptkit/pdf-renderer";
import { inspectStaticImage, validateJpegStructure, type StaticImageInspection } from "@pptkit/media-inspection";
const jpegPrecision: number = validateJpegStructure(new Uint8Array()).precision;
const staticImage: StaticImageInspection = inspectStaticImage(new Uint8Array());
void [jpegPrecision, staticImage];
declare const installedPdfBrowser: PdfBrowser;
const installedPdfResult: Promise<PdfRenderResult> = renderLayoutToPdf(resolveLayout(createPresentation()), { browser: installedPdfBrowser, fontPolicy: "require-embedded", imageAnimationPolicy: "require-static", isStaticImage: request => request.mimeType === "image/png", jpegDecodePolicy: "require-opaque", jpegPrecisionPolicy: "require-8-bit", jpegStructurePolicy: "require-valid-8-bit", maxStructureInspectionBytes: 1000, maxJpegInspectionPixels: 1000 });
const installedAutomaticPdf: Promise<PdfRenderResult> = renderLayoutToPdf(resolveLayout(createPresentation()), { browser: installedPdfBrowser, imageAnimationPolicy: "require-structurally-static", maxStaticInspectionBytes: 1000 });
void installedAutomaticPdf;
void installedPdfResult;
import { planVisibleImageReplacement, planNestedVisibleImageReplacement, planSharedVisibleImageReplacement, type SharedVisibleImageReplacementPlan, type NestedVisibleImageReplacementPlan, type NestedVisibleImageReplacementOptions, type SharedVisibleImageReplacementOptions, type VisibleImageReplacementFit, type VisibleImageReplacementOptions } from "@pptkit/layout";
import { planImageReplacement, evaluatePresentationField, fieldContextForOccurrences, fieldContextForSlide, materializeFieldText, formatListMarker, type PresentationField, type FieldEvaluationContext, type ImageReplacementPolicy, type TextBulletInput, type TextTabStop } from "@pptkit/core";
const installedTabStop: TextTabStop = { position: 72, align: "decimal" };
// @ts-expect-error Unsupported tab alignment.
const invalidTabStop: TextTabStop = { position: 72, align: "justify" };
void [installedTabStop, invalidTabStop];
const installedContinuation: TextBulletInput = { type: "number", startAt: "continue" };
// @ts-expect-error Only the explicit continuation keyword is accepted.
const invalidContinuation: TextBulletInput = { type: "number", startAt: "next" };
void [installedContinuation, invalidContinuation];
const installedListMarker: string = formatListMarker({ type: "number", style: "alphaUpperPeriod", startAt: 53 });
void installedListMarker;
const installedField: PresentationField = { id: "page", kind: "slide-number" };
const installedContext: FieldEvaluationContext = { pageOrder: ["page"], pageId: "page", firstPageNumber: 1 };
evaluatePresentationField(installedField, installedContext);
materializeFieldText([{ runs: [{ field: installedField }] }], installedContext);
materializeFieldText([{ runs: [{ field: { id: "38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a", kind: "slide-number" }, output: "native" }] }], installedContext);
materializeFieldText([{ runs: [{ field: { id: "38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a", kind: "dynamic-date", dateStyle: "short", updateFormat: "day-month-year" }, output: "native" }] }], { instantMs: 0, locale: "en-US", timeZone: "UTC" });
materializeFieldText([{ runs: [{ field: { id: "38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a", kind: "dynamic-date", timeStyle: "medium", updateFormat: "time-24-seconds" }, output: "native" }] }], { instantMs: 0, locale: "en-US", timeZone: "UTC" });
fieldContextForSlide(createPresentation(), { slideId: "page", firstPageNumber: 1, includeHidden: false });
fieldContextForOccurrences(resolveLayout(createPresentation()), { occurrences: [{ id: "visit", slideId: "page" }], occurrenceId: "visit", firstPageNumber: 1, includeHidden: false });
import { normalizeJpegOrientationInBrowser, orientJpegCrop, type BrowserJpegNormalizationResult } from "@pptkit/pptx-exporter/dom";
import { createPresentation, normalizePresentation, type PresentationDocument, type NormalizedPresentation, type PresentationDiagnostic } from "@pptkit/core";
import { validatePresentation } from "@pptkit/core/browser";
import { resolveLayout, type LayoutResult, type LayoutOptions, type TextMeasurer, type LayoutComposition } from "@pptkit/layout";
const fieldRefreshLayoutOptions: LayoutOptions = { fieldRefresh: { firstPageNumber: 1, includeHidden: false, instantMs: 0, locale: "en-US", timeZone: "UTC", dateStyle: "short" } };
void fieldRefreshLayoutOptions;
import { resolveNormalizedLayout } from "@pptkit/layout/browser";
import { generatePptx, inspectGif, inspectPng, inspectApng, inspectJpegExifOrientation, type JpegExifOrientationInspection, type GifInspection, type PngInspection, type ApngInspection, type GeneratePptxResult } from "@pptkit/pptx-exporter";
import { generatePptx as generateBrowserPptx } from "@pptkit/pptx-exporter/browser";
import { writePptx, type WritePptxResult } from "@pptkit/pptx-exporter/node";
import { renderLayoutToSvg, type SvgRenderResult } from "@pptkit/svg-renderer";
import { renderPresentationToSvg } from "@pptkit/svg-renderer/browser";
import { runCli, type CliRunResult } from "@pptkit/cli";
import type { PptxPlaceholderInheritance, PptxPlaceholderInfo, PptxTextStyleInheritance, PptxTextStylePropertyInfo, PptxSessionLimits, ImportReport, ImportedNotesTextRunSource, ImportedTableTextRunSource } from "@pptkit/pptx-parser";
import { importPptxFile, writePptxFile, type PptxFileSnapshot, type PptxFileWriteOptions, type PptxFileWriteResult } from "@pptkit/pptx-parser/node";
const fieldImportOptions: PptxImportOptions = { projection: "content", fieldTextProfile: "date-time-and-number-v1" };
const sparseScatterSaveOptions: PptxImportOptions = { projection: "content", chartEditSaveProfile: "existing-values-v2" };
void sparseScatterSaveOptions;
const descriptionImportOptions: PptxImportOptions = { projection: "content", descriptionEditSaveProfile: "existing-v1" };
void descriptionImportOptions;
const descriptionAttributeOptions: PptxImportOptions = { projection: "content", descriptionEditSaveProfile: "attribute-v2" };
void descriptionAttributeOptions;
const descriptionBatchOptions: PptxImportOptions = { projection: "content", descriptionEditSaveProfile: "batch-v3" };
void descriptionBatchOptions;
const descriptionMixedOptions: PptxImportOptions = { projection: "content", editSaveProfile: "text-runs-v3", descriptionEditSaveProfile: "mixed-v4" };
void descriptionMixedOptions;
const descriptionMixedChartOptions: PptxImportOptions = { projection: "content", editSaveProfile: "text-runs-v3", descriptionEditSaveProfile: "mixed-v5", chartEditSaveProfile: "existing-values-v1", firstSlideNumberSaveProfile: "attribute-v2" };
void descriptionMixedChartOptions;
const drawingNameOptions: PptxImportOptions = { projection: "content", drawingNameSaveProfile: "existing-v1" };
void drawingNameOptions;
const combinedNameOptions: PptxImportOptions = { projection: "content", editSaveProfile: "text-runs-v3", descriptionEditSaveProfile: "mixed-v5", drawingNameSaveProfile: "existing-v1", chartEditSaveProfile: "existing-values-v1", firstSlideNumberSaveProfile: "attribute-v2" };
void combinedNameOptions;
const decorativeOptions: PptxImportOptions = { projection: "content", editSaveProfile: "text-runs-v3", descriptionEditSaveProfile: "mixed-v5", drawingNameSaveProfile: "existing-v1", decorativeSaveProfile: "existing-v1" };
void decorativeOptions;
const visibilityOptions: PptxImportOptions = { projection: "content", slideVisibilitySaveProfile: "attribute-v1" };
void visibilityOptions;
const metadataOptions: PptxImportOptions = { projection: "content", metadataSaveProfile: "core-strings-v1" };
void metadataOptions;
const metadataSourceType = (report: ImportReport): import("@pptkit/pptx-parser").ImportedCoreMetadataSource | undefined => report.coreMetadataSources?.[0];
void metadataSourceType;
createPresentation().setMetadata({ title: "Title", revision: 7 });
const visibilitySourceType = (report: ImportReport): import("@pptkit/pptx-parser").ImportedSlideVisibilitySource | undefined => report.slideVisibilitySources?.[0];
void visibilitySourceType;
createPresentation().addSlide().setHidden(true);
void fieldImportOptions;

function placeholderProvenance(report: ImportReport): PptxPlaceholderInfo | undefined {
  const inheritance: PptxPlaceholderInheritance | undefined = report.placeholderInheritance;
  const entry = inheritance?.placeholders[0];
  if (entry?.dimensions.width) {
    const emu: number = entry.dimensions.width.valueEmu;
    // @ts-expect-error Source provenance is an immutable report snapshot.
    entry.dimensions.width.valueEmu = emu;
  }
  return entry;
}
void placeholderProvenance;
function textStyleProvenance(report: ImportReport): PptxTextStylePropertyInfo | undefined {
  const inheritance: PptxTextStyleInheritance | undefined = report.textStyleInheritance;
  const first = inheritance?.properties[0];
  if (first) {
    const property: string = first.property;
    const level: string = first.level;
    const dependencyPart: string | undefined = first.resolvedFrom?.[0]?.part;
    const limits: Partial<PptxSessionLimits> = { maxTextStyleProperties: 500_000, maxTextStyleDependencies: 1_000_000 };
    const limit: number | undefined = limits.maxTextStyleProperties;
    void [property, level, limit, dependencyPart];
    // @ts-expect-error Source provenance is an immutable report snapshot.
    first.status = "unsupported";
  }
  return first;
}
void textStyleProvenance;

const document: PresentationDocument = createPresentation();
document.addSlide().addElement({ type: "text", content: "Typed consumer", box: { x: 0, y: 0, width: 200, height: 50 } });
document.addSlide().addElement({ type: "text", content: [{ runs: [{ text: "Typed alpha", style: { colorOpacity: .25 } }] }], box: { x: 0, y: 0, width: 200, height: 50 } });
const normalized: NormalizedPresentation = normalizePresentation(document);
const diagnostics: PresentationDiagnostic[] = validatePresentation(document);
const layouts: LayoutResult[] = [resolveLayout(document), resolveNormalizedLayout(normalized)];
const generated: GeneratePptxResult[] = await Promise.all([generatePptx(document), generateBrowserPptx(document)]);
const written: WritePptxResult = await writePptx(document, { output: "typed.pptx" });
const previews: SvgRenderResult[] = await Promise.all([renderLayoutToSvg(layouts[0]!), renderPresentationToSvg(document)]);
const cli: CliRunResult = runCli(["--help"]);
void [diagnostics, generated, written, previews, cli];


const measure: TextMeasurer = (request) => {
  // @ts-expect-error Measurement requests are deeply readonly.
  request.paragraphs[0]!.runs[0]!.style.fontSize = 12;
  return { bounds: { x: 0, y: 0, width: request.availableWidth, height: 20 } };
};
// @ts-expect-error Load fonts before resolving Layout; asynchronous providers are unsupported.
const asynchronous: TextMeasurer = async () => ({ bounds: { x: 0, y: 0, width: 1, height: 1 } });
const options: LayoutOptions = { textMeasurement: { measure, resolveFont: ({ family }) => ({ family }) }, diagnostics: { textOverflow: true, slideBounds: true } };
const analyzed = resolveNormalizedLayout(normalized, options);
void analyzed.analysis?.measurements[0]?.contentBox;
void asynchronous;
await generateBrowserPptx(document, { layout: options });
await generatePptx(document, { layout: options });
await writePptx(document, { output: "analyzed.pptx", layout: options });
await renderPresentationToSvg(document, { layout: options, idNamespace: "consumer-instance" });
await renderPresentationToSvg(document, { pngAnimationDiagnostics: "warn" });


import { createBrowserTextMeasurement, type BrowserMeasurementFont } from "@pptkit/layout/dom";
// Typecheck the browser adapter without invoking DOM APIs in Node consumers.
function browserAnalysis(browserDocument: Document, fonts: readonly BrowserMeasurementFont[]) {
  const adapter = createBrowserTextMeasurement({ document: browserDocument, fonts });
  try { return resolveLayout(document, { textMeasurement: adapter }); }
  finally { adapter.dispose(); }
}
void browserAnalysis;

import { mountBrowserMedia, resolveAndMountBrowserMedia, mountLayoutMedia, mountLayoutMediaOverlay, mountLayoutMediaInPlace, createSharedBrowserMediaPool, type BrowserLayoutMedia, type BrowserMediaState, type BrowserMediaResolveState } from "@pptkit/svg-renderer/media";
function browserMedia(host: HTMLElement, source: Blob) {
  const pool = createSharedBrowserMediaPool(window, { maxBytes: 1000, resolve: (_key, { signal }) => { signal.throwIfAborted(); return source; } });
  void pool.mount(host, "movie", { kind: "video", contentType: "video/mp4" });
  const controller = mountBrowserMedia(host, { kind: "video", source, contentType: "video/mp4", muted: true });
  const observe = (state: BrowserMediaState) => state.status;
  const unsubscribe = controller.subscribe(observe);
  unsubscribe(); controller.dispose();
  return resolveAndMountBrowserMedia(host, { kind: "video", contentType: "video/mp4", maxBytes: 1000, onResolveState: (state: BrowserMediaResolveState) => state.status, resolve: async ({ signal }) => { signal.throwIfAborted(); return source; } });
}
void browserMedia;
function layoutMedia(host: HTMLElement, layout: ReturnType<typeof resolveLayout>) {
  const result: Promise<BrowserLayoutMedia> = mountLayoutMedia(host, layout, { slideId: 'page', elementId: 'video', maxBytes: 1000 });
  return result;
}
void layoutMedia;
function layoutMediaOverlay(svg: SVGSVGElement, layout: ReturnType<typeof resolveLayout>) {
  const result: Promise<BrowserLayoutMedia> = mountLayoutMediaOverlay(svg, layout, { slideId: 'page', elementId: 'video', maxBytes: 1000 });
  void mountLayoutMediaInPlace(svg, layout, { slideId: 'page', elementId: 'video', maxBytes: 1000 });
  return result;
}
void layoutMediaOverlay;



const composition: LayoutComposition[] = [{
  type: "stack", direction: "horizontal", container: { kind: "slide", id: document.slides[0]!.id },
  elementIds: [document.slides[0]!.elements[0]!.id], padding: 10, gap: 5, align: "center",
}];
resolveLayout(document, { composition });
await generatePptx(document, { layout: { composition } });
await renderPresentationToSvg(document, { layout: { composition } });


resolveLayout(document, {
  textMeasurement: { measure },
  tableSizing: [{ elementId: "table", autoRows: [{ row: 0, minHeight: 10, maxHeight: 100 }], overflow: "preserve" }],
});


import type { IntrinsicSizingRule } from "@pptkit/layout";
const intrinsicSizing: readonly IntrinsicSizingRule[] = [{ elementId: "text", width: { type: "parent", fraction: 0.5 }, height: { type: "content", min: 20, max: 200 } }];
resolveLayout(document, { textMeasurement: { measure }, intrinsicSizing });
await generatePptx(document, { layout: { textMeasurement: { measure }, intrinsicSizing } });
// @ts-expect-error Automatic dimensions require an explicit supported source.
const invalidIntrinsic: IntrinsicSizingRule = { elementId: "text", width: { type: "auto" } };
void invalidIntrinsic;

import type { GridComposition, GridTrack } from "@pptkit/layout";
const track: GridTrack = { type: "fraction", weight: 2, max: 300 };
const grid: GridComposition = { type: "grid", container: { kind: "slide", id: document.slides[0]!.id }, columns: [track], rows: [{ type: "content" }], placements: [{ elementId: document.slides[0]!.elements[0]!.id, row: 0, column: 0 }] };
resolveLayout(document, { composition: [grid] });
// @ts-expect-error Content tracks do not accept fraction weights.
const invalidTrack: GridTrack = { type: "content", weight: 2 };
void invalidTrack;

import type { ConstraintsComposition, ConstraintExpression } from "@pptkit/layout";
const expression: ConstraintExpression = { terms: [{ source: "container", metric: "width", factor: 0.5 }], constant: -10 };
const constraints: ConstraintsComposition = { type: "constraints", container: { kind: "slide", id: document.slides[0]!.id }, elements: [{ elementId: document.slides[0]!.elements[0]!.id, width: { value: expression, min: 0 } }] };
resolveLayout(document, { composition: [constraints] });
await generatePptx(document, { layout: { composition: [constraints] } });
// @ts-expect-error Constraint geometry must use unrotated frames.
const invalidConstraints: ConstraintsComposition = { ...constraints, bounds: "visual" };
void invalidConstraints;

import type { TablePaginationRule, TablePaginationFragment } from "@pptkit/layout";
const tablePagination: readonly TablePaginationRule[] = [{ elementId: "table", area: { x: 10, y: 10, width: 200, height: 150 }, headerRows: 1, continuation: { area: { x: 10, y: 20, width: 200, height: 200 } }, nativeRowRegions: "source-boundaries", maxPages: 100 }];
const paginated = resolveLayout(document, { tablePagination });
const fragments: readonly TablePaginationFragment[] = paginated.pagination?.fragments ?? [];
await generatePptx(document, { layout: { tablePagination } });
await renderPresentationToSvg(document, { layout: { tablePagination } });
void fragments;
// @ts-expect-error Pagination does not silently clip or split an indivisible row.
const invalidPagination: TablePaginationRule = { ...tablePagination[0]!, overflow: "clip" };
void invalidPagination;

import { inspectZipArchive, type ZipArchiveInspection } from "@pptkit/pptx-parser";
const archive: ZipArchiveInspection = await inspectZipArchive(generated[0]!.bytes, { limits: { maxEntries: 2000 } });
// @ts-expect-error Inventory entries are readonly at the public contract.
archive.entries.push({});
void archive;

import { inspectXmlPart, type XmlPartInspection } from "@pptkit/pptx-parser";
const xml: XmlPartInspection = await inspectXmlPart(new TextEncoder().encode("<r/>"), { partName: "ppt/presentation.xml", limits: { maxDepth: 64 } });
// @ts-expect-error XML names are readonly inventory, not mutable authoring state.
xml.root.local = "changed";
// @ts-expect-error Namespace inventory is readonly.
xml.namespaceUris.push("urn:changed");
void xml;

import { inspectOpcPackage, type OpcPackageInspection } from "@pptkit/pptx-parser";
const opc: OpcPackageInspection = await inspectOpcPackage(generated[0]!.bytes, { limits: { maxRelationships: 1000 }, xmlLimits: { maxDepth: 64 } });
// @ts-expect-error OPC relationships are readonly inventory.
opc.relationships.push({});
// @ts-expect-error Aggregate counts are readonly.
opc.xmlTotals.nodes = 0;
// @ts-expect-error Content-type mappings are readonly.
opc.contentTypes.defaults[0]!.extension = "changed";
void opc;

import { inspectPresentationPackage, type PresentationPackageInspection } from "@pptkit/pptx-parser";
const presentationInspection: PresentationPackageInspection = await inspectPresentationPackage(generated[0]!.bytes);
const dialect: "strict" | "transitional" = presentationInspection.dialect;
// @ts-expect-error inspection metadata is readonly
presentationInspection.mainPart = "/changed.xml";
// @ts-expect-error result has no semantic slide collection
presentationInspection.slides;
void dialect;

import { inspectMarkupCompatibility, type MarkupCompatibilityInspection } from "@pptkit/pptx-parser";
const compatibility: MarkupCompatibilityInspection = await inspectMarkupCompatibility(new Uint8Array(), { understoodNamespaces: ["urn:test"], extensionElements: [{ namespaceUri: "urn:test", local: "ext" }] });
// @ts-expect-error namespace capabilities are mandatory
inspectMarkupCompatibility(new Uint8Array(), {});
// @ts-expect-error compatibility decisions are readonly
compatibility.decisions.push({});
// @ts-expect-error effective root list is readonly
compatibility.output.roots.push({});
void compatibility;

import { inspectPptx, type PptxInspection, type PptxDocumentMetadata, type MetadataPropertyValue } from "@pptkit/pptx-parser";
const structural: PptxInspection = await inspectPptx(generated[0]!.bytes);
import type { PptxDiagramDataModelInfo, PptxInspectionOptions, PptxDiagramPresentationMapping, PptxDiagramCacheMapping } from "@pptkit/pptx-parser";
const graphOptions: PptxInspectionOptions = { diagramLimits: { maxWorkUnits: 10000, maxFindings: 100 } };
const sourceGraph: PptxDiagramDataModelInfo | undefined = (await inspectPptx(generated[0]!.bytes, graphOptions)).diagramInventory?.dataModels?.[0];
const presentationMapping: PptxDiagramPresentationMapping | undefined = sourceGraph?.presentationMappings?.[0];
const sourceCacheMapping: PptxDiagramCacheMapping | undefined = structural.diagramInventory?.cacheMappings?.[0];
if (sourceCacheMapping) {
  // @ts-expect-error Cache mappings are readonly source evidence.
  sourceCacheMapping.shapes.push({});
}
if (presentationMapping) {
  // @ts-expect-error Resolved references cannot be rebound by editing an inspection.
  presentationMapping.sources[0]!.reference.targetElementIndex = 1;
}
if (sourceGraph) {
  // @ts-expect-error Source graphs do not expose mutable authoring collections.
  sourceGraph.points.push(sourceGraph.points[0]!);
  // @ts-expect-error Source model identities are readonly.
  sourceGraph.points[0]!.modelId = "1";
  // @ts-expect-error This inventory is not a native diagram editor.
  sourceGraph.addNode({});
}
import type { ImportedDiagramPreview } from "@pptkit/pptx-parser";
const cachedSession = await importPptx(generated[0]!.bytes, { projection: "content", diagramPreview: "cache", sessionLimits: { maxDiagramPreviewWork: 100000 } });
const cachePreview: ImportedDiagramPreview | undefined = cachedSession.report.diagramPreviews?.[0];
if (cachePreview) {
  // @ts-expect-error Cached preview is not a semantic SmartArt node.
  cachePreview.nodes;
  // @ts-expect-error Freshness is an immutable unknown source property.
  cachePreview.freshness = "unknown";
}
import type { PptxDiagramInfo, PptxDiagramComponentInfo, PptxDiagramComponentRole, PptxDiagramResourceInventory, PptxDiagramCacheInfo } from "@pptkit/pptx-parser";
const diagram: PptxDiagramInfo | undefined = structural.slides[0]?.diagrams?.[0];
if (diagram) {
  const component: PptxDiagramComponentInfo = diagram.components[0]!;
  const role: PptxDiagramComponentRole = component.role;
  // @ts-expect-error Source identities are readonly inspection data.
  diagram.elementIndex = 0;
  // @ts-expect-error Source components do not grant authoring mutations.
  diagram.components.push(component);
  const cache: PptxDiagramCacheInfo | undefined = diagram.cache;
  if (cache) {
    // @ts-expect-error Cache freshness is an inspection result, not authoring state.
    cache.freshness = "current";
  }
  void role;
}
const diagramResources: PptxDiagramResourceInventory | undefined = structural.diagramInventory;
if (diagramResources) {
  // @ts-expect-error Resource owners are readonly source identities.
  diagramResources.parts[0]!.owners.push({ slidePart: "/fake", elementIndex: 1, role: "data" });
}
const sourceSlideId: number = structural.slides[0]!.id;
// @ts-expect-error inspected source slide order is readonly
structural.slides.reverse();
// @ts-expect-error structural inspection is not an editable document
structural.document.addSlide();
void sourceSlideId;
const documentMetadata: PptxDocumentMetadata = structural.metadata;
// @ts-expect-error metadata collections are readonly
documentMetadata.core!.properties.push({ name: "title", value: "changed" });
// @ts-expect-error metadata values are readonly
documentMetadata.core!.properties[0]!.value = "changed";
function propertyText(value: MetadataPropertyValue): string | undefined {
  if (value.status === "scalar") return value.value;
  // @ts-expect-error uninterpreted values expose no scalar text or raw XML
  return value.value;
}
void propertyText;
import type { PptxSectionListInfo, PptxExtensionInfo } from "@pptkit/pptx-parser";
const sourceSections: PptxSectionListInfo | undefined = structural.sectionList;
const sourceExtensions: readonly PptxExtensionInfo[] = structural.presentationExtensions;
if (sourceSections) {
  // @ts-expect-error native section members are readonly
  sourceSections.sections[0]!.slides.reverse();
  // @ts-expect-error source section GUID is readonly
  sourceSections.sections[0]!.id = "changed";
}
void sourceExtensions;
import { importPptx, type ImportedPresentation, type PptxSavePlan } from "@pptkit/pptx-parser";

async function checkDefaultMetadataRemovalTypes(session: ImportedPresentation): Promise<void> {
  await session.prepareSave({ mode: "edit", coreMetadataRemovals: ["title", "author", "language", "revision"] });
  // @ts-expect-error only explicitly supported defaulted fields may be selected
  await session.prepareSave({ mode: "edit", coreMetadataRemovals: ["company"] });
}
void checkDefaultMetadataRemovalTypes;
import type { PptxSaveChange, ImportedTextRunSource } from "@pptkit/pptx-parser";
import { compareOpcPackages, type OpcPackageComparison, type OpcComparisonOptions } from "@pptkit/pptx-parser";
import { inspectOpcClosure, type OpcClosureInspection, type OpcClosureOptions } from "@pptkit/pptx-parser";
const closureOptions: OpcClosureOptions = { roots: ["/"], closureLimits: { maxSourceRelationships: 10000, maxClosureExpandedBytes: 64000000 } };
const dependencyClosure: OpcClosureInspection = await inspectOpcClosure(generated[0]!.bytes, closureOptions);
void dependencyClosure.coverage.outgoingComplete;
// @ts-expect-error Roots are required; no implicit package-root query.
const invalidClosureOptions: OpcClosureOptions = {};
void invalidClosureOptions;
// @ts-expect-error Dependency inventories are immutable.
dependencyClosure.parts[0]!.name = "/changed.xml";
const comparisonOptions: OpcComparisonOptions = { profile: "opc-xml-v1", comparisonLimits: { maxDifferences: 20, maxWorkUnits: 1000000 }, zipLimits: { maxEntries: 1000 } };
const packageComparison: OpcPackageComparison = await compareOpcPackages(generated[0]!.bytes, generated[0]!.bytes, comparisonOptions);
void packageComparison.equal;
// @ts-expect-error Comparison results expose readonly findings.
packageComparison.differences.push({ code: "payload-changed", part: "/example.xml" });
// @ts-expect-error Profile selection is explicit and versioned.
const invalidComparisonOptions: OpcComparisonOptions = { profile: "lossless" };
void invalidComparisonOptions;
const importedSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "structure" });
const contentSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", sessionLimits: { maxImportedAssetBytes: 1024, maxTotalImportedAssetBytes: 4096 } });
const pictureSaveSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", pictureEditSaveProfile: "replace-v1" });
pictureSaveSession.report.pictureSources?.forEach(source => { const part: string = source.mediaPart; void part; });
pictureSaveSession.dispose();
const sharedPictureSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", masterArtwork: "pictures-v1" });
sharedPictureSession.dispose();
const sharedArtworkSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", masterArtwork: "artwork-v1" });
sharedArtworkSession.dispose();
const sharedArtworkTextSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", masterArtwork: "artwork-v2" });
sharedArtworkTextSession.dispose();
const themedArtworkSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", masterArtwork: "artwork-v3" });
themedArtworkSession.dispose();
const themedArtworkPaintSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", masterArtwork: "artwork-v4" });
themedArtworkPaintSession.dispose();
const backgroundArtworkSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", masterArtwork: "artwork-v5" });
backgroundArtworkSession.dispose();
const matrixBackgroundSession: ImportedPresentation = await importPptx(generated[0]!.bytes, { projection: "content", masterArtwork: "artwork-v6" });
matrixBackgroundSession.dispose();
contentSession.document.slides[0]!.setNotes([{ runs: [{ text: "Edited speaker notes", action: { type: "slide", slideId: contentSession.document.slides[0]!.id } }] }]);
contentSession.document.slides[0]!.setNotes();
const projectionMode: "structure" | "content" = contentSession.report.projection;
const editingSession = await importPptx(generated[0]!.bytes, { projection: "content", editSaveProfile: "text-runs-v1", sessionLimits: { maxSaveChanges: 100, maxSaveAnalysisWork: 100000, maxSaveValidationBytes: 64000000, maxSaveValidationWork: 10000000 } });
const importedRun: ImportedTextRunSource | undefined = editingSession.report.textRuns?.[0];
if (importedRun) {
  const rawIndex: number | undefined = importedRun.source.elementIndex;
  void rawIndex;
  // @ts-expect-error Source run provenance is immutable.
  importedRun.sourceElementIndex = 1;
}
const editedPreparation = await editingSession.prepareSave({ mode: "edit" });
const ownerSession = await importPptx(generated[0]!.bytes, { projection: "content", editSaveProfile: "text-runs-v2" });
const notesRun: ImportedNotesTextRunSource | undefined = ownerSession.report.notesTextRuns?.[0];
if (notesRun) {
  const slideId: string = notesRun.slideId;
  void slideId;
  // @ts-expect-error Notes provenance has no fabricated drawing element ID.
  void notesRun.elementId;
  // @ts-expect-error Notes source identities are immutable.
  notesRun.source.part = "/changed.xml";
  // @ts-expect-error Notes run indices are immutable.
  notesRun.runIndex = 1;
}
const drawingOwner: "text-box" | "shape-text" | undefined = ownerSession.report.textRuns?.[0]?.owner;
void drawingOwner;
const tableSession = await importPptx(generated[0]!.bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
const tableRun: ImportedTableTextRunSource | undefined = tableSession.report.tableTextRuns?.[0];
if (tableRun) {
  const coordinates: readonly number[] = [tableRun.rowIndex, tableRun.cellIndex, tableRun.columnIndex, tableRun.paragraphIndex, tableRun.runIndex];
  const elementId: string = tableRun.elementId;
  void [coordinates, elementId];
  // @ts-expect-error Table source coordinates are immutable.
  tableRun.columnIndex = 1;
  // @ts-expect-error Source identity is immutable.
  tableRun.source.elementIndex = 1;
}
if (editedPreparation.status === "ready") {
  const changes: readonly PptxSaveChange[] | undefined = editedPreparation.plan.report.changes;
  if (changes?.[0]?.text) {
    // @ts-expect-error A prepared report cannot be changed by callers.
    changes[0].text.after = "tampered";
  }
}
// @ts-expect-error Edit-save profiles are explicitly versioned.
importPptx(generated[0]!.bytes, { projection: "content", editSaveProfile: "all-edits" });
void projectionMode;
// @ts-expect-error structural projection requires explicit caller opt-in
importPptx(generated[0]!.bytes);
// @ts-expect-error session document reference cannot be replaced
importedSession.document = importedSession.document;
// @ts-expect-error report capability findings are readonly
importedSession.report.findings[0]!.tier = "editable";
const savePreparation = await importedSession.prepareSave({ mode: "no-op" });
if (savePreparation.status === "ready") {
  const savePlan: PptxSavePlan = savePreparation.plan;
  importedSession.assertSavePlanCurrent(savePlan, { acceptLossIds: [] });
  const importedFile = await importPptxFile("/example.pptx", { projection: "content", editSaveProfile: "text-runs-v3" });
  const fileSnapshot: PptxFileSnapshot = importedFile.file;
  const fileOptions: PptxFileWriteOptions = { mode: "replace", expected: fileSnapshot };
  const fileOutput: PptxFileWriteResult = await writePptxFile(importedSession, savePlan, "/example.pptx", fileOptions);
  const status: "committed" = fileOutput.status;
  void status;
  // @ts-expect-error File snapshots are nominal; a digest/path object cannot authorize overwrite.
  const forgedFile: PptxFileSnapshot = { path: "/example.pptx", byteLength: 1, sha256: "fake" };
  // @ts-expect-error Replace requires an explicit authentic file snapshot.
  const unguardedReplace: PptxFileWriteOptions = { mode: "replace" };
  // @ts-expect-error Create mode cannot carry overwrite authority.
  const overwriteCreate: PptxFileWriteOptions = { mode: "create", expected: fileSnapshot };
  // @ts-expect-error Snapshots and write warnings are immutable.
  fileOutput.file.sha256 = "changed";
  void [forgedFile, unguardedReplace, overwriteCreate];
  // @ts-expect-error caller cannot manufacture a branded save plan
  const forgedPlan: PptxSavePlan = { byteLength: 1, report: savePlan.report, dispose() {} };
  // @ts-expect-error no generic allowLossy consent override exists
  importedSession.save(savePlan, { allowLossy: true });
  void forgedPlan;
}

import type { PresentationAssetSourceInput, PresentationAssetSource } from "@pptkit/core";
import type { PptxAssetResolver } from "@pptkit/pptx-exporter";
const byteInput: PresentationAssetSourceInput = { type: "bytes", value: new Uint8Array([1]) };
const encodedInput: PresentationAssetSource = { type: "bytes", encoding: "base64", value: "AQ==" };
// @ts-expect-error stored source cannot contain a mutable typed array
const mutableStoredSource: PresentationAssetSource = byteInput;
const byteResolver: PptxAssetResolver = (asset, context) => {
  // @ts-expect-error resolver source is readonly
  asset.source.value = "changed";
  // @ts-expect-error resolver accessibility is readonly
  asset.accessibility.decorative = true;
  context.signal?.throwIfAborted();
  return { bytes: new Uint8Array([1]), mimeType: "image/png" };
};
void encodedInput; void mutableStoredSource; void byteResolver;

import type { AssetReference, AssetRetentionOptions, AssetRegistrationOptions, PresentationAssetReplacement } from "@pptkit/core";
const registrationOptions: AssetRegistrationOptions = { deduplicate: "content" };
const retainedOptions: AssetRetentionOptions = { retainAssetIds: ["kept"] };
const replacement: PresentationAssetReplacement = { kind: "image", source: byteInput };
const lifecycleAsset = document.registerAsset({ ...replacement, id: "kept" }, registrationOptions);
document.replaceAsset(lifecycleAsset.id, replacement);
const references: readonly AssetReference[] = document.getAssetReferences(lifecycleAsset.id);
// @ts-expect-error reference arrays are readonly
references.push({ assetId: "a", elementId: "e", path: "" });
// @ts-expect-error reference records are readonly
references[0]!.assetId = "changed";
// @ts-expect-error replacement cannot specify a new identity
document.replaceAsset("kept", { ...replacement, id: "other" });
// @ts-expect-error unsupported automatic destructive mode
const unsafeCleanup: AssetRetentionOptions = { force: true };
const removedAssets = document.removeUnusedAssets(retainedOptions);
// @ts-expect-error removed snapshot is readonly
removedAssets.pop();
void unsafeCleanup;

await generatePptx(document, { connectorMode: "native" });
await generatePptx(document, { connectorMode: "path" });
// @ts-expect-error export mode is an explicit union, not a generic boolean override
await generatePptx(document, { connectorMode: true });
import type { LayoutConnectorElement } from "@pptkit/layout";
function connectionIdentity(connector: LayoutConnectorElement): string | undefined {
  return connector.connections?.start?.elementId;
}
void connectionIdentity;

import type { SlideTransitionInput } from "@pptkit/core";
const fadeTransition: SlideTransitionInput = { type: "fade", durationMs: 1234, afterMs: 0, onClick: false };
document.slides[0]!.setTransition(fadeTransition);
// @ts-expect-error Unknown effects are not silently mapped to fade.
const unsupportedTransition: SlideTransitionInput = { type: "morph", durationMs: 1000 };
void unsupportedTransition;

const pushTransition: SlideTransitionInput = { type: "push", direction: "left", durationMs: 765 };
document.slides[0]!.setTransition(pushTransition);
// @ts-expect-error Side effects require an explicit direction.
const missingDirection: SlideTransitionInput = { type: "wipe", durationMs: 765 };
void missingDirection;

import type { PptxDiagramTextComparison } from "@pptkit/pptx-parser";
function diagramTextContract(value: PptxDiagramTextComparison) {
  const freshness: "unknown" = value.freshness;
  // @ts-expect-error Comparison snapshots are readonly.
  value.shapes.push(value.shapes[0]!);
  // @ts-expect-error Paragraph vectors are readonly.
  value.shapes[0]!.sourceParagraphs?.push("edited");
  // @ts-expect-error Equality cannot imply freshness.
  value.freshness = "current";
  return freshness;
}
void diagramTextContract;
await inspectPptx(generated[0]!.bytes, { diagramTextComparison: "direct-paragraphs" });
// @ts-expect-error Comparison requires an explicit supported profile.
await inspectPptx(generated[0]!.bytes, { diagramTextComparison: true });

import { runCliAsync, CLI_REPORT_SCHEMA_VERSION, type CliExecutionOptions } from "@pptkit/cli";
const cliOptions: CliExecutionOptions = { cwd: "/local", signal: new AbortController().signal };
// @ts-expect-error Execution options are readonly.
cliOptions.cwd = "/other";
const asyncCliResult: CliRunResult = await runCliAsync(["inspect", "example.pptx", "--json"], cliOptions);
const cliSchema: "pptkit-cli-report-v1" = CLI_REPORT_SCHEMA_VERSION;
void [asyncCliResult, cliSchema];

import { describeImportCapabilities, type ImportCapabilityReport } from "@pptkit/pptx-parser";
function capabilityContract(report: ImportReport): ImportCapabilityReport {
  const capabilities = describeImportCapabilities(report, { limits: { maxSubjects: 10 }, filter: { part: "/ppt/slides/slide1.xml" } });
  // @ts-expect-error Operation observations are immutable.
  capabilities.subjects[0]!.operations[0]!.observations.push({ id: "forged", tier: "editable", reason: "forged" });
  // @ts-expect-error Native validation was not performed by this pure reporter.
  capabilities.nativeVerification = "passed";
  // @ts-expect-error Task references are readonly.
  capabilities.subjects[0]!.references?.tasks.push("forged");
  return capabilities;
}
void capabilityContract;

import type { ImportedChartValueSource } from "@pptkit/pptx-parser";
function chartValueContract(report: ImportReport): ImportedChartValueSource | undefined {
  const source = report.chartValues?.[0];
  if (source?.workbook) {
    // @ts-expect-error Numeric source provenance is readonly.
    source.workbook.cellReference = "A1";
  }
  return source;
}
void chartValueContract;
await importPptx(generated[0]!.bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
// @ts-expect-error An explicit profile, not a generic bool, is required.
await importPptx(generated[0]!.bytes, { projection: "content", chartEditSaveProfile: true });

import { compileFiniteTimeline, sampleFiniteTimeline, FiniteTimelineError, type FiniteTimelineNode, type FiniteTimelinePlan, type FiniteTimelineLimits, type FiniteTimelineSample, type FiniteTimelineErrorCode } from "@pptkit/core";
import { sampleFiniteTimeline as sampleBrowserTimeline } from "@pptkit/core/browser";
const finiteInput: FiniteTimelineNode = { kind: "parallel", id: "root", children: [{ kind: "span", id: "a", durationMs: 10, repeat: 2, fill: "hold" }] };
const finiteLimits: FiniteTimelineLimits = { maxNodes: 20, maxDepth: 10, maxRepeat: 100, maxTimeMs: 1000, maxIdLength: 128, maxTotalIdLength: 1000 };
const finitePlan: FiniteTimelinePlan = compileFiniteTimeline(finiteInput, finiteLimits);
const finiteSamples: readonly FiniteTimelineSample[] = sampleFiniteTimeline(finitePlan, 10, finiteLimits);
sampleBrowserTimeline(finitePlan, 10, finiteLimits);
// @ts-expect-error Compiled plan rows are readonly snapshots.
finitePlan.leaves[0]!.endMs = 0;
// @ts-expect-error Sample collections are readonly.
finiteSamples.push({ id: "x", phase: "before", progress: null, iteration: null });
// @ts-expect-error Active samples require numeric progress and iteration.
const impossibleSample: FiniteTimelineSample = { id: "x", phase: "active", progress: null, iteration: null };
// @ts-expect-error A finite span does not authorize drawing/effect targets.
compileFiniteTimeline({ kind: "span", id: "x", durationMs: 10, target: { elementId: "x" } });
// @ts-expect-error Unknown node kinds are not mapped to finite spans.
compileFiniteTimeline({ kind: "effect", id: "x", durationMs: 10 });
// @ts-expect-error Input child arrays are readonly semantic snapshots.
finiteInput.children.push({ kind: "span", id: "b", durationMs: 10 });
const timelineFailure = new FiniteTimelineError("invalid-timeline-plan", "Example", "/leaves/0");
const finiteCode: FiniteTimelineErrorCode = timelineFailure.code;
void [impossibleSample, finiteCode];

import { compileEventTimeline, replayEventTimeline, EventTimelineError, type EventTimelineInput, type EventTimelinePlan, type EventTimelineReplayRequest, type EventTimelineReplay, type EventTimelineObservation, type EventTimelineErrorCode } from "@pptkit/core";
import { replayEventTimeline as browserReplayEvents } from "@pptkit/core/browser";
const eventInput: EventTimelineInput = { pageId: "p", elementIds: ["s", "b"], stages: [{ id: "gate", target: { pageId: "p", elementId: "s" }, trigger: { type: "object-click", elementId: "b" }, timing: { kind: "span", id: "leaf", durationMs: 20 } }] };
const eventPlan: EventTimelinePlan = compileEventTimeline(eventInput, { maxEvents: 10, maxStages: 2, maxElements: 4 });
const eventRequest: EventTimelineReplayRequest = { visitId: "v", timeMs: 10, elementIds: ["s", "b"], events: [{ id: "click", pageId: "p", visitId: "v", timeMs: 0, type: "object-click", elementId: "b" }] };
const eventResult: EventTimelineReplay = replayEventTimeline(eventPlan, eventRequest);
browserReplayEvents(eventPlan, eventRequest);
// @ts-expect-error Current inventory is mandatory for stale target detection.
replayEventTimeline(eventPlan, { visitId: "v", timeMs: 0, events: [] });
// @ts-expect-error Object-click events must identify a source object.
const missingEventSource: EventTimelineReplayRequest = { visitId: "v", timeMs: 0, elementIds: [], events: [{ id: "e", pageId: "p", visitId: "v", timeMs: 0, type: "object-click" }] };
// @ts-expect-error Stage effects do not admit subtarget indices.
compileEventTimeline({ ...eventInput, stages: [{ ...eventInput.stages[0]!, target: { pageId: "p", elementId: "s", paragraphIndex: 0 } }] });
// @ts-expect-error Observations do not mutate or dispatch the input log.
eventResult.events.push({ index: 0, eventId: "e", pageId: "p", visitId: "v", timeMs: 0, status: "future" });
// @ts-expect-error Whole-object targets are readonly snapshots.
eventPlan.stages[0]!.target.elementId = "deleted";
function consumedStage(observation: EventTimelineObservation): string | null {
  return observation.status === "consumed" ? observation.stageId : null;
}
const eventFailureCode: EventTimelineErrorCode = new EventTimelineError("invalid-event-timeline-plan", "Example", "/plan").code;
void [missingEventSource, consumedStage, eventFailureCode];

import type { SlideEventTimelineInput } from "@pptkit/core";
import { PptxExportError } from "@pptkit/pptx-exporter";
const slideTiming: SlideEventTimelineInput = { stages: [{ id: "s", target: { pageId: document.slides[0]!.id, elementId: document.slides[0]!.elements[0]!.id }, trigger: { type: "enter" }, timing: { kind: "span", id: "t", durationMs: 1 } }] };
document.slides[0]!.setEventTimeline(slideTiming);
const boundPlan: EventTimelinePlan | undefined = document.slides[0]!.getEventTimelinePlan();
const retainedTiming: EventTimelinePlan | undefined = resolveLayout(document).slides[0]!.eventTimeline?.plan;
document.slides[0]!.replaceElement(document.slides[0]!.elements[0]!.id, { type: "text", box: { x: 0, y: 0, width: 100, height: 40 }, content: "Replacement" });
// @ts-expect-error Authored timing trees are immutable snapshots.
document.slides[0]!.eventTimeline!.stages[0]!.target.elementId = "other";
await generatePptx(document, { eventTimelineMode: "omit" });
// @ts-expect-error Native timing export has no admitted mode.
await generatePptx(document, { eventTimelineMode: "native" });
const exportFailure: PptxExportError = new PptxExportError([]);
// @ts-expect-error Rejection diagnostics are readonly snapshots.
exportFailure.diagnostics.push({ code: "fake", message: "fake" });
void [boundPlan, retainedTiming];

import { inspectPptxTiming, type PptxTimingInspection, type PptxTimingReference } from "@pptkit/pptx-parser";
import { inspectPptxCaptionTracks, type PptxCaptionInspection } from "@pptkit/pptx-parser";
const sourceCaptions: PptxCaptionInspection = await inspectPptxCaptionTracks(generated[0]!.bytes, { captionLimits: { maxTracks: 10 } });
void sourceCaptions.tracks[0]?.language;
const sourceTiming: PptxTimingInspection = await inspectPptxTiming(generated[0]!.bytes, { timingLimits: { maxRecords: 10000, maxTextCharacters: 1000000 } });
const timingReference: PptxTimingReference | undefined = sourceTiming.slides[0]?.nodes[0]?.drawingReference;
// @ts-expect-error Located source records do not authorize a Core mapping.
sourceTiming.authoredMapping = "performed";
// @ts-expect-error Source inventories are deeply readonly.
sourceTiming.slides[0]!.nodes[0]!.attributes.push({ namespaceUri: "", local: "id", value: "fake" });
// @ts-expect-error Unknown timing budgets are not ignored.
await inspectPptxTiming(generated[0]!.bytes, { timingLimits: { maxEvents: 100 } });
void timingReference;

const paragraphTargetSession = await importPptx(generated[0]!.bytes, { projection: "content", editSaveProfile: "text-runs-v3", textTargetSaveProfile: "paragraph-stable-v1" });
paragraphTargetSession.dispose();
// @ts-expect-error Character target remapping has no admitted profile.
await importPptx(generated[0]!.bytes, { projection: "content", editSaveProfile: "text-runs-v3", textTargetSaveProfile: "character-ranges-v1" });

import { editTable, TableEditError, type TableEditOperation, type TableEditOptions, type TableEditResult, type TableCellEditMapping } from "@pptkit/core";
const tableOperations: readonly TableEditOperation[] = [{ type: "insert-column", index: 0, width: 40 }];
const tableEditOptions: TableEditOptions = { frame: "fit-grid", limits: { maxCells: 100 } };
const tableEdited: TableEditResult = editTable({ type: "table", columns: [20], rows: [{ height: 10, cells: [{ content: "Cell" }] }] }, tableOperations, tableEditOptions);
const mappedCell: TableCellEditMapping | undefined = tableEdited.cellMapping[0];
// @ts-expect-error A complete cell mapping is a readonly report.
tableEdited.cellMapping.push({});
// @ts-expect-error Geometry policy must be explicit.
editTable(tableEdited.table, tableOperations, {});
// @ts-expect-error Unknown table operation is not approximated.
editTable(tableEdited.table, [{ type: "delete-everything" }], tableEditOptions);
// @ts-expect-error Unknown limits cannot be silently accepted.
editTable(tableEdited.table, tableOperations, { frame: "preserve", limits: { maxUnknown: 1 } });
void [TableEditError, mappedCell];

import { resolveTableBorders, TableBorderError, type LayoutTableBorderGrid, type TableBorderLimits } from "@pptkit/layout";
import type { NormalizedTableElement, TableCellEdge } from "@pptkit/core";
const borderLimits: TableBorderLimits = { maxGridSlots: 100, maxEdges: 300, maxWorkUnits: 10000 };
const borderEdge: TableCellEdge = "right";
const normalizedTable = normalizePresentation(document).slides[0]?.elements.find(e => e.type === "table") as NormalizedTableElement | undefined;
if (normalizedTable) {
  const borderGrid: LayoutTableBorderGrid = resolveTableBorders(normalizedTable, borderLimits).grid;
  // @ts-expect-error Derived edge arrays are readonly.
  borderGrid.horizontal.push(borderGrid.vertical[0]!);
}
// @ts-expect-error Unknown resource limits are not ignored.
resolveTableBorders(normalizedTable!, { maxUnknown: 1 });
void [TableBorderError, borderEdge];


import type { TableStyleInput, TableRegionStyleInput, TableBandStyleInput } from "@pptkit/core";
const regionalTableStyle: TableStyleInput = { wholeTable: { text: { run: { fontFamily: { theme: "body" }, bold: false } } }, bandedRows: { size: 2, first: { fill: { type: "solid", color: { theme: "accent1" } } } }, borders: { innerHorizontal: { width: 0 } }, bottomRight: { margin: { left: 0 } } };
const region: TableRegionStyleInput = { borders: { right: { dash: "dot" } } };
const band: TableBandStyleInput = { first: region, size: 2 };
void [regionalTableStyle, band];
// @ts-expect-error Native style IDs do not enter the Core regional style contract.
const nativeRegion: TableStyleInput = { tableStyleId: "native-style-id" };
// @ts-expect-error Only supported region names are accepted.
const unknownRegion: TableStyleInput = { header: {} };
void [nativeRegion, unknownRegion];

import type { PptxTableStyleInheritance, PptxTableStylePropertyInfo } from "@pptkit/pptx-parser";
const tableStyleLimits: Partial<PptxSessionLimits> = { maxTableStyleProperties: 1000, maxTableStyleDependencies: 2000, maxTableStyleWork: 10000 };
function tableStyleProvenance(info: PptxTableStyleInheritance, property: PptxTableStylePropertyInfo) {
  const profile: "drawingml-direct-regions-v1" = info.profile;
  // @ts-expect-error Source provenance is readonly.
  info.properties.push(property);
  return [profile, property.target, info.tables[0]?.flags.firstRow];
}
void [tableStyleLimits, tableStyleProvenance];

import { resolveTextFontSegments, type TextFontSegment, type TextFontScript, type TextFontSegmentOptions, type TextRunStyleInput } from "@pptkit/core";
const scriptStyle: TextRunStyleInput = { fontFamily: "Latin", eastAsianFontFamily: { theme: "heading" }, complexScriptFontFamily: null };
const scriptOptions: TextFontSegmentOptions = { maxSegments: 1000 };
const fontSegments: readonly TextFontSegment[] = resolveTextFontSegments("A中ع", { fontFamily: "Latin", eastAsianFontFamily: "Eastern" }, normalized.theme, scriptOptions);
const fontScript: TextFontScript | undefined = fontSegments[0]?.script;
// @ts-expect-error Invalid semantic theme font role.
const invalidScriptFont: TextRunStyleInput = { eastAsianFontFamily: { theme: "eastAsian" } };
void [scriptStyle, fontScript, invalidScriptFont];

const scriptMap: import("@pptkit/core").ScriptFontFamiliesInput = { Arab: "Arabic", Hebr: { theme: "heading" }, Hant: null };
const mappedStyle: TextRunStyleInput = { scriptFontFamilies: scriptMap, language: "zh-TW" };
const scriptCode: import("@pptkit/core").TextFontScriptCode | undefined = fontSegments[0]?.scriptFont;
// @ts-expect-error Unadmitted script codes must not silently become arbitrary keys.
const invalidMap: TextRunStyleInput = { scriptFontFamilies: { Unknown: "Bad" } };
// @ts-expect-error Normalized maps contain no reset values.
const invalidNormalizedMap: import("@pptkit/core").NormalizedScriptFontFamilies = { Arab: null };
void [mappedStyle, scriptCode, invalidMap, invalidNormalizedMap];

import type { PptxImportOptions, PptxTextStyleProperty } from "@pptkit/pptx-parser";
const scriptImportOptions: PptxImportOptions = { projection: "content", textFontProfile: "script-families-v1" };
const scriptProperty: PptxTextStyleProperty = "run.eastAsianFontFamily";
function selectedFontProfile(report: ImportReport): PptxImportOptions["textFontProfile"] { return report.textFontProfile; }
const supplementalImportOptions: PptxImportOptions = { projection: "content", textFontProfile: "script-families-v2" };
const declaredSlotImportOptions: PptxImportOptions = { projection: "content", textFontProfile: "script-families-v3", notesTextProfile: "master-style-v1" };
void declaredSlotImportOptions;
const supplementalProperty: PptxTextStyleProperty = "run.scriptFontFamilies.Hant";
// @ts-expect-error Unadmitted map keys do not become public provenance properties.
const invalidSupplementalProperty: PptxTextStyleProperty = "run.scriptFontFamilies.Ethi";
// @ts-expect-error Only the named explicit font profile is admitted.
const badFontProfile: PptxImportOptions = { projection: "content", textFontProfile: "automatic" };
void [scriptImportOptions, scriptProperty, selectedFontProfile, badFontProfile, supplementalImportOptions, supplementalProperty, invalidSupplementalProperty];

const notesStyleOptions: PptxImportOptions = { projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v2", sessionLimits: { maxNotesStyleWork: 200_000 } };
function notesStyleProfile(report: ImportReport): "master-style-v1" | undefined { return report.notesTextProfile; }
const notesStyleLevel: import("@pptkit/pptx-parser").PptxTextStyleLevel = "notes-master-placeholder";
// @ts-expect-error Notes inheritance requires the named explicit profile.
const badNotesStyle: PptxImportOptions = { projection: "content", notesTextProfile: "automatic" };
void [notesStyleOptions, notesStyleProfile, notesStyleLevel, badNotesStyle];

import { inspectPptxNotes, type PptxNotesInspection, type PptxNotesRole } from "@pptkit/pptx-parser";
const sourceNotes: PptxNotesInspection = await inspectPptxNotes(generated[0]!.bytes, { notesLimits: { maxRecords: 10000 } });
const notesRole: PptxNotesRole | undefined = sourceNotes.parts[0]?.drawings[0]?.role;
// @ts-expect-error Source inventory cannot become authored geometry.
sourceNotes.geometryPolicy = "resolved";
// @ts-expect-error Inventory budgets are explicit.
await inspectPptxNotes(generated[0]!.bytes, { notesLimits: { maxPages: 100 } });
void notesRole;

import type { SlideMasterInput, NormalizedSlideMaster } from "@pptkit/core";
const authoredMaster: SlideMasterInput = { id: "isolated", name: "Isolated", theme: { colors: { accent1: "123456" } } };
const authoredMasters = createPresentation({ masters: [authoredMaster] });
authoredMasters.defineSlideLayout({ id: "owned", name: "Owned", masterId: "isolated" });
const normalizedMaster: NormalizedSlideMaster | undefined = normalizePresentation(authoredMasters).masters?.[0];
// @ts-expect-error Theme registry is owned readonly state.
authoredMasters.masters = [];
void normalizedMaster;

const managed: import("@pptkit/core").ManagedPresentationDocument = createPresentation();
const legacyCompatible: PresentationDocument = managed;
managed.defineSlideMaster({ id: "shared", name: "Shared" });
managed.replaceSlideMaster("shared", { name: "Replaced" });
managed.defineSlideLayout({ id: "shared-layout", name: "Shared" });
managed.replaceSlideLayout("shared-layout", { name: "Replaced", masterId: "shared" });
// @ts-expect-error A legacy structural document has no new editing requirement.
legacyCompatible.replaceSlideMaster("shared", { name: "Replaced" });
void legacyCompatible;

import { resolvePresentationInheritance, type PresentationInheritanceOrigin, type PresentationInheritanceResolution, type PresentationInheritanceTarget } from "@pptkit/core";
const inheritanceResolution: PresentationInheritanceResolution = resolvePresentationInheritance(managed);
const inheritedOrigin: PresentationInheritanceOrigin | undefined = inheritanceResolution.origins[0];
if (inheritedOrigin && inheritedOrigin.source.kind !== "default") {
  const declarationId: string = inheritedOrigin.source.id;
  void declarationId;
}
// @ts-expect-error Reports are immutable snapshots.
inheritanceResolution.origins.push({});
if (inheritedOrigin) {
  // @ts-expect-error Declaration paths cannot be edited.
  inheritedOrigin.source.path = "wrong";
}

interface CustomInheritanceTarget extends PresentationInheritanceTarget { readonly label?: string }
const preservedTargetInterface: CustomInheritanceTarget = { kind: "slide", id: "custom" };
if (inheritedOrigin?.target.kind === "placeholder") {
  const placeholderKey: string = inheritedOrigin.target.key;
  void placeholderKey;
}
const clearedState: "cleared" | undefined = inheritedOrigin?.state;
void [preservedTargetInterface, clearedState];

if (inheritedOrigin?.target.kind === "notes") {
  const notesSlideId: string = inheritedOrigin.target.id;
  void notesSlideId;
  // @ts-expect-error Only placeholder targets expose a key.
  inheritedOrigin.target.key;
}
const notesFormatOrigin: PresentationInheritanceOrigin = { target: { kind: "notes", id: "page" }, property: "notes.0.runs.0.style.bold", source: { kind: "slide", id: "page", path: "slides.0.notes.0.runs.0.style.bold" } };
const chartFormatOrigin: PresentationInheritanceOrigin = { target: { kind: "element", id: "chart" }, property: "axes.value.labelStyle.fontSize", source: { kind: "default", path: "defaults.chart.textStyle.fontSize" } };
void [notesFormatOrigin, chartFormatOrigin];

const tableFormatOrigin: PresentationInheritanceOrigin = { target: { kind: "element", id: "table" }, property: "rows.0.cells.0.content.0.runs.0.style.bold", source: { kind: "preset", id: "cell", path: "textStylePresets[\"cell\"].run.bold" } };
const tableBorderOrigin: PresentationInheritanceOrigin = { target: { kind: "element", id: "table" }, property: "rows.0.cells.0.style.borders.left.width", source: { kind: "element", id: "table", path: "slides.0.elements.0.style.borders.outer.width" } };
void [tableFormatOrigin, tableBorderOrigin];

import type { ManagedPresentationSlide } from "@pptkit/core";
const managedArtSlide: ManagedPresentationSlide = managed.addSlide({ showInheritedShapes: false });
managedArtSlide.setShowInheritedShapes(true); managedArtSlide.setShowInheritedShapes();
managedArtSlide.setBackground({ type: "solid", color: "123456" }); managedArtSlide.setBackground();
// @ts-expect-error background must be a paint declaration
managedArtSlide.setBackground("blue");
// @ts-expect-error Visibility accepts booleans only.
managedArtSlide.setShowInheritedShapes("false");
managed.defineSlideMaster({ id: "art-master", name: "Artwork", elements: [{ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 10, height: 10 } }] });
managed.defineSlideLayout({ id: "art-layout", name: "Art", masterId: "art-master", showMasterShapes: false });
import type { LayoutSlideMaster } from "@pptkit/layout";
const masterLayoutType: LayoutSlideMaster | undefined = undefined;
void masterLayoutType;
// @ts-expect-error Legacy slides need no new mutation capability.
legacyCompatible.slides[0]!.setShowInheritedShapes(false);


import { prepareLayoutRebind, LayoutRebindError } from "@pptkit/core";
import type { PresentationData, LayoutRebindPlan, LayoutRebindApplyResult, LayoutRebindRequest } from "@pptkit/core";
const authoringValues: PresentationData = managed;
const rebindRequests: readonly LayoutRebindRequest[] = [];
const rebindPlan: LayoutRebindPlan = prepareLayoutRebind(authoringValues, rebindRequests, { localAppearance: "adopt", reset: ["geometry", "text"] });
const rebindResult: LayoutRebindApplyResult = managed.applyLayoutRebind(rebindPlan);
const snapshotString: string = rebindPlan.sourceSnapshot;
const staleRebind: LayoutRebindError = new LayoutRebindError("stale-plan", "Replan after edits");
// @ts-expect-error A legacy document is not required to expose factory transaction methods.
legacyCompatible.applyLayoutRebind(rebindPlan);
// @ts-expect-error Plans are immutable review snapshots.
rebindPlan.requests.push({ slideId: "page", layoutId: "target" });
// @ts-expect-error Local appearance policy is explicit.
prepareLayoutRebind(authoringValues, [], {});
void [rebindResult, snapshotString, staleRebind];
const indexedPlaceholder: import("@pptkit/core").PlaceholderDefinitionInput = { key: "body", kind: "body", index: 0, box: { x: 0, y: 0, width: 100, height: 50 } };
const indexMatchMethod: import("@pptkit/core").LayoutRebindMatch["method"] = "index";
const optionalSourceIndex: number | undefined = rebindPlan.matches[0]?.sourceIndex;
// @ts-expect-error Numeric identity does not accept a source XML string.
indexedPlaceholder.index = "0";
void [indexedPlaceholder,indexMatchMethod,optionalSourceIndex];

function inspectConsumerGif(bytes: Uint8Array): GifInspection { return inspectGif(bytes, { maxFrames: 4 }); }
void inspectConsumerGif;

void generatePptx(document, { gifDiagnostics: "warn" });

void generatePptx(document, { resolveAsset: () => ({ bytes: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]), animationFallback: "first-frame" }) });

function inspectConsumerJpeg(bytes: Uint8Array): JpegExifOrientationInspection { return inspectJpegExifOrientation(bytes, { maxSegments: 32 }); }
void inspectConsumerJpeg;
function inspectConsumerPng(bytes: Uint8Array): PngInspection { return inspectPng(bytes, { maxChunks: 32 }); }
const pngIccCount: number | undefined = inspectConsumerPng(new Uint8Array()).iccProfileChunks;
void pngIccCount;
void inspectConsumerPng;
function inspectConsumerApng(bytes: Uint8Array): ApngInspection { return inspectApng(bytes, { maxFrames: 32 }); }
void inspectConsumerApng;
void generatePptx(document, { pngDiagnostics: "warn" });

void generatePptx(document, { jpegOrientationDiagnostics: "warn" });

void generatePptx(document, { resolveAsset: () => ({ bytes: new Uint8Array([137,80,78,71,13,10,26,10]), mimeType: "image/png", orientationNormalization: "pixels" }) });
void generatePptx(document, { resolveAsset: () => ({ bytes: new Uint8Array([137,80,78,71,13,10,26,10]), mimeType: "image/png", vectorRasterization: "emf" }) });
void generatePptx(document, { resolveAsset: () => ({ bytes: new Uint8Array([137,80,78,71,13,10,26,10]), mimeType: "image/png", vectorRasterization: "svg" }) });
void generatePptx(document, { resolveAsset: () => ({ bytes: new Uint8Array([137,80,78,71,13,10,26,10]), mimeType: "image/png", svgPair: new Uint8Array([60,115,118,103,47,62]) }) });
void generatePptx(document, { deduplicateResolvedAssets: true });
import { inspectJpegIcc, extractJpegIccProfile } from "@pptkit/pptx-exporter";
import { inspectPngIcc } from "@pptkit/pptx-exporter";
import { inspectIccProfile } from "@pptkit/pptx-exporter";
import { extractPngIccProfile } from "@pptkit/pptx-exporter/node";
import { extractPngIccProfileInBrowser } from "@pptkit/pptx-exporter/dom";
void inspectJpegIcc(new Uint8Array(), { maxProfileBytes: 1024 });
void extractJpegIccProfile(new Uint8Array(), { maxProfileBytes: 1024 });
void inspectPngIcc(new Uint8Array(), { maxCompressedProfileBytes: 1024 });
void inspectIccProfile(new Uint8Array(), { maxTags: 1024 });
void extractPngIccProfile(new Uint8Array(), { maxProfileBytes: 2048 });
void extractPngIccProfileInBrowser(new Uint8Array(), { maxProfileBytes: 2048 });
void generatePptx(document, { jpegIccDiagnostics: "warn" });
// @ts-expect-error Only the explicitly supported vector declarations are accepted.
void generatePptx(document, { resolveAsset: () => ({ bytes: new Uint8Array(), vectorRasterization: "pdf" }) });

async function normalizeConsumerJpeg(bytes: Uint8Array): Promise<BrowserJpegNormalizationResult> { return normalizeJpegOrientationInBrowser(bytes, { maxPixels: 100 }); }
void normalizeConsumerJpeg;

const normalizedConsumerCrop = orientJpegCrop({ left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 }, 6);
void normalizedConsumerCrop;

const replacementPolicy: ImageReplacementPolicy = { fit: "cover" };
const imageReplacementPlan = planImageReplacement(
  { type: "image", assetId: "old", box: { x: 0, y: 0, width: 100, height: 100 } },
  { id: "new", kind: "image", source: { type: "url", value: "local" }, width: 200, height: 100 },
  replacementPolicy,
);
void imageReplacementPlan;

const visibleReplacementFit: VisibleImageReplacementFit = "stretch";
const visibleReplacementOptions: VisibleImageReplacementOptions = { composition: [{ type: "align", container: { kind: "slide", id: "slide" }, elementIds: ["pic"], horizontal: "center" }] };
void planVisibleImageReplacement; void planNestedVisibleImageReplacement; void planSharedVisibleImageReplacement; void visibleReplacementFit; void visibleReplacementOptions;
const nestedCommitOptions: NestedVisibleImageReplacementOptions = { commitAncestorGeometry: true };
const sharedCommitOptions: SharedVisibleImageReplacementOptions = { commitAncestorGeometry: true };
void nestedCommitOptions; void sharedCommitOptions;
type NestedPlanContract = NestedVisibleImageReplacementPlan;
type SharedPlanContract = SharedVisibleImageReplacementPlan;

async function inspectConsumerMasterPictures(bytes: Uint8Array): Promise<MasterPictureInventory> { return inspectPptxMasterPictures(bytes, { masterPictureLimits: { maxRecords: 100 } }); }
void inspectConsumerMasterPictures;

import type { AssetRegistrationRequest, ManagedPresentationDocument } from '@pptkit/core';
const batchDocument: ManagedPresentationDocument = createPresentation();
const batchRequests: readonly AssetRegistrationRequest[] = [{ asset: { kind: 'image', source: { type: 'bytes', value: new Uint8Array([1]) } }, options: { deduplicate: 'content' } }];
const batchAssets = batchDocument.registerAssets(batchRequests);
// @ts-expect-error Batch result is readonly.
batchAssets.push(batchAssets[0]!);
// @ts-expect-error Batch registration is an additive managed-document capability.
document.registerAssets(batchRequests);

import type { AssetCopyOptions, AssetCopyMapping } from '@pptkit/core';
const assetCopyOptions: AssetCopyOptions = { deduplicate: 'content', external: 'reject' };
const assetCopyMapping: readonly AssetCopyMapping[] = batchDocument.copyAssetsFrom(createPresentation(), [], assetCopyOptions);
// @ts-expect-error Mappings are immutable snapshots.
assetCopyMapping[0]!.destinationId = 'changed';
// @ts-expect-error Both copy policies are required.
batchDocument.copyAssetsFrom(createPresentation(), [], { deduplicate: 'content' });
void assetCopyMapping;

const installedSmallDoughnut: import("@pptkit/core").DoughnutChartElementInput = { type: "chart", chartType: "doughnut", holeSize: 1, categories: ["A"], series: [{ name: "S", values: [1] }], box: { x: 0, y: 0, width: 100, height: 100 } };
void installedSmallDoughnut;

const installedArea: import('@pptkit/core').AreaChartElementInput = {type:'chart',chartType:'area',grouping:'percentStacked',box:{x:0,y:0,width:100,height:100},categories:['A'],series:[{name:'S',values:[null]}]};
void installedArea;

const installedScatter: import('@pptkit/core').ScatterChartElementInput = { type:'chart',chartType:'scatter',box:{x:0,y:0,width:100,height:100},series:[{name:'S',xValues:[1,null],yValues:[2,3]}],axes:{x:{scale:{min:-1,max:1,majorUnit:.5}},y:{scale:'auto'}} };
void installedScatter;

const installedScatterPoint: import('@pptkit/layout').ScatterGeometryPoint = {x:1,y:2,sourceIndex:0,blank:true};
void installedScatterPoint;
