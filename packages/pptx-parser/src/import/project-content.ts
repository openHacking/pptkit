import { readParallelogramSlant } from "./parallelogram-slant.js";
import type { ConnectorElementInput, ElementAction, ManagedPresentationDocument, PresentationElementInput, ShapeKind } from "@pptkit/core";
import { attribute, children, integer, invalidStructure } from "../presentation/structure-xml.js";
import type { structureContext } from "../presentation/structure-context.js";
import type { XmlElement } from "../xml/read-xml.js";
import { PptxImportError, type ImportedNotesTextRunSource, type ImportedPictureSource, type ImportedTableTextRunSource, type ImportedTextRunSource, type ImportFinding, type ImportReport, type PptxImportOptions, type PptxSourceLocation, type PptxSessionLimits } from "../types/import.js";
import type { OpcInspectionOptions } from "../types/opc.js";
import { freezeReport } from "./state.js";
import { contentReader } from "./content-reader.js";
import { readDecorative } from "./decorative.js";
import { bindDrawingDescription, bindDrawingDecorative, type DrawingDescriptionBinding } from "./drawing-description-binding.js";
import { contentText, type ContentTextRun } from "./content-text.js";
import { bindTextRuns, type TextRunBinding } from "./text-run-bindings.js";
import { contentImages } from "./content-images.js";
import { audioSources } from "./audio-sources.js";
import { audioTrim } from "./audio-trim.js";
import { audioPlayback } from "./audio-playback.js";
import { audioObjectStart } from "./audio-object-start.js";
import { audioMainStart } from "./audio-main-start.js";
import { plainVideoTimingSource } from "./video-timing-source.js";
import { videoSources } from "./video-sources.js";
import { contentAction } from "./content-actions.js";
import { connectorAttachment, contentConnector, type ContentConnection, type ContentDrawing } from "./content-connectors.js";
import { contentRoundPath } from "./content-round-path.js";
import { contentNotes } from "./content-notes.js";
import { contentTable, type ContentTableTextRun } from "./content-table.js";
import { contentCharts } from "./content-chart.js";
import { placeholderInheritance } from "./placeholder-inheritance.js";
import { tableStyleInheritance } from "./table-style-inheritance.js";
import { textInheritance } from "./text-inheritance.js";
import { diagramPreview } from "./diagram-preview.js";
import type { ChartValueBinding } from "./chart-value-bindings.js";
import type { ChartLabelBinding } from "./chart-label-bindings.js";
import type { ChartSeriesOrderBinding } from "./chart-series-order.js";
import type { ImportedDiagramPreview } from "../types/import.js";
import { artworkPaintReader } from "./artwork-paint-reader.js";
import { projectMasterPictures } from "./project-master-pictures.js";
import type { MasterOwner } from "./project-structure.js";

const shapes: Readonly<Record<string, ShapeKind>> = { rect: "rect", roundRect: "roundRect", ellipse: "ellipse", triangle: "triangle", diamond: "diamond", rightArrow: "arrow", chevron: "chevron", parallelogram: "parallelogram" };

export async function projectContent(document: ManagedPresentationDocument, base: ImportReport, context: Awaited<ReturnType<typeof structureContext>>, limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>, options: OpcInspectionOptions, bindings?: TextRunBinding[], profile?: PptxImportOptions["editSaveProfile"], diagramPreviewMode?: PptxImportOptions["diagramPreview"], chartBindings?: ChartValueBinding[], textFontProfile?: PptxImportOptions["textFontProfile"], notesTextProfile?: PptxImportOptions["notesTextProfile"], masterOwners?: readonly MasterOwner[], masterShapes = false, masterText = false, masterTextTheme = false, masterPaintTheme = false, backgrounds = false, backgroundMatrix = false, backgroundGradient = false, shapeGradient = false, slideShapeGradient = false, shapePattern = false, shapeShadow = false, pictureSources?: ImportedPictureSource[], fieldTextProfile?: PptxImportOptions["fieldTextProfile"], descriptionBindings?: DrawingDescriptionBinding[], descriptionAllowPresenceChange = false, drawingName = false, descriptionAndName = false, decorativeEnabled = false, decorativeOnly = false, chartSparseWrites = false, chartV3Writes = false, chartLabelBindings?: ChartLabelBinding[], chartSeriesNames = false, chartOrderBindings?: ChartSeriesOrderBinding[], researchColumnLegend = false, mediaPreviewMode?: PptxImportOptions["mediaPreview"], audioSourceProfile?: PptxImportOptions["audioSourceProfile"], audioProjection?: PptxImportOptions["audioProjection"], videoSourceProfile?: PptxImportOptions["videoSourceProfile"], videoProjection?: PptxImportOptions["videoProjection"]): Promise<ImportReport> {
  const includeTextOwners = profile === "text-runs-v2" || profile === "text-runs-v3";
  const report = structuredClone(base);
  const diagramPreviews: ImportedDiagramPreview[] = [];
  const audio = audioSourceProfile || audioProjection ? audioSources(context, limits, checkpoint) : undefined;
  const video = videoSourceProfile || videoProjection ? videoSources(context, limits, checkpoint) : undefined;
  let previewWork = 0;
  const previewCheckpoint = async () => {
    if (++previewWork > limits.maxDiagramPreviewWork) throw new PptxImportError("import-limit-exceeded", "Cached diagram traversal exceeds maxDiagramPreviewWork.");
    await checkpoint();
  };
  const findings = report.findings.map(f => f.feature === "source-part-content" ? { ...f, reason: "This part has no complete semantic mapping. Separately reported structure/content fields are the mapped subset." } : f.feature === "slide-order-and-visibility" ? { ...f, reason: "Source slide order and hidden state are mapped; drawing fields have separate capability findings." } : f);
  const identities = [...report.identities], diagnostics = report.diagnostics.filter(d => d.code !== "structural-projection-only");
  let slidePaintDependencies = 0;
  const projectedParts = new Set<string>(), audioAssets = new Map<string, string>(), videoAssets = new Map<string, string>();
  let projectedBytes = 0;
  const chargeAsset = audioProjection || videoProjection ? (part: string, size: number) => {
    if (projectedParts.has(part)) return;
    if (size > limits.maxImportedAssetBytes || projectedBytes + size > limits.maxTotalImportedAssetBytes) throw new PptxImportError("import-limit-exceeded", "Projected media and images exceed imported asset byte limits.", { part });
    projectedParts.add(part); projectedBytes += size;
  } : undefined;
  const image = contentImages(document, context, limits, identities, checkpoint, chargeAsset);
  const chart = contentCharts(context, limits, options, base.inspection.package.relationships, checkpoint, chartBindings !== undefined || chartLabelBindings !== undefined || chartOrderBindings !== undefined, chartSparseWrites, chartV3Writes, chartLabelBindings !== undefined, chartSeriesNames, chartOrderBindings !== undefined, researchColumnLegend);
  const inheritance = placeholderInheritance(context, base.inspection, limits, checkpoint);
  const slideIds = new Map(base.inspection.slides.map((source, index) => [source.part, document.slides[index]!.id]));
  const finding = (feature: string, operation: ImportFinding["operation"], tier: ImportFinding["tier"], source: PptxSourceLocation, reason: string, slideId?: string, elementId?: string, resolvedFrom?: readonly PptxSourceLocation[]) => {
    if (findings.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Import capability report exceeds maxReportFindings.", source);
    findings.push({ id: `finding-${findings.length + 1}`, feature, operation, tier, source, reason, ...(slideId ? { slideId } : {}), ...(elementId ? { elementId } : {}), ...(resolvedFrom?.length ? { resolvedFrom } : {}) });
  };
  const textCascade = textInheritance(context, base.inspection, limits, checkpoint, inheritance.paths,
    (source, reason) => finding("drawing.inherited-text.unmapped-property", "import", "fallback", source, reason), textFontProfile, (source, reason) => finding("notes.inherited-text.unmapped-property", "import", "fallback", source, reason));
  if (masterOwners) {
    const masterParts = new Map(masterOwners.filter(owner => owner.kind === "master").map(owner => [owner.id, owner.part]));
    await projectMasterPictures(document, masterOwners, context, image, identities, checkpoint, finding, masterShapes, masterText,
      masterTextTheme ? (owner, r, paths, at) => owner.kind === "slide" ? textCascade.forSlideArtwork(owner.masterId ? masterParts.get(owner.masterId) : undefined, owner.layoutPart, owner.part, r, paths, at) : textCascade.forArtwork(owner.kind === "master" ? owner.part : masterParts.get(owner.masterId!)!, owner.kind === "layout" ? owner.part : undefined, r, paths, at) : undefined, masterPaintTheme, limits.maxTextStyleDependencies, backgrounds, backgroundMatrix, backgroundGradient, shapeGradient, shapePattern, shapeShadow);
    diagnostics.push({ code: "master-artwork-partial", severity: "warning", source: { part: base.inspection.mainPart }, message: `Only ordinary master/layout pictures${masterShapes ? ", eligible shapes" : ""}${masterText ? " with partial direct text" : ""}, containing groups and verified listed ownership paths are projected. Placeholders, other artwork, themes, backgrounds and effects remain incomplete; source-preserving edits are not enabled.` });
  }
  const tableCascade = tableStyleInheritance(context, base.inspection, limits, checkpoint, (source, reason) => finding("drawing.table-style.unmapped-property", "import", "fallback", source, reason), textFontProfile);
  for (let index = 0; index < base.inspection.slides.length; index++) {
    await checkpoint();
    const slide = document.slides[index]!, sourceSlide = base.inspection.slides[index]!, part = sourceSlide.part;
    const diagramByIndex = new Map(sourceSlide.diagrams?.map(d => [d.elementIndex, d]));
    const root = await context.root(part, "sld"), r = contentReader(part, context.p, context.a);
    const audioSettings = audioProjection ? await audioPlayback(root, context.p, checkpoint) : undefined;
    const audioStart = audioProjection ? await audioObjectStart(root, context.p, checkpoint) : undefined;
    const videoPauses = videoProjection ? await audioObjectStart(root, context.p, checkpoint, "video-pause") : undefined;
    const audioMain = audioProjection || videoProjection ? await audioMainStart(root, context.p, checkpoint, videoPauses) : undefined;
    const audioMainInputs = new Map<string, { target: string; duration?: number; trimUnsupported: boolean; playbackMapped: boolean }>();
    const videoMainInputs = new Map<string, { target: string; duration?: number; eligible: boolean }>();
    const placeholders = await inheritance.slide(part, sourceSlide.layoutPart);
    r.known(root); r.attr(root, "show");
    const common = r.one(root, r.p, "cSld"), tree = common ? r.one(common, r.p, "spTree") : undefined;
    if (common) { r.known(common); r.attr(common, "name"); }
    const seen = new Set<number>(), checkedIdentities = new WeakSet<XmlElement>(), owners = new Map<string, { id?: string; sourceObjectId?: string }>();
    const drawings = new Map<number, ContentDrawing>(), pendingConnections: { input: ConnectorElementInput; scope: string; source: PptxSourceLocation; connections: ContentConnection[] }[] = [];
    const drawingIdentity = (node: XmlElement): number => {
      const value = integer(attribute(node, "id"), 0, 4294967295, "cNvPr/@id", part);
      if (!checkedIdentities.has(node)) {
        if (seen.has(value)) invalidStructure("Duplicate drawing identity in a slide part.", part);
        seen.add(value); checkedIdentities.add(node);
      }
      return value;
    };
    const pathByNode = placeholders.paths;
    const slideOwner = masterOwners?.find(owner => owner.kind === "slide" && owner.id === slide.id);
    const masterPart = slideOwner?.masterId ? masterOwners?.find(owner => owner.kind === "master" && owner.id === slideOwner.masterId)?.part : undefined;
    const gradientReader = slideShapeGradient ? artworkPaintReader(r, checkpoint, (color, source) => textCascade.forSlideArtwork(masterPart, sourceSlide.layoutPart, part, r, pathByNode, source).local.resolveColor!(color), (count, source) => {
      if (count > limits.maxTextStyleDependencies - slidePaintDependencies) throw new PptxImportError("import-limit-exceeded", "Slide artwork paint dependencies exceed maxTextStyleDependencies.", source);
      slidePaintDependencies += count;
    }) : undefined;
    const projectedBackground = backgrounds && common ? r.one(common, r.p, "bg") : undefined;
    const projectedBackgroundPath = projectedBackground ? pathByNode.get(projectedBackground) : undefined;
    const drawingNodes = tree ? [tree] : [];
    while (drawingNodes.length) {
      await checkpoint();
      const node = drawingNodes.pop()!;
      if (node.namespaceUri !== r.p) continue;
      const container = { spTree: "nvGrpSpPr", grpSp: "nvGrpSpPr", sp: "nvSpPr", pic: "nvPicPr", cxnSp: "nvCxnSpPr", graphicFrame: "nvGraphicFramePr" }[node.local];
      if (!container) continue;
      const nv = r.one(node, r.p, container), cnv = nv ? r.one(nv, r.p, "cNvPr") : undefined;
      if (cnv) {
        const sourceId = drawingIdentity(cnv);
        if (audio && node.local === "pic") await audio.inspect(part, slide.id, String(sourceId), nv ? r.one(nv, r.p, "nvPr") : undefined, pathByNode);
        if (video && node.local === "pic") await video.inspect(part, slide.id, String(sourceId), node, pathByNode);
      }
      if (node.local === "spTree" || node.local === "grpSp") for (const child of children(node).reverse()) drawingNodes.push(child);
    }
    const capabilities = (feature: string, source: PptxSourceLocation, reason: string, mapped: boolean, id?: string, preview?: ImportFinding["tier"]) => {
      for (const operation of ["create", "import", "edit", "preview", "export", "no-op-save", "edit-save"] as const) {
        const tier = operation === "no-op-save" ? "preserved" : operation === "edit-save" || !mapped ? "unsupported" : operation === "preview" && preview ? preview : "editable";
        finding(feature, operation, tier, source, reason, slide.id, id);
      }
    };
    const metadata = (node: XmlElement, container: string) => {
      const nv = r.one(node, r.p, container), cnv = nv ? r.one(nv, r.p, "cNvPr") : undefined;
      if (!cnv) invalidStructure(`Missing ${container}/cNvPr drawing identity.`, part);
      r.known(nv!);
      r.attr(cnv!, "id");
      const sourceId = drawingIdentity(cnv!);
      const name = r.attr(cnv!, "name") ?? "", description = r.attr(cnv!, "descr"), hidden = r.bool(cnv!, "hidden"), decorative = readDecorative(cnv!, r);
      return { sourceId, name, hidden, cnv: cnv!, ...(decorative ? { decorativeSource: decorative.node } : {}), ...(description !== undefined || decorative !== undefined ? { accessibility: { ...(description !== undefined ? { description } : {}), ...(decorative !== undefined ? { decorative: decorative.value } : {}) } } : {}) };
    };
    const readAction = (owner: XmlElement, source: PptxSourceLocation, elementId: string): ElementAction | undefined => {
      const result = contentAction(owner, r, context, slideIds);
      if (!result) return undefined;
      const at = { ...source, path: pathByNode.get(result.node)!, ...(result.relationshipId !== undefined ? { relationshipId: result.relationshipId } : {}) };
      capabilities("drawing.click-action", at, result.reason, Boolean(result.action), elementId, result.action?.type === "slide" ? "fallback" : undefined);
      if (!result.action) {
        r.note(result.node, result.reason);
        diagnostics.push({ code: "click-action-unmapped", severity: "warning", source: at, slideId: slide.id, elementId, message: result.reason });
      }
      if (result.action?.type === "slide") finding("drawing.click-action.navigation", "preview", "fallback", at, "Standalone SVG retains slide-action metadata; consumer-controlled navigation is not implemented by the renderer.", slide.id, elementId);
      return result.action;
    };
    const pictureMedia = async (node: XmlElement): Promise<XmlElement | undefined> => {
      const nv = r.one(node, r.p, "nvPicPr"), nvPr = nv && r.one(nv, r.p, "nvPr");
      if (!nvPr) return undefined;
      const items = children(nvPr);
      for (const item of items) {
        await checkpoint();
        if ((item.namespaceUri === r.a || item.namespaceUri === r.p) && ["audioFile", "videoFile", "quickTimeFile", "audioCd", "wavAudioFile"].includes(item.local)) return item;
      }
      for (const list of items) if (list.namespaceUri === r.p && list.local === "extLst") for (const ext of children(list)) {
        await checkpoint();
        if (ext.namespaceUri !== r.p || ext.local !== "ext" || attribute(ext, "uri")?.toUpperCase() !== "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}") continue;
        for (const child of children(ext)) { await checkpoint(); if (child.namespaceUri === "http://schemas.microsoft.com/office/powerpoint/2010/main" && child.local === "media") return child; }
      }
      return undefined;
    };
    const parse = async (node: XmlElement, origin: { x: number; y: number }, scope: string): Promise<PresentationElementInput | undefined> => {
      await checkpoint();
      const path = pathByNode.get(node)!;
      if (node.namespaceUri !== r.p) return undefined;
      const container = { sp: "nvSpPr", grpSp: "nvGrpSpPr", pic: "nvPicPr", cxnSp: "nvCxnSpPr", graphicFrame: "nvGraphicFramePr" }[node.local];
      if (!container) return undefined;
      const meta = metadata(node, container);
      const at = { part, path, sourceObjectId: String(meta.sourceId) };
      const id = `${slide.id}-object-${meta.sourceId}`;
      const media = node.local === "pic" ? await pictureMedia(node) : undefined;
      const reportMedia = (posterMapped: boolean, placeholder = false, projected?: "audio" | "video") => {
        if (!media) return;
        const mediaAt = { ...at, path: pathByNode.get(media)! };
        const reason = projected ? `The ${projected === "audio" ? "PCM audio" : "video"} and optional poster are projected, but some source playback settings and timing are not projected or executed.` : posterMapped ? "The static image is a media poster only; media identity, playback and timing are not projected or executed." : placeholder ? "The media object has no mapped static poster; a static placeholder is shown without media identity, playback or timing." : "The media object has no mapped static poster; media identity, playback and timing are not projected or executed.";
        capabilities("drawing.media", mediaAt, reason, false, posterMapped || placeholder ? id : undefined);
        finding("drawing.media.poster", "preview", posterMapped || placeholder ? "fallback" : "unsupported", mediaAt, reason, slide.id, posterMapped || placeholder ? id : undefined);
        diagnostics.push({ code: "media-playback-unmapped", severity: "warning", source: mediaAt, slideId: slide.id, ...(posterMapped || placeholder ? { elementId: id } : {}), message: reason });
      };
      const placeholder = placeholders.get(node);
      if (placeholder) {
        const nv = r.one(node, r.p, container)!, nvPr = r.one(nv, r.p, "nvPr")!, ph = r.one(nvPr, r.p, "ph")!;
        r.known(nvPr); r.attr(ph, "type"); r.attr(ph, "idx");
      }
      const reportPlaceholder = (elementId?: string) => {
        if (!placeholder) return;
        const parent = placeholder.parent!;
        finding("drawing.placeholder-match", "import", parent.status === "matched" ? "editable" : "fallback", placeholder.declaration,
          `Placeholder ${parent.rule} ${JSON.stringify(parent.key)}: ${parent.status}; ${parent.candidateCount} candidates. ${parent.reason ?? "Matching identifies source ancestry only; styles and template rebinding remain pending."}`, slide.id, elementId);
      };
      owners.set(path, { sourceObjectId: at.sourceObjectId });
      const unsupported = (reason: string) => { reportMedia(false); reportPlaceholder(); capabilities(`drawing.${node.local}`, at, reason, false); return undefined; };
      const properties = node.local === "graphicFrame" ? node : r.one(node, r.p, node.local === "grpSp" ? "grpSpPr" : "spPr");
      r.known(node);
      if (!properties) return unsupported("Drawing properties are absent; inherited geometry is not synthesized.");
      r.known(properties);
      const localGeometry = r.transform(properties, node.local === "graphicFrame" ? r.p : r.a);
      const inheritedGeometry = placeholders.geometry(node);
      const geometry = localGeometry ?? inheritedGeometry;
      if (!geometry) return unsupported(placeholder?.reason ?? "Complete local or unambiguous supported placeholder offset/extent is required.");
      if (!localGeometry && inheritedGeometry) {
        const xfrm = r.one(properties, node.local === "graphicFrame" ? r.p : r.a, "xfrm");
        if (xfrm) {
          r.attr(xfrm, "rot"); r.attr(xfrm, "flipH"); r.attr(xfrm, "flipV");
          const off = r.one(xfrm, r.a, "off"), ext = r.one(xfrm, r.a, "ext");
          if (off) { r.attr(off, "x"); r.attr(off, "y"); }
          if (ext) { r.attr(ext, "cx"); r.attr(ext, "cy"); }
        }
      }
      geometry.box.x -= origin.x; geometry.box.y -= origin.y;
      const baseInput = { id, name: meta.name, hidden: meta.hidden, ...geometry, ...(meta.accessibility ? { accessibility: meta.accessibility } : {}) };
      let input: PresentationElementInput;
      let previewFallback = false;
      let cachedPreview = false;
      let mediaPlaceholder = false;
      if (node.local === "graphicFrame") {
        const graphic = r.one(node, r.a, "graphic"), data = graphic ? r.one(graphic, r.a, "graphicData") : undefined;
        const uri = data ? attribute(data, "uri") : undefined;
        const framePr = r.one(r.one(node, r.p, container)!, r.p, "cNvGraphicFramePr");
        if (framePr) r.known(framePr);
        if (diagramPreviewMode && uri === r.a.replace(/\/main$/, "/diagram")) {
          r.known(graphic!); r.attr(data!, "uri");
          const rawIndex = (await context.sourceElement(part, node)).elementIndex;
          const sourceDiagram = diagramByIndex.get(rawIndex);
          const result = await diagramPreview(sourceDiagram, baseInput, context, image, previewCheckpoint, (source, reason) => finding("smartart.cache.unmapped-property", "preview", "fallback", source, reason, slide.id, id));
          input = result.input; cachedPreview = true; previewFallback = true;
          diagramPreviews.push({ slideId: slide.id, elementId: id, source: at, status: result.status, freshness: "unknown", ...(result.cachePart ? { cachePart: result.cachePart } : {}), ...(result.reason ? { reason: result.reason } : {}) });
          for (const operation of ["create", "import", "edit", "preview", "export", "no-op-save", "edit-save"] as const) finding("smartart.cache-preview", operation,
            operation === "no-op-save" ? "preserved" : operation === "import" || operation === "preview" ? "fallback" : "unsupported", at,
            result.reason ?? "Cached drawing approximation of unknown freshness; derived Core shapes are not semantic SmartArt nodes or source-edit bindings.", slide.id, id);
        } else if (["http://schemas.openxmlformats.org/drawingml/2006/chart", "http://purl.oclc.org/ooxml/drawingml/chart"].includes(uri ?? "")) {
          r.known(graphic!); r.attr(data!, "uri");
          const chartFindings: { feature: string; source: PptxSourceLocation; reason: string }[] = [];
          const result = await chart(data!, geometry.box, r, at.sourceObjectId, (feature, source, reason) => {
            if (chartFindings.length + findings.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Chart report exceeds maxReportFindings.", source);
            chartFindings.push({ feature, source, reason });
          });
          const mappedId = "fields" in result ? id : undefined;
          for (const item of chartFindings) {
            if (item.feature === "chart.unmapped-property") capabilities(item.feature, item.source, item.reason, false, mappedId);
            else finding(item.feature, "import", "fallback", item.source, item.reason, slide.id, mappedId);
          }
          if ("reason" in result) return unsupported(result.reason);
          if (chartBindings) for (const value of result.valueBindings ?? []) {
            if (chartBindings.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Chart value bindings exceed maxReportFindings.", value.source);
            chartBindings.push({ ...value, slideId: slide.id, elementId: id });
          }
          if (chartLabelBindings && "categories" in result.fields) for (const closure of result.labelCandidates ?? []) {
            const match = /^category:(\d+)$/u.exec(closure.logicalKey), categoryIndex = match ? Number(match[1]) : -1;
            const baselineValue = result.fields.categories[categoryIndex];
            if (!Number.isSafeInteger(categoryIndex) || categoryIndex < 0 || typeof baselineValue !== "string") throw new PptxImportError("import-limit-exceeded", "Chart label candidate identity is invalid.", at);
            if (chartLabelBindings.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Chart label bindings exceed maxReportFindings.", at);
            chartLabelBindings.push({ slideId: slide.id, elementId: id, categoryIndex, baselineValue, closure });
          }
          if (chartLabelBindings && chartSeriesNames) for (const closure of result.seriesNameCandidates ?? []) {
            const match = /^series:(\d+)$/u.exec(closure.logicalKey), seriesIndex = match ? Number(match[1]) : -1;
            const baselineValue = result.fields.series[seriesIndex]?.name;
            if (!Number.isSafeInteger(seriesIndex) || seriesIndex < 0 || typeof baselineValue !== "string") throw new PptxImportError("import-limit-exceeded", "Chart series-name candidate identity is invalid.", at);
            if (chartLabelBindings.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Chart label bindings exceed maxReportFindings.", at);
            chartLabelBindings.push({ slideId: slide.id, elementId: id, seriesIndex, baselineValue, closure });
          }
          if (chartOrderBindings && result.seriesOrderCandidate) {
            if (chartOrderBindings.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Chart series-order bindings exceed maxReportFindings.", at);
            chartOrderBindings.push({ slideId: slide.id, elementId: id, slidePart: at.part, ...result.seriesOrderCandidate });
          }
          input = { ...baseInput, ...result.fields }; previewFallback = result.previewFallback;
        } else {
          const tableRuns: ContentTableTextRun[] | undefined = bindings && profile === "text-runs-v3" ? [] : undefined;
          const result = await contentTable(node, geometry.box, r, checkpoint, owner => readAction(owner, at, id), tableRuns ? run => {
            if (bindings!.length + tableRuns.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Text run source bindings exceed maxReportFindings.", at);
            tableRuns.push(run);
          } : undefined, props => tableCascade.forTable(props, r, pathByNode, at, sourceSlide.layoutPart), fieldTextProfile);
          if ("reason" in result) return unsupported(result.reason);
          input = { ...baseInput, ...result.fields }; previewFallback = result.previewFallback;
          if (bindings && tableRuns) for (const run of tableRuns) {
            const target = { slideId: slide.id, elementId: id, owner: "table-cell" as const, rowIndex: run.rowIndex, cellIndex: run.cellIndex, columnIndex: run.columnIndex };
            const source = { ...at, path: pathByNode.get(run.cell)! };
            for (const binding of await bindTextRuns([run], target, source, context, checkpoint)) bindings.push(binding);
          }
        }
      } else if (node.local === "cxnSp") {
        const result = contentConnector(node, properties, geometry, r);
        if ("reason" in result) return unsupported(result.reason);
        input = { ...baseInput, ...result.fields };
        pendingConnections.push({ input, scope, source: at, connections: result.connections });
      } else if (node.local === "pic") {
        const nv = r.one(node, r.p, container)!, nvPr = r.one(nv, r.p, "nvPr");
        const sources = audio?.forOwner(nvPr) ?? [], direct = sources.filter(item => item.declaration !== "p14:media"), companion = sources.filter(item => item.declaration === "p14:media");
        const declarations = nvPr ? children(nvPr).filter(item => (item.namespaceUri === r.a || item.namespaceUri === r.p) && ["audioFile", "wavAudioFile", "videoFile", "quickTimeFile", "audioCd"].includes(item.local)) : [];
        const primary = direct.length === 1 ? direct[0] : undefined;
        const office = companion.length === 1 ? companion[0] : undefined;
        const trim = audioTrim(nvPr, r.p, (primary?.status === "wave-pcm" ? primary : office)?.pcm?.durationSeconds);
        const nativeTrimmed = primary?.declaration === "audioFile" && primary.status === "external" && primary.externalTarget === "NULL" && office?.relationshipAttribute === "embed" && office.status === "wave-pcm" && office.companionStatus === "unverified" && trim.status === "mapped" && (trim.trimStartMs > 0 || trim.trimEndMs > 0);
        const candidate = audioProjection && declarations.length === 1 && geometry.box.width > 0 && geometry.box.height > 0 ? nativeTrimmed ? office : primary?.status === "wave-pcm" && primary.mediaPart && (companion.length === 0 || office?.companionStatus === "same-part" && office.status === "wave-pcm") ? primary : undefined : undefined;
        const videoSources = video?.forPicture(node) ?? [], videoDirect = videoSources.filter(item => item.declaration === "videoFile"), videoCompanion = videoSources.filter(item => item.declaration === "p14:media");
        const videoCandidate = videoProjection && videoDirect.length === 1 && videoDirect[0]!.status === "iso-bmff-framing" && videoDirect[0]!.mediaPart && declarations.length === 1 && (videoCompanion.length === 0 || videoCompanion.length === 1 && videoCompanion[0]!.companionStatus === "same-part" && videoCompanion[0]!.status === "iso-bmff-framing") && geometry.box.width > 0 && geometry.box.height > 0 ? videoDirect[0] : undefined;
        const result = await image(node, properties, r, at);
        const posterSupported = !("reason" in result) && result.fields.fit === "stretch" && (result.fields.opacity === undefined || result.fields.opacity === 1);
        if (audioProjection && sources.length && (!candidate || !("reason" in result) && !posterSupported)) finding("drawing.audio.projected-fields", "import", "unsupported", at, "Audio projection requires one unambiguous embedded PCM source, matching Office companion references, nonzero geometry and a representable poster.", slide.id);
        if (videoProjection && videoSources.length && (!videoCandidate || !("reason" in result) && !posterSupported)) finding("drawing.video.projected-fields", "import", "unsupported", at, "Video projection requires one unambiguous embedded MP4 framing source, matching Office companion references, nonzero geometry and a representable poster.", slide.id);
        if (candidate && ("reason" in result || posterSupported)) {
          const mediaPart = candidate.mediaPart!, payload = context.payload(mediaPart);
          let audioAssetId = audioAssets.get(mediaPart);
          if (!audioAssetId) {
            chargeAsset!(mediaPart, payload.bytes.byteLength);
            audioAssetId = `source-audio-${audioAssets.size + 1}`;
            document.registerAsset({ id: audioAssetId, kind: "audio", mimeType: payload.contentType, source: { type: "bytes", value: payload.bytes } });
            audioAssets.set(mediaPart, audioAssetId); identities.push({ kind: "asset", id: audioAssetId, source: { part: mediaPart } });
          }
          identities.push({ kind: "asset", id: audioAssetId, source: candidate.source });
          const playback = audioSettings?.get(at.sourceObjectId), start = audioStart?.get(at.sourceObjectId);
          const duration = candidate.pcm ? Math.round(candidate.pcm.durationSeconds * 1000) - (trim.status === "mapped" ? trim.trimStartMs + trim.trimEndMs : 0) : undefined;
          const startMapped = start?.status === "mapped" && trim.status !== "unsupported" && start.durationMs === duration;
          input = { ...baseInput, type: "audio", audioAssetId, ...(trim.status === "mapped" ? { trimStartMs: trim.trimStartMs, trimEndMs: trim.trimEndMs } : {}), ...(playback?.status === "mapped" ? playback.settings : {}), ...(startMapped ? { playbackStart: "object-click" as const } : {}), ...(!("reason" in result) ? { posterAssetId: result.fields.assetId } : {}) };
          if (start && !startMapped) finding("drawing.audio.start", "import", "unsupported", at, "Ambiguous or unsupported object-click timing was not projected.", slide.id, id);
          audioMainInputs.set(id, { target: at.sourceObjectId, ...(duration === undefined ? {} : { duration }), trimUnsupported: trim.status === "unsupported", playbackMapped: playback?.status === "mapped" });
          if (playback?.status === "unsupported") finding("drawing.audio.playback", "import", "unsupported", at, "Ambiguous or unsupported audio media timing settings were not projected.", slide.id, id);
          if (trim.status === "unsupported") finding("drawing.audio.trim", "import", "unsupported", at, "Audio trim cannot be projected as bounded integer milliseconds; playback uses the full source.", slide.id, id);
          cachedPreview = true; mediaPlaceholder = true;
          reportMedia(!("reason" in result), "reason" in result, "audio");
          for (const operation of ["create", "import", "edit", "preview", "export", "no-op-save", "edit-save"] as const) {
            const tier = operation === "no-op-save" ? "preserved" : operation === "edit-save" || operation === "export" ? "unsupported" : operation === "import" || operation === "preview" ? "fallback" : "editable";
            finding("drawing.audio.projected-fields", operation, tier, at, "Audio bytes, source identity, poster, geometry, supported trim, media settings and bounded object-click/main-sequence starts are projected; other playback timing, general export and source edit-save remain unsupported.", slide.id, id);
          }
          if ("reason" in result) finding("drawing.audio.poster", "preview", "fallback", at, `No supported poster was projected: ${result.reason}`, slide.id, id);
        } else if (videoCandidate && ("reason" in result || posterSupported)) {
          const mediaPart = videoCandidate.mediaPart!, payload = context.payload(mediaPart);
          let videoAssetId = videoAssets.get(mediaPart);
          if (!videoAssetId) {
            chargeAsset!(mediaPart, payload.bytes.byteLength);
            videoAssetId = `source-video-${videoAssets.size + 1}`;
            document.registerAsset({ id: videoAssetId, kind: "video", mimeType: payload.contentType, source: { type: "bytes", value: payload.bytes } });
            videoAssets.set(mediaPart, videoAssetId); identities.push({ kind: "asset", id: videoAssetId, source: { part: mediaPart } });
          }
          identities.push({ kind: "asset", id: videoAssetId, source: videoCandidate.source });
          input = { ...baseInput, type: "video", videoAssetId, ...(!("reason" in result) ? { posterAssetId: result.fields.assetId } : {}) };
          videoMainInputs.set(id, { target: at.sourceObjectId, ...(videoCandidate.movieDurationMs === undefined ? {} : { duration: videoCandidate.movieDurationMs }), eligible: !baseInput.hidden && plainVideoTimingSource(nvPr, r.p, r.a, context.r) });
          cachedPreview = true; mediaPlaceholder = true;
          reportMedia(!("reason" in result), "reason" in result, "video");
          for (const operation of ["create", "import", "edit", "preview", "export", "no-op-save", "edit-save"] as const) {
            const tier = operation === "no-op-save" ? "preserved" : operation === "edit-save" || operation === "export" ? "unsupported" : operation === "import" || operation === "preview" ? "fallback" : "editable";
            finding("drawing.video.projected-fields", operation, tier, at, "Video bytes, source identity, poster, geometry and bounded main-sequence starts are projected; codec validation, other playback settings, general export and source edit-save remain unsupported.", slide.id, id);
          }
          if ("reason" in result) finding("drawing.video.poster", "preview", "fallback", at, `No supported poster was projected: ${result.reason}`, slide.id, id);
        } else if ("reason" in result) {
          if (!media || !mediaPreviewMode || !geometry.box.width || !geometry.box.height) return unsupported(result.reason);
          input = { ...baseInput, type: "shape", shape: "rect", style: { fill: { type: "solid", color: "F2F2F2" }, stroke: { paint: { type: "solid", color: "777777" }, width: 1, dash: "dash" } }, text: { content: "Media", frame: { verticalAlign: "middle" } } };
          cachedPreview = true; mediaPlaceholder = true;
          reportMedia(false, true);
          finding("drawing.media.placeholder", "preview", "fallback", at, `Static placeholder replaces an unmapped poster: ${result.reason}`, slide.id, id);
        } else {
          input = { ...baseInput, ...result.fields };
          if (pictureSources) {
            if (pictureSources.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Picture source bindings exceed maxReportFindings.", at);
            const raw = await context.sourceElement(part, result.source.blip);
            if (raw.node.namespaceUri !== context.a || raw.node.local !== "blip" || attribute(raw.node, "embed", context.r) !== result.source.relationshipId) invalidStructure("Picture reference disagrees with its original source identity.", part);
            let hasSvgExtension = false; const rawStack = [raw.node];
            while (rawStack.length) {
              await checkpoint(); const current = rawStack.pop()!;
              if (current.namespaceUri === "http://schemas.microsoft.com/office/drawing/2016/SVG/main" && current.local === "svgBlip") hasSvgExtension = true;
              for (const child of current.children) if (child.kind === "element") rawStack.push(child);
            }
            pictureSources.push({ slideId: slide.id, elementId: id, assetId: result.fields.assetId, source: { ...at, elementIndex: raw.elementIndex, relationshipId: result.source.relationshipId }, sourceElementIndex: raw.elementIndex, relationshipId: result.source.relationshipId, mediaPart: result.source.mediaPart, hasSvgExtension, hasExternalLink: attribute(raw.node, "link", context.r) !== undefined });
          }
          const nv = r.one(node, r.p, container)!, picPr = r.one(nv, r.p, "cNvPicPr");
          if (picPr) r.known(picPr);
          if (media) {
            previewFallback = true;
            reportMedia(true);
          }
        }
      } else if (node.local === "grpSp") {
        const xfrm = r.one(properties, r.a, "xfrm")!, chOff = r.one(xfrm, r.a, "chOff"), chExt = r.one(xfrm, r.a, "chExt");
        if (!chOff || !chExt) return unsupported("Group child offset/extent is required for local coordinate mapping.");
        const childOrigin = { x: r.number(chOff, "x", -27273042329600, 27273042316900) / 12700, y: r.number(chOff, "y", -27273042329600, 27273042316900) / 12700 };
        const coordinateSize = { width: r.number(chExt, "cx", 0, 27273042316900) / 12700, height: r.number(chExt, "cy", 0, 27273042316900) / 12700 };
        if (!coordinateSize.width || !coordinateSize.height) return unsupported("Zero group child extent cannot form an invertible Core coordinate space.");
        const result: PresentationElementInput[] = [];
        for (const child of children(node)) {
          if (child.namespaceUri !== r.p || ["nvGrpSpPr", "grpSpPr"].includes(child.local)) continue;
          const value = await parse(child, childOrigin, id); if (value) result.push(value);
        }
        input = { ...baseInput, type: "group", coordinateSize, children: result };
      } else if (node.local === "sp" && r.one(properties, r.a, "custGeom")) {
        const result = contentRoundPath(node, properties, geometry, r);
        if ("reason" in result) return unsupported(result.reason);
        input = { ...baseInput, ...result.fields };
      } else {
        const localPreset = r.one(properties, r.a, "prstGeom"), localCustom = r.one(properties, r.a, "custGeom");
        if (localPreset && localCustom) invalidStructure("Conflicting drawing geometry choices.", part);
        const inheritedPreset = placeholder ? placeholders.preset(node) : undefined;
        const preset = localCustom ? undefined : localPreset ?? (inheritedPreset?.node.local === "prstGeom" ? inheritedPreset.node : undefined);
        const kind = preset ? attribute(preset, "prst") : undefined;
        if (!kind || !Object.hasOwn(shapes, kind)) return unsupported("Preset/custom geometry is outside the mapped Core presets.");
        const slant = kind === "parallelogram" ? readParallelogramSlant(preset!, r.a) : undefined;
        if (slant && !slant.supported) return unsupported("Parallelogram adjustments are outside the supported literal profile.");
        if (preset === localPreset) {
          r.attr(preset!, "prst");
          const adjustments = r.one(preset!, r.a, "avLst");
          if (adjustments) r.known(adjustments);
          if (slant?.supported && slant.guide) { r.known(slant.guide); r.attr(slant.guide, "name"); r.attr(slant.guide, "fmla"); }
        } else {
          finding("drawing.placeholder-preset", "import", "fallback", inheritedPreset!.source, slant?.supported && slant.value !== undefined ? "A matched ancestor supplies the preset and bounded literal parallelogram slant; fill, effects and text are not inherited by this projection." : "A matched ancestor supplies the preset name; its adjustments, fill, effects and text are not inherited by this projection.", slide.id, id);
          if (slant?.supported && slant.guide) {
            const path = inheritance.paths(inheritedPreset!.source.part)?.get(slant.guide);
            if (!path) throw new Error("Missing inherited parallelogram adjustment provenance.");
            finding("drawing.placeholder-slant", "import", "fallback", { ...inheritedPreset!.source, path }, "A matched ancestor supplies one bounded literal parallelogram slant; unsupported formulas and other geometry remain source-owned.", slide.id, id);
          }
        }
        const nv = r.one(node, r.p, "nvSpPr")!, shapePr = r.one(nv, r.p, "cNvSpPr");
        const textBox = shapePr ? r.bool(shapePr, "txBox") : false;
        const textBody = r.one(node, r.p, "txBody");
        const textRuns: ContentTextRun[] | undefined = bindings ? [] : undefined;
        const text = textBody ? await contentText(textBody, r, checkpoint, owner => readAction(owner, at, id), true,
          await textCascade.forDrawing(sourceSlide.layoutPart, at, r, pathByNode, placeholder, placeholders.ancestors(node)),
          textRuns ? run => {
            if (textRuns.length + (bindings?.length ?? 0) >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Text run source bindings exceed maxReportFindings.", at);
            textRuns.push(run);
          } : undefined, fieldTextProfile) : undefined;
        const directGradient = gradientReader && r.one(properties, r.a, "gradFill"), directPattern = shapePattern && gradientReader && r.one(properties, r.a, "pattFill");
        let gradientDependencies: readonly PptxSourceLocation[] | undefined;
        const fill = directGradient ? await gradientReader!.gradient(properties, at, node => { const path = pathByNode.get(node); return path ? { part, path } : undefined; }, locations => { gradientDependencies = locations; }) : directPattern ? await gradientReader!.pattern(properties, at, node => { const path = pathByNode.get(node); return path ? { part, path } : undefined; }, locations => { gradientDependencies = locations; }) : r.paint(properties), stroke = r.stroke(properties);
        if ((directGradient || directPattern) && !fill) return unsupported(directPattern ? "Shape pattern is outside the supported profile; the shape was omitted rather than painted as no-fill." : "Shape gradient is outside the supported profile; the shape was omitted rather than painted as no-fill.");
        if (gradientDependencies) finding(directPattern ? "drawing.pattern-dependencies" : "drawing.gradient-dependencies", "import", "fallback", at, "Selected fill and resolved color dependencies; this is not a complete source reference graph.", slide.id, id, gradientDependencies);
        let shadowDependencies: readonly PptxSourceLocation[] | undefined;
        const shadow = shapeShadow && gradientReader ? await gradientReader.shadow(properties, at, node => { const path = pathByNode.get(node); return path ? { part, path } : undefined; }, locations => { shadowDependencies = locations; }) : undefined;
        if (shadowDependencies) finding("drawing.shadow-dependencies", "import", "fallback", at, "Selected explicit shadow and resolved color dependencies; this is not a complete source reference graph.", slide.id, id, shadowDependencies);
        if (textBox && kind === "rect" && text && fill?.type === "none" && !stroke && !shadow) {
          input = { ...baseInput, type: "text", ...text };
        } else {
          input = { ...baseInput, type: "shape", shape: shapes[kind]!, ...(slant?.supported && slant.value !== undefined ? { slant: slant.value } : {}), style: { fill: fill ?? { type: "none" }, stroke: stroke ?? { paint: { type: "none" } }, ...(shadow ? { shadow } : {}) }, ...(text ? { text } : {}) };
        }
        if (bindings && textRuns && (input.type === "text" || includeTextOwners)) {
          if (bindings.length + textRuns.length > limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Text run source bindings exceed maxReportFindings.", at);
          const target = { slideId: slide.id, elementId: id, ...(includeTextOwners ? { owner: input.type === "text" ? "text-box" as const : "shape-text" as const } : {}) };
          for (const binding of await bindTextRuns(textRuns, target, at, context, checkpoint)) bindings.push(binding);
        }
      }
      const nv = r.one(node, r.p, container)!;
      const nvPr = r.one(nv, r.p, "nvPr"); if (nvPr) r.known(nvPr);
      const groupPr = r.one(nv, r.p, "cNvGrpSpPr"); if (groupPr) r.known(groupPr);
      const action = readAction(meta.cnv, at, id); if (action && !mediaPlaceholder) input.action = action;
      if (descriptionBindings && !cachedPreview) {
        const properties: ('description' | 'name' | 'decorative')[] = decorativeOnly ? [] : descriptionAndName ? ['description', 'name'] : [drawingName ? 'name' : 'description'];
        if (decorativeEnabled) properties.push('decorative');
        for (const property of properties) {
          if (descriptionBindings.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Drawing property source bindings exceed maxReportFindings.", at);
          const binding = property === 'decorative' ? await bindDrawingDecorative(context, part, slide.id, id, at.sourceObjectId, node, nv, meta.cnv, meta.accessibility?.description, meta.decorativeSource, meta.accessibility?.decorative) : await bindDrawingDescription(context, part, slide.id, id, at.sourceObjectId, node, nv, meta.cnv, property === 'name' ? meta.name : meta.accessibility?.description, property === 'name' ? false : descriptionAllowPresenceChange, property === 'name' ? 'name' : undefined);
          descriptionBindings.push(binding);
          const feature = property === 'decorative' ? 'drawing.decorative-existing' : property === 'name' ? 'drawing.name-existing' : descriptionAllowPresenceChange ? 'drawing.description-attribute' : 'drawing.description-existing';
          finding(feature, 'edit-save', binding.status === 'bound' ? 'editable' : 'unsupported', { ...at, ...(binding.elementIndex ? { elementIndex: binding.elementIndex } : {}) }, binding.status === 'bound' ? 'Drawing property has a direct source binding; candidate save still checks package guards, closure and full mapped state.' : binding.reason!, slide.id, id);
        }
      }
      drawings.set(meta.sourceId, { input, scope });
      owners.set(path, { id, sourceObjectId: at.sourceObjectId });
      identities.push({ kind: "element", id, source: at });
      if (meta.decorativeSource) finding("drawing.decorative", "import", "editable", { ...at, path: pathByNode.get(meta.decorativeSource)! }, "Explicit decorative metadata is projected; source edit-save and application behavior require separate verification.", slide.id, id);
      reportPlaceholder(id);
      if (!localGeometry && inheritedGeometry) finding("drawing.placeholder-dimensions", "import", "editable", at, "Offset and extent use property-level slide/layout/master provenance in placeholderInheritance. This creates local Core geometry; dynamic template binding and source edit-save are not implemented.", slide.id, id);
      if (!cachedPreview) capabilities(`drawing.${input.type}.mapped-fields`, at, "Only explicitly interpreted fields are editable in Core; this is not complete object fidelity or edit-save support. See property and inheritance findings.", true, id, previewFallback ? "fallback" : undefined);
      finding("drawing.inheritance-and-defaults", "import", "fallback", at, "Placeholder dimensions/preset and the supported frame/paragraph/list/run text properties have separate provenance. Theme fonts/colors, backgrounds, paint, effects, unsupported text properties and template rebinding remain unresolved. Core defaults/explicit neutral paint are projection values, not inferred source appearance.", slide.id, id);
      return input;
    };
    if (tree) {
      const inputs: PresentationElementInput[] = [];
      r.known(tree);
      const rootNv = r.one(tree, r.p, "nvGrpSpPr");
      if (rootNv) {
        const meta = metadata(tree, "nvGrpSpPr"); void meta;
        const group = r.one(rootNv, r.p, "cNvGrpSpPr"), nv = r.one(rootNv, r.p, "nvPr");
        if (group) r.known(group); if (nv) r.known(nv);
      }
      const rootProperties = r.one(tree, r.p, "grpSpPr");
      if (rootProperties) r.known(rootProperties);
      for (const node of children(tree)) {
        if (node.namespaceUri === r.p && ["nvGrpSpPr", "grpSpPr"].includes(node.local)) continue;
        const input = await parse(node, { x: 0, y: 0 }, slide.id);
        if (input) inputs.push(input);
      }
      for (const pending of pendingConnections) for (const connection of pending.connections) {
        await checkpoint();
        const resolved = connectorAttachment(connection, pending.scope, drawings);
        const source: PptxSourceLocation = { ...pending.source, path: pathByNode.get(connection.node)! };
        if ("endpoint" in resolved) {
          pending.input[connection.side] = resolved.endpoint;
          capabilities("drawing.connector.attachment", source, "Native cardinal site maps to a same-scope Core target identity. Target geometry takes precedence over source cached line coordinates.", true, pending.input.id);
        } else {
          r.note(connection.node, resolved.reason);
          finding("drawing.connector.attachment", "import", "fallback", source, resolved.reason, slide.id, pending.input.id);
          diagnostics.push({ code: "connector-attachment-unmapped", severity: "warning", source, slideId: slide.id, elementId: pending.input.id!, message: resolved.reason });
        }
      }
      if (audioMain) {
        const mediaInputs = inputs.filter((input): input is Extract<PresentationElementInput, { type: "audio" | "video" }> => input.type === "video" || input.type === "audio" && input.playbackStart !== "object-click");
        const compatible = mediaInputs.length === audioMain.length && mediaInputs.every((input, index) => {
          const effect = audioMain[index]!, previous = mediaInputs[index - 1];
          if (previous?.type === "audio" && effect.start === "automatic" && previous.loop) return false;
          if (input.type === "video") { const item = videoMainInputs.get(input.id!); return item && item.eligible && item.target === effect.target && item.duration === effect.durationMs; }
          const item = audioMainInputs.get(input.id!);
          return item && !item.trimUnsupported && item.playbackMapped && item.target === effect.target && item.duration === effect.durationMs;
        });
        if (compatible) for (const [index, input] of mediaInputs.entries()) input.playbackStart = audioMain[index]!.start;
      }
      for (const input of inputs) slide.addElement(input);
    }
    await r.audit(root, pathByNode.get(root)!, (path, reason) => {
      // The background-specific audit already records mapped and unsupported fields.
      if (projectedBackgroundPath && (path === projectedBackgroundPath || path.startsWith(`${projectedBackgroundPath}/`))) return;
      let ownerPath = path;
      while (ownerPath && !owners.has(ownerPath)) ownerPath = ownerPath.slice(0, ownerPath.lastIndexOf("/"));
      const owner = owners.get(ownerPath);
      capabilities("drawing.unmapped-property", { part, path, ...(owner?.sourceObjectId ? { sourceObjectId: owner.sourceObjectId } : {}) }, reason, false, owner?.id);
    }, checkpoint);
    if (sourceSlide.notesPart) {
      const noteRuns: ContentTextRun[] | undefined = bindings && includeTextOwners ? [] : undefined;
      const notes = await contentNotes(sourceSlide.notesPart, slide.id, context, slideIds, checkpoint, (feature, at, reason, mapping) => {
        if (mapping === "fallback") finding(feature, "import", "fallback", at, reason, slide.id);
        else capabilities(feature, at, reason, mapping === "editable", undefined, "unsupported");
      }, noteRuns ? run => {
        if (noteRuns.length + bindings!.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Text run source bindings exceed maxReportFindings.", { part: sourceSlide.notesPart! });
        noteRuns.push(run);
      } : undefined, notesTextProfile ? textCascade.forNotes : textFontProfile ? textCascade.forLocal : undefined, notesTextProfile, fieldTextProfile);
      if (bindings && noteRuns && notes.identity) {
        for (const binding of await bindTextRuns(noteRuns, { slideId: slide.id, owner: "notes" }, notes.identity.source, context, checkpoint)) bindings.push(binding);
      }
      if (notes.content) slide.setNotes(notes.content);
      if (notes.identity) identities.push(notes.identity);
      diagnostics.push(...notes.diagnostics);
    }
  }
  diagnostics.push({ code: "content-projection-partial", severity: "warning", source: { part: base.inspection.mainPart }, message: "Slide-local text, seven preset shapes, embedded PNG/JPEG/GIF pictures, native straight, single-elbow and bounded double-elbow connectors, direct URL/slide clicks, nested groups, direct speaker-notes text table grids/merged rich-text cells and bounded bar/line/pie/doughnut/area/scatter/bubble/combo chart data have an explicit mapped subset. Broader geometry/actions, other chart families/styles, native table styles, notes-page layout, inheritance and other source semantics remain unsupported or fallback; original-byte no-op save alone preserves the full source." });
  const textRuns: ImportedTextRunSource[] = [], notesTextRuns: ImportedNotesTextRunSource[] = [], tableTextRuns: ImportedTableTextRunSource[] = [];
  for (const binding of bindings ?? []) {
    const { baselineText: _text, namespaceUri: _ns, rawRunElementIndex: _run, rawParagraphElementIndex: _paragraph, ...source } = binding;
    if (source.owner === "notes") {
      const { owner: _owner, elementId: _element, ...run } = source;
      notesTextRuns.push(run);
    } else if (source.owner === "table-cell") {
      const { owner: _owner, ...run } = source;
      tableTextRuns.push(run);
    } else textRuns.push(source);
  }
  const chartValues = chartBindings?.map(({ baselineValue: _value, sparseProfile: _sparseProfile, v3Profile: _v3Profile, cacheLexical: _cache, cacheNamespace: _cacheNamespace, workbookLexical: _book, workbookNamespace: _bookNamespace, ...binding }) => binding);
  const labelSource = (binding: ChartLabelBinding) => ({ slideId: binding.slideId, elementId: binding.elementId, status: binding.closure.status, ...(binding.closure.reason ? { reason: binding.closure.reason } : {}), sources: binding.closure.uses.map(({ cache, cell }) => ({ cache: { part: cache.part, elementIndex: cache.elementIndex }, workbook: { part: cell.workbookPart, worksheetPart: cell.worksheetPart, cellReference: cell.reference, sourceElementIndex: cell.elementIndex, preservesWhitespace: cell.preserveSpace } })) });
  const chartLabels = chartLabelBindings?.filter(binding => binding.categoryIndex !== undefined).map(binding => ({ ...labelSource(binding), categoryIndex: binding.categoryIndex! }));
  const chartSeriesNamesReport = chartSeriesNames ? chartLabelBindings?.filter(binding => binding.seriesIndex !== undefined).map(binding => ({ ...labelSource(binding), seriesIndex: binding.seriesIndex! })) : undefined;
  const chartSeriesOrders = chartOrderBindings?.map(binding => ({ slideId: binding.slideId, elementId: binding.elementId, source: { part: binding.part }, status: binding.complete ? "bound" as const : "unmapped" as const, ...(!binding.complete ? { reason: binding.localDependencyReason ?? "Series source identity or dependency coverage is incomplete." } : {}), series: binding.complete ? [...binding.sources].sort((a, b) => a.order - b.order).map((source, seriesIndex) => ({ seriesIndex, sourceId: source.identity })) : [] }));
  return freezeReport({ ...report, ...(audio ? { audioSources: audio.result } : {}), ...(video ? { videoSources: video.result } : {}), ...(pictureSources ? { pictureSources } : {}), ...(fieldTextProfile ? { fieldTextProfile } : {}), ...(textFontProfile ? { textFontProfile } : {}), ...(notesTextProfile ? { notesTextProfile } : {}), ...(chartValues ? { chartValues } : {}), ...(chartLabels ? { chartLabels } : {}), ...(chartSeriesNamesReport ? { chartSeriesNames: chartSeriesNamesReport } : {}), ...(chartSeriesOrders ? { chartSeriesOrders } : {}), ...(diagramPreviewMode ? { diagramPreviews } : {}), inspection: { ...report.inspection, markupCompatibility: structuredClone(context.reports) }, projection: "content", ...(bindings ? { textRuns, ...(includeTextOwners ? { notesTextRuns } : {}), ...(profile === "text-runs-v3" ? { tableTextRuns } : {}) } : {}), placeholderInheritance: inheritance.report(), textStyleInheritance: textCascade.report(), tableStyleInheritance: tableCascade.report(), findings, identities, diagnostics });
}
