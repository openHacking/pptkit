import type { NativeInheritedRunProperty, NativeInheritedParagraphProperty, NativeInheritedFrameProperty, NativeInheritedProperty, NativeInheritedPlaceholderDefaults, NativeInheritedRuns } from "../types/native-placeholder-inheritance.js";
import type { NativeConnectorResult } from "./native-connectors.js";
import type {
  ElementAction,
  NormalizedElement,
  NormalizedPlaceholderDefinition,
  NormalizedTableCell,
  NormalizedStrokeStyle,
  NormalizedTextFrameStyle,
  NormalizedTextParagraph,
  NormalizedPresentationTheme,
  PlaceholderKind,
} from "@pptkit/core";
import { resolveTextFontSegments } from "@pptkit/core";
import type { LayoutChartElement, LayoutConnectorElement, LayoutElement, LayoutGroupElement } from "@pptkit/layout";
import type { ExportWarning } from "../types/export.js";
import { shapeFillXml, colorXml, emu, escapeXml, groupTransformXml, paintXml, paragraphAlignment, strokeXml, transformXml } from "./xml.js";

const dateTypes = { "application-default": "datetime", "month-day-year": "datetime1", "weekday-month-day-year": "datetime2", "day-month-year": "datetime3", "month-day-year-long": "datetime4", "automatic-month-day-year": "datetimeFigureOut", "day-short-month-year": "datetime5", "month-year": "datetime6", "short-month-year": "datetime7", "month-day-year-time-12": "datetime8", "month-day-year-time-12-seconds": "datetime9", "time-24": "datetime10", "time-24-seconds": "datetime11", "time-12": "datetime12", "time-12-seconds": "datetime13" } as const;

export interface ElementXmlContext {
  theme: NormalizedPresentationTheme;
  nativeInheritedPlaceholders?: NativeInheritedPlaceholderDefaults;
  nativeInheritedRuns?: NativeInheritedRuns;
  elementId?: string;
  textPath?: string;
  slideId?: string;
  layoutId?: string;
  masterId?: string;
  warnings: ExportWarning[];
  placeholders: ReadonlyMap<string, NormalizedPlaceholderDefinition>;
  placeholderIndices: ReadonlyMap<string, number>;
  nextObjectId(): number;
  objectId?(elementId: string): number;
  nativeConnector?(element: LayoutConnectorElement): NativeConnectorResult;
  imageRelationship(assetId: string): { base: string; svg?: string } | undefined;
  audioRelationship(assetId: string, posterAssetId: string | undefined, objectId: number, playbackStart: "click-sequence" | "object-click" | "automatic" | undefined, playbackVolume: "low" | "medium" | "high" | "mute" | undefined, loop: boolean | undefined, rewindAfterPlayback: boolean | undefined, trimStartMs: number | undefined, trimEndMs: number | undefined): { audio: string; media: string; poster: string };
  videoRelationship(assetId: string, posterAssetId: string | undefined, objectId: number, playbackStart: "click-sequence" | "object-click" | "automatic" | undefined, trimStartMs: number | undefined, trimEndMs: number | undefined, playbackVolume: "low" | "medium" | "high" | undefined, muted: boolean | undefined, loop: boolean | undefined, rewindAfterPlayback: boolean | undefined): { video: string; media: string; poster: string };
  actionRelationship(action: ElementAction): string | undefined;
  nextChartId(): number;
  chartRelationship(chartId: number): string;
  registerChart(chartId: number, chart: LayoutChartElement): void;
}

function actionXml(action: ElementAction | undefined, context: ElementXmlContext): string {
  if (action === undefined) return "";
  const id = context.actionRelationship(action);
  if (id === undefined) return "";
  return `<a:hlinkClick r:id="${id}"${action.type === "slide" ? ' action="ppaction://hlinksldjump"' : ""}${action.tooltip === undefined ? "" : ` tooltip="${escapeXml(action.tooltip)}"`}/>`;
}

function accessibility(element: LayoutElement): string {
  return element.accessibility.description ?? element.name;
}

function decorativeXml(element: LayoutElement): string {
  return element.accessibility.decorative ? '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>' : '';
}

function placeholderType(kind: PlaceholderKind): string {
  if (kind === "subtitle") return "subTitle";
  if (kind === "image") return "pic";
  if (kind === "slideNumber") return "sldNum";
  if (kind === "table") return "tbl";
  return kind;
}

function nonVisualProperties(element: LayoutElement, context: ElementXmlContext, kind: "shape" | "picture" | "connector" | "graphic", connections = ""): string {
  const id = context.objectId?.(element.id) ?? context.nextObjectId();
  const placeholder = element.placeholderKey === undefined ? undefined : context.placeholders.get(element.placeholderKey);
  // Layout-owned objects may use placeholder geometry/style, but remain visible
  // artwork. Only the dedicated layout placeholder owns native prompt identity.
  const placeholderXml = placeholder === undefined || context.slideId === undefined ? "" : `<p:ph type="${placeholderType(placeholder.kind)}" idx="${context.placeholderIndices.get(placeholder.key)!}"/>`;
  const cNvPr = `<p:cNvPr id="${id}" name="${escapeXml(element.name)}" descr="${escapeXml(accessibility(element))}">${actionXml(element.action, context)}${decorativeXml(element)}</p:cNvPr>`;
  if (kind === "picture") return `<p:nvPicPr>${cNvPr}<p:cNvPicPr preferRelativeResize="0"/><p:nvPr>${placeholderXml}</p:nvPr></p:nvPicPr>`;
  if (kind === "connector") return `<p:nvCxnSpPr>${cNvPr}<p:cNvCxnSpPr>${connections}</p:cNvCxnSpPr><p:nvPr>${placeholderXml}</p:nvPr></p:nvCxnSpPr>`;
  if (kind === "graphic") return `<p:nvGraphicFramePr>${cNvPr}<p:cNvGraphicFramePr/><p:nvPr>${placeholderXml}</p:nvPr></p:nvGraphicFramePr>`;
  return `<p:nvSpPr>${cNvPr}<p:cNvSpPr${element.type === "text" ? ' txBox="1"' : ""}/><p:nvPr>${placeholderXml}</p:nvPr></p:nvSpPr>`;
}

function fontTypeface(value: NormalizedTextParagraph["runs"][number]["style"]["fontFamily"]): string {
  if (typeof value === "string") return value;
  return value.theme === "heading" ? "+mj-lt" : "+mn-lt";
}

function bulletXml(bullet: NormalizedTextParagraph["style"]["bullet"]): string {
  if (bullet.type === "none") return "<a:buNone/>";
  if (bullet.type === "bullet") return `<a:buChar char="${escapeXml(bullet.character)}"/>`;
  if (!Number.isSafeInteger(bullet.startAt) || bullet.startAt < 1 || bullet.startAt > 32767) throw new RangeError("PPTX numbered-list startAt must be an integer from 1 to 32767.");
  const mapping = { arabicPeriod: "arabicPeriod", arabicParen: "arabicParenR", alphaLowerPeriod: "alphaLcPeriod", alphaUpperPeriod: "alphaUcPeriod" } as const;
  return `<a:buAutoNum type="${mapping[bullet.style]}" startAt="${bullet.startAt}"/>`;
}

function runXml(run: NormalizedTextParagraph["runs"][number], context: ElementXmlContext, opacity: number, inherited: readonly NativeInheritedRunProperty[] = []): string {
  const style = run.style;
  if (run.field) {
    let resolved = style;
    if (style.scriptFontFamilies && Object.keys(style.scriptFontFamilies).length) {
      fontMapWarning(context);
      const family = resolveTextFontSegments(run.text, style, context.theme)[0]!.family;
      resolved = { ...style, fontFamily: family, eastAsianFontFamily: family, complexScriptFontFamily: family };
    }
    const fieldType = run.field.kind === "slide-number" ? "slidenum" : dateTypes[run.field.updateFormat];
    return `<a:fld id="{${run.field.id.toUpperCase()}}" type="${fieldType}">${runPropertiesXml(resolved, context, opacity, "rPr", run.action, inherited)}<a:t>${escapeXml(run.text)}</a:t></a:fld>`;
  }
  if (style.scriptFontFamilies && Object.keys(style.scriptFontFamilies).length) {
    fontMapWarning(context);
    const segments: { text: string; family: string }[] = [];
    for (const segment of resolveTextFontSegments(run.text, style, context.theme)) {
      const previous = segments.at(-1);
      if (previous?.family === segment.family) previous.text += segment.text;
      else segments.push({ text: segment.text, family: segment.family });
    }
    const { scriptFontFamilies: _map, ...concreteStyle } = style;
    return segments.map(segment => runXml({ ...run, text: segment.text, style: { ...concreteStyle, fontFamily: segment.family, eastAsianFontFamily: segment.family, complexScriptFontFamily: segment.family } }, context, opacity)).join("");
  }
  return `<a:r>${runPropertiesXml(style, context, opacity, "rPr", run.action, inherited)}<a:t>${escapeXml(run.text)}</a:t></a:r>`;
}

function runPropertiesXml(style: NormalizedTextParagraph["runs"][number]["style"], context: ElementXmlContext, opacity: number, tag: "rPr" | "endParaRPr" | "defRPr", action?: ElementAction, inherited: readonly NativeInheritedRunProperty[] = []): string {
  const typeface = escapeXml(fontTypeface(style.fontFamily));
  // Theme roles identify a Core font family, so keep -lt even when placed in
  // another script slot; -ea/-cs would select different, unmodeled theme fonts.
  const eastAsian = style.eastAsianFontFamily === undefined ? typeface : escapeXml(fontTypeface(style.eastAsianFontFamily));
  const complexScript = style.complexScriptFontFamily === undefined ? typeface : escapeXml(fontTypeface(style.complexScriptFontFamily));
  const properties = `${inherited.includes("language") ? "" : `lang="${escapeXml(style.language)}"`}${inherited.includes("fontSize") ? "" : ` sz="${Math.round(style.fontSize * 100)}"`}${inherited.includes("bold") ? "" : ` b="${style.bold ? 1 : 0}"`}${inherited.includes("italic") ? "" : ` i="${style.italic ? 1 : 0}"`}${inherited.includes("underline") ? "" : ` u="${style.underline ? "sng" : "none"}"`}${inherited.includes("strike") ? "" : ` strike="${style.strike ? "sngStrike" : "noStrike"}"`}`;
  const fill = inherited.includes("color") ? "" : `<a:solidFill>${colorXml(style.color, opacity * (style.colorOpacity ?? 1))}</a:solidFill>`;
  const latin = inherited.includes("fontFamily") ? "" : `<a:latin typeface="${typeface}"/>`;
  const ea = inherited.includes("fontFamily") && style.eastAsianFontFamily === undefined ? "" : `<a:ea typeface="${eastAsian}"/>`;
  const cs = inherited.includes("fontFamily") && style.complexScriptFontFamily === undefined ? "" : `<a:cs typeface="${complexScript}"/>`;
  const rtl = inherited.includes("rightToLeft") || style.rightToLeft === undefined ? "" : `<a:rtl val="${style.rightToLeft ? 1 : 0}"/>`;
  return `<a:${tag} ${properties}>${fill}${latin}${ea}${cs}${actionXml(action, context)}${rtl}</a:${tag}>`;
}

function fontMapWarning(context: ElementXmlContext): void {
  const code = "script-font-map-materialized", path = context.textPath ?? "content";
  if (context.warnings.some(w => w.code === code && w.slideId === context.slideId && w.layoutId === context.layoutId && w.masterId === context.masterId && w.elementId === context.elementId && w.path === path)) return;
  context.warnings.push({ code, message: "Script font maps are materialized as current concrete text fonts. Native reusable script/language rules are not retained; text runs may split.", ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}), ...(context.elementId !== undefined ? { elementId: context.elementId } : {}), path });
}

export function textParagraphsXml(paragraphs: NormalizedTextParagraph[], context: ElementXmlContext, opacity = 1, retainParagraphFont = false): string {
  return paragraphs.map((paragraph, paragraphIndex) => {
    const style = paragraph.style;
    const inherited = context.elementId ? context.nativeInheritedRuns?.get(context.elementId)?.paragraphs[paragraphIndex] : undefined;
    if (style.tabStops?.length && style.hanging === 0 && paragraph.runs.some(run => run.text.includes("\t")) && !inherited?.includes("paragraph.hanging")) {
      context.warnings.push({ code: "tab-stops-libreoffice-zero-indent", message: "LibreOfficeDev was observed to ignore explicit tab stops when this paragraph also emits zero first-line indent; verify the target application.", ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}), ...(context.elementId !== undefined ? { elementId: context.elementId } : {}), ...(context.textPath === undefined ? {} : { path: `${context.textPath}.${paragraphIndex}.style.tabStops` }) });
    }
    const pPr = paragraphPropertiesXml(style, "", inherited);
    const runs = paragraph.runs.map((run, runIndex) => runXml(run, context, opacity, context.elementId ? context.nativeInheritedRuns?.get(context.elementId)?.runs[paragraphIndex]?.[runIndex] : undefined)).join("");
    const last = paragraph.runs.at(-1);
    let end = '<a:endParaRPr lang="en-US"/>';
    if (last && paragraph.runs.every(run => run.text.length === 0)) {
      // Office removes zero-character runs on save. The paragraph mark owns
      // the insertion style; retain the final empty run's style there as well.
      const insertionStyle = emptyRunStyle(last.style, context);
      // Actions stay on their authored runs, never on newly inserted text.
      end = runPropertiesXml(insertionStyle, context, opacity, "endParaRPr", undefined, context.elementId ? context.nativeInheritedRuns?.get(context.elementId)?.runs[paragraphIndex]?.[paragraph.runs.length - 1] : undefined);
    } else if (last && retainParagraphFont) {
      // A table paragraph mark otherwise inherits the owner's theme font even
      // when every visible run has a concrete font. Office includes that mark
      // in line metrics, so moving the table to another master moves its text.
      // Keep only the resolved font slots here; no new size or action is owned.
      const insertionStyle = emptyRunStyle(last.style, context);
      const latin = escapeXml(fontTypeface(insertionStyle.fontFamily));
      const eastAsian = escapeXml(fontTypeface(insertionStyle.eastAsianFontFamily ?? insertionStyle.fontFamily));
      const complexScript = escapeXml(fontTypeface(insertionStyle.complexScriptFontFamily ?? insertionStyle.fontFamily));
      end = `<a:endParaRPr lang="en-US"><a:latin typeface="${latin}"/><a:ea typeface="${eastAsian}"/><a:cs typeface="${complexScript}"/></a:endParaRPr>`;
    }
    return `<a:p>${pPr}${runs}${end}</a:p>`;
  }).join("");
}

function emptyRunStyle(style: NormalizedTextParagraph["runs"][number]["style"], context: ElementXmlContext): typeof style {
  if (!style.scriptFontFamilies || Object.keys(style.scriptFontFamilies).length === 0) return style;
  const family = resolveTextFontSegments("", style, context.theme)[0]!.family;
  return { ...style, fontFamily: family, eastAsianFontFamily: family, complexScriptFontFamily: family };
}

function paragraphPropertiesXml(style: NormalizedTextParagraph["style"], defaultRun = "", inherited: readonly NativeInheritedParagraphProperty[] = []): string {
  const align = inherited.includes("paragraph.align") ? "" : ` algn="${paragraphAlignment(style.align)}"`;
  const direction = style.direction === undefined || inherited.includes("paragraph.direction") ? "" : ` rtl="${style.direction === "rtl" ? 1 : 0}"`;
  const indent = inherited.includes("paragraph.indent") ? "" : ` marL="${emu(style.indent)}"`;
  const hanging = inherited.includes("paragraph.hanging") ? "" : ` indent="-${emu(style.hanging)}"`;
  const line = inherited.includes("paragraph.lineSpacing") ? "" : `<a:lnSpc><a:spcPct val="${Math.round(style.lineSpacing * 100000)}"/></a:lnSpc>`;
  const before = inherited.includes("paragraph.spaceBefore") ? "" : `<a:spcBef><a:spcPts val="${Math.round(style.spaceBefore * 100)}"/></a:spcBef>`;
  const after = inherited.includes("paragraph.spaceAfter") ? "" : `<a:spcAft><a:spcPts val="${Math.round(style.spaceAfter * 100)}"/></a:spcAft>`;
  const bullet = inherited.includes("paragraph.bullet") ? "" : bulletXml(style.bullet);
  const alignments = { left: "l", center: "ctr", right: "r", decimal: "dec" } as const;
  const tabs = style.tabStops === undefined ? "" : `<a:tabLst>${style.tabStops.map(stop => `<a:tab pos="${emu(stop.position)}" algn="${alignments[stop.align]}"/>`).join("")}</a:tabLst>`;
  return `<a:pPr${align}${direction}${indent}${hanging}>${line}${before}${after}${bullet}${tabs}${defaultRun}</a:pPr>`;
}

const textDirections = {
  horizontal: "horz", rotate90: "vert", rotate270: "vert270", eastAsianVertical: "eaVert",
  mongolianVertical: "mongolianVert", stackedVertical: "wordArtVert", stackedVerticalRtl: "wordArtVertRtl",
} satisfies Record<NonNullable<NormalizedTextFrameStyle["textDirection"]>, string>;

function bodyPropertiesXml(style: NormalizedTextFrameStyle, inherited: readonly NativeInheritedFrameProperty[] = []): string {
  const anchor = style.verticalAlign === "middle" ? "ctr" : style.verticalAlign === "bottom" ? "b" : "t";
  const fit = inherited.includes("frame.autoFit") ? "" : style.autoFit.mode === "none" ? "<a:noAutofit/>" : style.autoFit.mode === "resize" ? "<a:spAutoFit/>" : `<a:normAutofit fontScale="${Math.round(style.autoFit.fontScale * 100000)}"/>`;
  const attrs = `${inherited.includes("frame.wrap") ? "" : ` wrap="${style.wrap ? "square" : "none"}"`}${inherited.includes("frame.verticalAlign") ? "" : ` anchor="${anchor}"`}${inherited.includes("frame.margin.left") ? "" : ` lIns="${emu(style.margin.left)}"`}${inherited.includes("frame.margin.top") ? "" : ` tIns="${emu(style.margin.top)}"`}${inherited.includes("frame.margin.right") ? "" : ` rIns="${emu(style.margin.right)}"`}${inherited.includes("frame.margin.bottom") ? "" : ` bIns="${emu(style.margin.bottom)}"`}`;
  const direction = inherited.includes("frame.textDirection") || style.textDirection === undefined ? "" : ` vert="${textDirections[style.textDirection]}"`;
  return `<a:bodyPr${attrs}${direction}>${fit}</a:bodyPr>`;
}

function textBodyXml(content: NormalizedTextParagraph[], frame: NormalizedTextFrameStyle, context: ElementXmlContext, opacity: number): string {
  return `<p:txBody>${bodyPropertiesXml(frame, context.elementId ? context.nativeInheritedRuns?.get(context.elementId)?.frame : undefined)}<a:lstStyle/>${textParagraphsXml(content, context, opacity)}</p:txBody>`;
}

function textXml(element: Extract<LayoutElement, { type: "text" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  const opacity = inheritedOpacity * element.opacity;
  return `<p:sp>${nonVisualProperties(element, context, "shape")}<p:spPr>${transformXml(element.box, element.transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>${textBodyXml(element.content, element.frame, context, opacity)}</p:sp>`;
}

function imageXml(element: Extract<LayoutElement, { type: "image" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  const relationship = context.imageRelationship(element.assetId);
  if (relationship === undefined) return "";
  if (element.fit === "contain" || element.fit === "cover") {
    context.warnings.push({ code: "image-fit-unresolved", message: `Image ${element.fit} could not be resolved from source and frame dimensions; PPTX output stretches the image.`, elementId: element.id, assetId: element.assetId, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  }
  const cropEdges = [element.crop.left, element.crop.top, element.crop.right, element.crop.bottom].map(value => Math.round(value * 100000));
  if (element.fit === "crop" && (cropEdges[0]! + cropEdges[2]! >= 100000 || cropEdges[1]! + cropEdges[3]! >= 100000)) {
    context.warnings.push({ code: "image-crop-quantized-empty", message: "PPTX crop precision collapses the visible image area on at least one axis.", elementId: element.id, assetId: element.assetId, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  }
  const crop = element.fit === "crop" ? `<a:srcRect l="${cropEdges[0]}" t="${cropEdges[1]}" r="${cropEdges[2]}" b="${cropEdges[3]}"/>` : "";
  const alpha = inheritedOpacity * element.opacity < 1 ? `<a:alphaModFix amt="${Math.round(inheritedOpacity * element.opacity * 100000)}"/>` : "";
  const svg = relationship.svg ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${relationship.svg}"/></a:ext></a:extLst>` : "";
  return `<p:pic>${nonVisualProperties(element, context, "picture")}<p:blipFill><a:blip r:embed="${relationship.base}">${alpha}${svg}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${transformXml(element.box, element.transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function audioXml(element: Extract<LayoutElement, { type: "audio" }>, context: ElementXmlContext): string {
  const id = context.objectId?.(element.id) ?? context.nextObjectId();
  const relationships = context.audioRelationship(element.audioAssetId, element.posterAssetId, id, element.playbackStart, element.playbackVolume, element.loop, element.rewindAfterPlayback, element.trimStartMs, element.trimEndMs);
  context.warnings.push({ code: "audio-playback-unverified", message: "PCM WAVE bytes and a native media timing pattern were packaged; application playback and codec support are not verified.", elementId: element.id, assetId: element.audioAssetId, ...(context.slideId ? { slideId: context.slideId } : {}) });
  const props = `<p:nvPicPr><p:cNvPr id="${id}" name="${escapeXml(element.name)}" descr="${escapeXml(accessibility(element))}"><a:hlinkClick r:id="" action="ppaction://media"/>${decorativeXml(element)}</p:cNvPr><p:cNvPicPr/><p:nvPr><a:audioFile r:link="${relationships.audio}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${relationships.media}"${element.trimStartMs || element.trimEndMs ? `><p14:trim st="${element.trimStartMs ?? 0}" end="${element.trimEndMs ?? 0}"/></p14:media>` : "/>"}</p:ext></p:extLst></p:nvPr></p:nvPicPr>`;
  return `<p:pic>${props}<p:blipFill><a:blip r:embed="${relationships.poster}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${transformXml(element.box, element.transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function videoXml(element: Extract<LayoutElement, { type: "video" }>, context: ElementXmlContext): string {
  const id = context.objectId?.(element.id) ?? context.nextObjectId();
  const relationships = context.videoRelationship(element.videoAssetId, element.posterAssetId, id, element.playbackStart, element.trimStartMs, element.trimEndMs, element.playbackVolume, element.muted, element.loop, element.rewindAfterPlayback);
  context.warnings.push({ code: "video-playback-unverified", message: "MP4 framing and movie duration were inspected, but codec, track integrity and application playback are not verified.", elementId: element.id, assetId: element.videoAssetId, ...(context.slideId ? { slideId: context.slideId } : {}) });
  const props = `<p:nvPicPr><p:cNvPr id="${id}" name="${escapeXml(element.name)}" descr="${escapeXml(accessibility(element))}"><a:hlinkClick r:id="" action="ppaction://media"/>${decorativeXml(element)}</p:cNvPr><p:cNvPicPr/><p:nvPr><a:videoFile r:link="${relationships.video}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${relationships.media}"${element.trimStartMs || element.trimEndMs ? `><p14:trim st="${element.trimStartMs ?? 0}" end="${element.trimEndMs ?? 0}"/></p14:media>` : "/>"}</p:ext></p:extLst></p:nvPr></p:nvPicPr>`;
  return `<p:pic>${props}<p:blipFill><a:blip r:embed="${relationships.poster}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${transformXml(element.box, element.transform)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

function shapeGeometry(shape: Extract<LayoutElement, { type: "shape" }>["shape"]): string {
  return { rect: "rect", roundRect: "roundRect", ellipse: "ellipse", triangle: "triangle", diamond: "diamond", arrow: "rightArrow", chevron: "chevron", parallelogram: "parallelogram" }[shape];
}

function shapeXml(element: Extract<LayoutElement, { type: "shape" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  const opacity = inheritedOpacity * element.opacity;
  const text = element.text === undefined ? "" : textBodyXml(element.text.content, element.text.frame, context, opacity);
  const shadow = element.style.shadow;
  const fill = element.style.fill;
  const transparentFill = fill.type === "solid" ? fill.opacity < 1 : fill.type === "linearGradient" ? fill.stops.some(stop => stop.opacity < 1) : fill.type === "pattern" ? fill.foregroundOpacity < 1 || fill.backgroundOpacity < 1 : false;
  if (shadow && (opacity < 1 || transparentFill || (element.style.stroke.paint.type === "solid" && element.style.stroke.paint.opacity < 1))) context.warnings.push({ code: "shape-shadow-transparency-application-difference", message: "Applications may compose shadow and shape transparency differently; verify the target application.", elementId: element.id, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  const slant = element.slant === undefined ? undefined : Math.round(element.slant * 100000);
  if (slant !== undefined && slant / 100000 !== element.slant) context.warnings.push({ code: "shape-slant-quantized", message: "PPTX stores parallelogram slant in increments of 1/100000 of the shorter side.", elementId: element.id, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  const adjustments = slant === undefined ? "<a:avLst/>" : `<a:avLst><a:gd name="adj" fmla="val ${slant}"/></a:avLst>`;
  // Keep the authored effect opacity separate from shape paint opacity.
  const effect = shadow === undefined ? "" : `<a:effectLst><a:outerShdw blurRad="${emu(shadow.blur)}" dist="${emu(shadow.distance)}" dir="${Math.round(shadow.direction * 60000) % 21600000}" algn="ctr" rotWithShape="1">${colorXml(shadow.color, shadow.opacity)}</a:outerShdw></a:effectLst>`;
  return `<p:sp>${nonVisualProperties(element, context, "shape")}<p:spPr>${transformXml(element.box, element.transform)}<a:prstGeom prst="${shapeGeometry(element.shape)}">${adjustments}</a:prstGeom>${shapeFillXml(element.style.fill, opacity)}${strokeXml(element.style.stroke, opacity)}${effect}</p:spPr>${text}</p:sp>`;
}

function connectorXml(element: LayoutConnectorElement, context: ElementXmlContext, inheritedOpacity: number): string {
  if (context.nativeConnector && element.pathMode !== undefined) context.warnings.push({ code: "native-connector-fallback", message: "Round custom paths use editable custom geometry rather than a native connector preset.", elementId: element.id, ...(context.slideId ? { slideId: context.slideId } : {}), ...(context.layoutId ? { layoutId: context.layoutId } : {}), ...(context.masterId ? { masterId: context.masterId } : {}) });
  if (context.nativeConnector && element.pathMode === undefined) {
    const native = context.nativeConnector(element);
    if (native.supported) return `<p:cxnSp>${nonVisualProperties(element, context, "connector", native.connections)}<p:spPr>${transformXml(element.box, native.transform)}<a:prstGeom prst="${native.preset}">${native.adjustment === undefined ? "<a:avLst/>" : `<a:avLst><a:gd name="adj1" fmla="val ${native.adjustment}"/></a:avLst>`}</a:prstGeom>${strokeXml(element.style, inheritedOpacity * element.opacity)}</p:spPr></p:cxnSp>`;
    context.warnings.push({ code: "native-connector-fallback", message: native.reason, elementId: element.id, ...(context.slideId ? { slideId: context.slideId } : {}), ...(context.layoutId ? { layoutId: context.layoutId } : {}), ...(context.masterId ? { masterId: context.masterId } : {}) });
  }
  const points = [element.start, ...element.route, element.end];
  const width = Math.max(element.box.width, 0.001);
  const height = Math.max(element.box.height, 0.001);
  // Retain the established minimum extent for legacy custom-path output.
  // DrawingML permits zero extents; native straight output keeps them exact.
  const transformBox = {
    ...element.box,
    width,
    height,
  };
  const pathPoints = element.pathMode === "closed-round" ? points.slice(0, -1) : points;
  const path = pathPoints.map((point, index) => {
    const x = emu(point.x - element.box.x);
    const y = emu(point.y - element.box.y);
    return index === 0 ? `<a:moveTo><a:pt x="${x}" y="${y}"/></a:moveTo>` : `<a:lnTo><a:pt x="${x}" y="${y}"/></a:lnTo>`;
  }).join("") + (element.pathMode === "closed-round" ? "<a:close/>" : "");
  const geometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path w="${emu(width)}" h="${emu(height)}" fill="none">${path}</a:path></a:pathLst></a:custGeom>`;
  // Legacy path mode preserves explicit geometry as an ordinary shape.
  // Native preset/attachment support is selected independently above.
  return `<p:sp>${nonVisualProperties(element, context, "shape")}<p:spPr>${transformXml(transformBox, element.transform)}${geometry}${strokeXml(element.style, inheritedOpacity * element.opacity, "a:ln", element.pathMode !== undefined)}</p:spPr></p:sp>`;
}

function groupXml(element: LayoutGroupElement, context: ElementXmlContext, inheritedOpacity: number): string {
  const id = context.objectId?.(element.id) ?? context.nextObjectId();
  const opacity = inheritedOpacity * element.opacity;
  return `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${escapeXml(element.name)}" descr="${escapeXml(accessibility(element))}">${actionXml(element.action, context)}${decorativeXml(element)}</p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml(element.box, element.transform, element.coordinateSize.width, element.coordinateSize.height)}</p:grpSpPr>${element.children.map((child) => elementXml(child, context, opacity)).join("")}</p:grpSp>`;
}

function emptyTableCell(attributes: string, borders?: string): string {
  return `<a:tc${attributes}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></a:txBody>${borders === undefined ? "<a:tcPr/>" : `<a:tcPr>${borders}</a:tcPr>`}</a:tc>`;
}

function tableCellXml(cell: NormalizedTableCell, context: ElementXmlContext, opacity: number, attributes = "", borders?: string): string {
  const style = cell.style;
  const border = borders ?? ["lnL", "lnR", "lnT", "lnB"].map((tag) => strokeXml(style.stroke, opacity, `a:${tag}`)).join("");
  const anchor = style.verticalAlign === "middle" ? "ctr" : style.verticalAlign === "bottom" ? "b" : "t";
  // CT_TableCellProperties owns cell insets/anchor and orders borders before fill.
  // PowerPoint ignores bodyPr insets here and drops borders after a fill node.
  const direction = style.textDirection === undefined ? "" : ` vert="${textDirections[style.textDirection]}"`;
  return `<a:tc${attributes}><a:txBody><a:bodyPr/><a:lstStyle/>${textParagraphsXml(cell.content, context, opacity, true)}</a:txBody><a:tcPr marL="${emu(style.margin.left)}" marR="${emu(style.margin.right)}" marT="${emu(style.margin.top)}" marB="${emu(style.margin.bottom)}" anchor="${anchor}"${direction}>${border}${paintXml(style.fill, opacity)}</a:tcPr></a:tc>`;
}

function tableRowsXml(element: Extract<LayoutElement, { type: "table" }>, context: ElementXmlContext, opacity: number): string {
  const active = new Map<number, { remaining: number; horizontalOffset: number; columns: number }>();
  let segmented = false;
  const borders = (row: number, column: number, cell?: NormalizedTableCell): string | undefined => {
    const grid = element.borderGrid; if (!grid) return undefined;
    const width = grid.columnCount;
    const logicalSide = (styles: NormalizedStrokeStyle[]): NormalizedStrokeStyle => {
      const first = styles[0]!;
      const encoded = strokeXml(first, opacity);
      if (styles.every(style => strokeXml(style, opacity) === encoded)) return first;
      segmented = true;
      return { ...first, width: 0, paint: { type: "none" } };
    };
    // Merged anchors carry logical outer sides; continuations retain physical
    // perimeter hints. Office producers may discard continuation-only borders.
    const edges = cell && (cell.rowSpan > 1 || cell.colSpan > 1)
      ? [
        ["lnL", logicalSide(Array.from({ length: cell.rowSpan }, (_, i) => grid.vertical[(row + i) * (width + 1) + column]!))],
        ["lnR", logicalSide(Array.from({ length: cell.rowSpan }, (_, i) => grid.vertical[(row + i) * (width + 1) + column + cell.colSpan]!))],
        ["lnT", logicalSide(Array.from({ length: cell.colSpan }, (_, i) => grid.horizontal[row * width + column + i]!))],
        ["lnB", logicalSide(Array.from({ length: cell.colSpan }, (_, i) => grid.horizontal[(row + cell.rowSpan) * width + column + i]!))],
      ] as const
      : [["lnL", grid.vertical[row * (width + 1) + column]!], ["lnR", grid.vertical[row * (width + 1) + column + 1]!], ["lnT", grid.horizontal[row * width + column]!], ["lnB", grid.horizontal[(row + 1) * width + column]!]] as const;
    return edges.map(([tag, style]) => strokeXml(style, opacity, `a:${tag}`)).join("");
  };
  const result = element.rows.map((row, rowIndex) => {
    const cells: string[] = [];
    let inputIndex = 0;
    for (let column = 0; column < element.columns.length; column += 1) {
      const merge = active.get(column);
      if (merge !== undefined) {
        cells.push(emptyTableCell(` vMerge="1"${merge.horizontalOffset > 0 ? ' hMerge="1"' : merge.columns > 1 ? ` gridSpan="${merge.columns}"` : ""}`, borders(rowIndex, column)));
        merge.remaining -= 1;
        if (merge.remaining === 0) active.delete(column);
        continue;
      }
      const cell = row.cells[inputIndex++];
      if (cell === undefined) {
        cells.push(emptyTableCell("", borders(rowIndex, column)));
        continue;
      }
      const attributes = `${cell.colSpan > 1 ? ` gridSpan="${cell.colSpan}"` : ""}${cell.rowSpan > 1 ? ` rowSpan="${cell.rowSpan}"` : ""}`;
      cells.push(tableCellXml(cell, context, opacity, attributes, borders(rowIndex, column, cell)));
      const anchorColumn = column;
      for (let offset = 0; offset < cell.colSpan; offset += 1) {
        if (cell.rowSpan > 1) active.set(anchorColumn + offset, { remaining: cell.rowSpan - 1, horizontalOffset: offset, columns: cell.colSpan });
        if (offset > 0) {
          column += 1;
          cells.push(emptyTableCell(` hMerge="1"${cell.rowSpan > 1 ? ` rowSpan="${cell.rowSpan}"` : ""}`, borders(rowIndex, column)));
        }
      }
    }
    return `<a:tr h="${emu(row.height)}">${cells.join("")}</a:tr>`;
  }).join("");
  if (segmented) context.warnings.push({ code: "table-border-segmented-native", message: "Merged-cell outer sides contain unequal segments; native applications may render or rewrite these borders differently.", elementId: element.id, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  return result;
}

function tableFrameXml(element: Extract<LayoutElement, { type: "table" }>, context: ElementXmlContext, opacity: number): string {
  const flags = element.nativeTableStyleFlags;
  const styleId = element.nativeTableStylePreset === "medium-2-accent-2"
    ? "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}" : "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";
  const optionalFlags = [
    ["lastRow", flags?.lastRow], ["firstCol", flags?.firstColumn],
    ["lastCol", flags?.lastColumn], ["bandCol", flags?.bandedColumns],
  ].map(([name, value]) => value === true ? ` ${name}="1"` : "").join("");
  return `<p:graphicFrame>${nonVisualProperties(element, context, "graphic")}${transformXml(element.box, element.transform, "p")}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="${element.nativeFirstRowStyle === false ? 0 : 1}" bandRow="${flags?.bandedRows === false ? 0 : 1}"${optionalFlags}><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr><a:tblGrid>${element.columns.map((width) => `<a:gridCol w="${emu(width)}"/>`).join("")}</a:tblGrid>${tableRowsXml(element, context, opacity)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

function tableXml(element: Extract<LayoutElement, { type: "table" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  if (element.rows.some(row => row.cells.some(cell => cell.headerScope !== undefined))) context.warnings.push({ code: "table-header-semantics-not-exported", message: "PPTX does not retain the authored per-cell accessibility header roles.", elementId: element.id, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  if (element.nativeTableStylePreset !== undefined) context.warnings.push({ code: "native-table-preset-direct-formatting", message: "The native preset reference is exported with direct cell formatting that can mask its appearance; the preset definition is not embedded.", elementId: element.id, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  if (element.styleSource === "materialized-regions") context.warnings.push({ code: "table-style-materialized", message: "Authoring table regions are exported as effective direct cell formatting; native table-style rules are not retained.", elementId: element.id, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  const opacity = inheritedOpacity * element.opacity;
  if (element.transform.rotation === 0 && !element.transform.flipH && !element.transform.flipV) return tableFrameXml(element, context, opacity);
  // Office ignores graphic-frame transforms. Keep the authored drawing ID on
  // a native group; its editable table gets a fresh private physical identity.
  const rootId = context.objectId?.(element.id) ?? context.nextObjectId(), childId = context.nextObjectId();
  const child = { ...element, box: { x: 0, y: 0, width: element.box.width, height: element.box.height }, transform: { rotation: 0, flipH: false, flipV: false } };
  delete child.action;
  const childContext: ElementXmlContext = { ...context, objectId: () => childId };
  context.warnings.push({ code: "table-transform-grouped", message: "Rotated or flipped tables are exported as editable grouped tables.", elementId: element.id, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}), ...(context.layoutId !== undefined ? { layoutId: context.layoutId } : {}), ...(context.masterId !== undefined ? { masterId: context.masterId } : {}) });
  const nonVisual = `<p:nvGrpSpPr><p:cNvPr id="${rootId}" name="${escapeXml(element.name)}" descr="${escapeXml(accessibility(element))}">${actionXml(element.action, context)}${decorativeXml(element)}</p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>`;
  return `<p:grpSp>${nonVisual}<p:grpSpPr>${groupTransformXml(element.box, element.transform, element.box.width, element.box.height)}</p:grpSpPr>${tableFrameXml(child, childContext, opacity)}</p:grpSp>`;
}

function chartXml(element: Extract<LayoutElement, { type: "chart" }>, context: ElementXmlContext, inheritedOpacity: number): string {
  const chartId = context.nextChartId();
  context.registerChart(chartId, element);
  const rId = context.chartRelationship(chartId);
  return `<p:graphicFrame>${nonVisualProperties(element, context, "graphic")}${transformXml(element.box, element.transform, "p")}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></a:graphicData></a:graphic></p:graphicFrame>`;
}

export function elementXml(element: LayoutElement, context: ElementXmlContext, inheritedOpacity = 1): string {
  if (element.hidden) return "";
  context = { ...context, elementId: element.id };
  if (element.type === "text") return textXml(element, context, inheritedOpacity);
  if (element.type === "image") return imageXml(element, context, inheritedOpacity);
  if (element.type === "audio") return audioXml(element, context);
  if (element.type === "video") return videoXml(element, context);
  if (element.type === "shape") return shapeXml(element, context, inheritedOpacity);
  if (element.type === "connector") return connectorXml(element, context, inheritedOpacity);
  if (element.type === "group") return groupXml(element, context, inheritedOpacity);
  if (element.type === "table") return tableXml(element, context, inheritedOpacity);
  if (element.type === "chart") return chartXml(element, context, inheritedOpacity);
  context.warnings.push({ code: "unsupported-element", message: `Unsupported element type was omitted.`, ...(context.slideId !== undefined ? { slideId: context.slideId } : {}) });
  return "";
}

/** Emit only properties requested by an inheriting slide; unrelated defaults remain inherited. */
function nativePlaceholderDefaults(style: NormalizedTextParagraph["runs"][number]["style"], paragraph: NormalizedTextParagraph["style"], properties: ReadonlySet<NativeInheritedProperty>): string {
  const attributes = `${properties.has("language") ? ` lang="${escapeXml(style.language)}"` : ""}${properties.has("fontSize") ? ` sz="${Math.round(style.fontSize * 100)}"` : ""}${properties.has("bold") ? ` b="${style.bold ? 1 : 0}"` : ""}${properties.has("italic") ? ` i="${style.italic ? 1 : 0}"` : ""}${properties.has("underline") ? ` u="${style.underline ? "sng" : "none"}"` : ""}${properties.has("strike") ? ` strike="${style.strike ? "sngStrike" : "noStrike"}"` : ""}`;
  const fill = properties.has("color") ? `<a:solidFill>${colorXml(style.color, style.colorOpacity ?? 1)}</a:solidFill>` : "";
  const family = escapeXml(fontTypeface(style.fontFamily));
  const fonts = properties.has("fontFamily") ? `<a:latin typeface="${family}"/><a:ea typeface="${escapeXml(fontTypeface(style.eastAsianFontFamily ?? style.fontFamily))}"/><a:cs typeface="${escapeXml(fontTypeface(style.complexScriptFontFamily ?? style.fontFamily))}"/>` : "";
  const rtl = properties.has("rightToLeft") && style.rightToLeft !== undefined ? `<a:rtl val="${style.rightToLeft ? 1 : 0}"/>` : "";
  const runSelected = rtl !== "" || ["fontFamily", "color", "bold", "italic", "underline", "strike", "fontSize", "language"].some(key => properties.has(key as NativeInheritedRunProperty));
  const defaults = runSelected ? `<a:defRPr${attributes}>${fill}${fonts}${rtl}</a:defRPr>` : "";
  const all: NativeInheritedParagraphProperty[] = ["paragraph.align", "paragraph.direction", "paragraph.indent", "paragraph.hanging", "paragraph.lineSpacing", "paragraph.spaceBefore", "paragraph.spaceAfter", "paragraph.bullet"];
  const level = paragraphPropertiesXml(paragraph, defaults, all.filter(key => !properties.has(key))).replace("<a:pPr", "<a:lvl1pPr").replace("</a:pPr>", "</a:lvl1pPr>");
  return `<a:lstStyle>${level}</a:lstStyle>`;
}

export function placeholderShapeXml(placeholder: NormalizedPlaceholderDefinition, index: number, context: ElementXmlContext): string {
  if (placeholder.textStyle.run.scriptFontFamilies && Object.keys(placeholder.textStyle.run.scriptFontFamilies).length) fontMapWarning({ ...context, textPath: `placeholders.${index}.textStyle.run.scriptFontFamilies` });
  const id = context.nextObjectId();
  const type = placeholderType(placeholder.kind);
  const defaults = runPropertiesXml(emptyRunStyle(placeholder.textStyle.run, context), context, 1, "defRPr");
  const paragraph = paragraphPropertiesXml(placeholder.textStyle.paragraph, defaults);
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(placeholder.key)}"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="${type}" idx="${context.placeholderIndices.get(placeholder.key)!}"/></p:nvPr></p:nvSpPr><p:spPr>${transformXml(placeholder.box, { rotation: 0, flipH: false, flipV: false })}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody>${bodyPropertiesXml(placeholder.textStyle.frame)}${context.nativeInheritedPlaceholders?.has(placeholder.key) ? nativePlaceholderDefaults(placeholder.textStyle.run, placeholder.textStyle.paragraph, context.nativeInheritedPlaceholders.get(placeholder.key)!) : "<a:lstStyle/>"}<a:p>${paragraph}<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
}
