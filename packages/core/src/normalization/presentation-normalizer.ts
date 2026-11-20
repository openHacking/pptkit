import { bubbleDataIssue } from "../chart/bubble-data.js";
import { assertCustomDocumentProperties, isCanonicalUtcDate } from "../document/custom-document-properties.js";
import { snapshotScatterData } from "../chart/scatter-data.js";
import { compileSlideEventTimeline } from "../document/slide-event-timeline.js";
import { StyleInputTrace } from "./style-input-origins.js";
import { tableStyleCells } from "./table-style.js";
import { backgroundSource, defaultBackground, defaultTheme, recordOrigins, resolveBackground, resolveTheme, type OriginSink, type ResolvedBackground, type ResolvedTheme } from "./inheritance.js";
import { StyleTrace, selectStyle, type TextStyleOrigins, type StyleOriginMap } from "./style-origins.js";
import { childSource, presetSource, styleOrigins, traceStyles, type ElementOriginContext, type TextOriginContext } from "./text-origins.js";
import type { PresentationInheritanceSource, PresentationInheritanceTextTarget } from "../types/inheritance.js";
import {
  DEFAULT_BACKGROUND,
  DEFAULT_LAYOUT_ID,
  DEFAULT_PRESENTATION_SIZE,
  DEFAULT_TEXT_FRAME_STYLE,
  DEFAULT_TEXT_PARAGRAPH_STYLE,
  DEFAULT_TEXT_RUN_STYLE,
  DEFAULT_THEME,
} from "../constants/presentation.js";
import type { NormalizedAsset } from "../types/asset.js";
import type {
  ChartMarkerInput,
  ChartTextStyleInput,
  NormalizedChartMarkerStyle,
  NormalizedChartAxes,
  NormalizedChartTextStyle,
  NormalizedElement,
  NormalizedElementBase,
  NormalizedPlaceholderDefinition,
  NormalizedTextParagraph,
  PlaceholderDefinitionInput,
  PresentationElement,
  TableCellInput,
  TextParagraphInput,
  TextContentInput,
  ValueAxisInput,
} from "../types/element.js";
import type { Box } from "../types/geometry.js";
import type {
  NormalizedPresentation,
  NormalizedPresentationMetadata,
  NormalizedSlideLayout,
  NormalizedSlideMaster,
  PresentationData,
  PresentationSlideData,
  SlideLayoutDefinition,
} from "../types/presentation.js";
import type { NormalizedPresentationTheme } from "../types/theme.js";
import type { ColorValue, NormalizedPaint, NormalizedStrokeStyle, NormalizedTextFrameStyle, NormalizedTextParagraphStyle, NormalizedTextRunStyle, PaintInput, StrokeStyleInput, TextStylePresetResetProperty, TextStylePresetInput, ThemeColorRole } from "../types/style.js";
import { resetFrameMarginSide } from "../text/reset-frame-margin.js";
import { resetScriptFont } from "../text/reset-script-font.js";
import type { TextFontScriptCode } from "../text/font-scripts.js";
import type { Insets } from "../types/geometry.js";
import { deepClone } from "../utils/clone.js";
import { normalizeSize } from "../validation/geometry.js";
import { estimateTextHeight } from "./measure-text.js";
import {
  normalizePaint,
  normalizeShapeStyle,
  normalizeStroke,
  normalizeTableCellStyle,
  normalizeTextFrame,
  normalizeTextParagraphStyle,
  normalizeTextRunStyle,
} from "./normalize-styles.js";

function normalizeMetadata(document: PresentationData): NormalizedPresentationMetadata {
  if (document.metadata.customProperties !== undefined) assertCustomDocumentProperties(document.metadata.customProperties);
  if ([document.metadata.createdAt, document.metadata.modifiedAt].some(value => value !== undefined && !isCanonicalUtcDate(value))) throw new TypeError("Metadata dates must be canonical UTC timestamps.");
  return {
    title: document.metadata.title ?? "PPTKit Presentation",
    author: document.metadata.author ?? "PPTKit",
    ...(document.metadata.company !== undefined ? { company: document.metadata.company } : {}),
    ...(document.metadata.subject !== undefined ? { subject: document.metadata.subject } : {}),
    ...(document.metadata.description !== undefined ? { description: document.metadata.description } : {}),
    language: document.metadata.language ?? "en-US",
    keywords: [...(document.metadata.keywords ?? [])],
    revision: document.metadata.revision ?? 1,
    ...(document.metadata.createdAt === undefined ? {} : { createdAt: document.metadata.createdAt }),
    ...(document.metadata.modifiedAt === undefined ? {} : { modifiedAt: document.metadata.modifiedAt }),
    ...(document.metadata.customProperties === undefined ? {} : { customProperties: document.metadata.customProperties.map(property => ({ ...property })) }),
  };
}

export function normalizeTextContent(
  content: TextContentInput,
  paragraphFallback = DEFAULT_TEXT_PARAGRAPH_STYLE,
  runFallback = DEFAULT_TEXT_RUN_STYLE,
  context?: TextOriginContext,
): NormalizedTextParagraph[] {
  const paragraphs: TextParagraphInput[] = typeof content === "string"
    ? content.split(/\r\n|\r|\n/).map((text) => ({ runs: [{ text }] }))
    : content;
  let previousNumber: { style: string; value: number } | undefined;
  return paragraphs.map((paragraph, paragraphIndex) => {
    const prefix = `${context?.prefix ?? "content"}.${paragraphIndex}`;
    const trace = context ? new StyleTrace(childSource(context.owner, `${prefix}.style`), "defaults.text.paragraph", context.fallback?.paragraph, context.declarations?.[paragraphIndex]?.style) : undefined;
    const style = normalizeTextParagraphStyle(paragraph.style, paragraphFallback, trace);
    const declared = paragraph.style?.bullet;
    if (declared?.type === "number" && declared.startAt === "continue") {
      if (!previousNumber || style.bullet.type !== "number" || style.bullet.style !== previousNumber.style || !Number.isSafeInteger(previousNumber.value + 1)) throw new RangeError("Numbered-list continuation requires a previous explicit numbered paragraph with the same style and an available next value.");
      style.bullet.startAt = previousNumber.value + 1;
      if (trace) trace.choose("bullet.startAt", style.bullet.startAt, style.bullet.startAt, "bullet.startAt");
    }
    previousNumber = declared?.type === "number" && style.bullet.type === "number" ? { style: style.bullet.style, value: style.bullet.startAt } : undefined;
    if (context && trace) trace.emit(context.target, `${prefix}.style`, context.sink);
    return { style, runs: paragraph.runs.map((run, runIndex) => {
      const runPrefix = `${prefix}.runs.${runIndex}.style`;
      const runTrace = context ? new StyleTrace(childSource(context.owner, runPrefix), "defaults.text.run", context.fallback?.run, context.declarations?.[paragraphIndex]?.runs[runIndex]) : undefined;
      const style = normalizeTextRunStyle(run.style, runFallback, runTrace);
      if (context && runTrace) runTrace.emit(context.target, runPrefix, context.sink);
      return { text: run.text, style, ...(run.field !== undefined ? { field: deepClone(run.field) } : {}), ...(run.action !== undefined ? { action: deepClone(run.action) } : {}) };
    }) };
  });
}

function normalizeTextFallbacks(
  preset: TextStylePresetInput | undefined,
  placeholder?: NormalizedPlaceholderDefinition,
  fallbackOrigins?: TextStyleOrigins,
  owner?: PresentationInheritanceSource,
  scalarPresetMargin = false,
  nullPresetScriptMap = false,
): { frame: NormalizedTextFrameStyle; paragraph: NormalizedTextParagraphStyle; run: NormalizedTextRunStyle; origins: TextStyleOrigins | undefined } {
  const traces = traceStyles(owner, fallbackOrigins);
  const frame = normalizeTextFrame(preset?.frame, placeholder?.textStyle.frame ?? DEFAULT_TEXT_FRAME_STYLE, traces.frame);
  if (scalarPresetMargin && owner && traces.frame && typeof preset?.frame?.margin === "object") {
    for (const side of ["top", "right", "bottom", "left"] as const) {
      if (preset.frame.margin[side] !== undefined) traces.frame.origins[`margin.${side}`] = { source: childSource(owner, "frame.margin") };
    }
  }
  const run = normalizeTextRunStyle(preset?.run, placeholder?.textStyle.run ?? DEFAULT_TEXT_RUN_STYLE, traces.run);
  if (nullPresetScriptMap && owner && traces.run && preset?.run?.scriptFontFamilies) {
    for (const [code, value] of Object.entries(preset.run.scriptFontFamilies)) if (value === null) traces.run.origins[`scriptFontFamilies.${code}`] = { source: childSource(owner, "run.scriptFontFamilies"), state: "cleared" };
  }
  return {
    frame,
    paragraph: normalizeTextParagraphStyle(preset?.paragraph, placeholder?.textStyle.paragraph ?? DEFAULT_TEXT_PARAGRAPH_STYLE, traces.paragraph),
    run,
    origins: styleOrigins(traces),
  };
}

function presetWithRunReset(preset: TextStylePresetInput | undefined, reset: readonly TextStylePresetResetProperty[] | undefined): TextStylePresetInput | undefined {
  if (!preset || !reset?.length) return preset;
  const run = preset.run && { ...preset.run }, paragraph = preset.paragraph && { ...preset.paragraph }, frame = preset.frame && { ...preset.frame };
  for (const property of reset) {
    if (property.startsWith("paragraph.")) { if (paragraph) delete paragraph[property.slice(10) as keyof typeof paragraph]; }
    else if (property.startsWith("frame.margin.")) resetFrameMarginSide(frame, property.slice(13) as keyof Insets);
    else if (property.startsWith("frame.")) { if (frame) delete frame[property.slice(6) as keyof typeof frame]; }
    else if (property.startsWith("scriptFontFamilies.")) resetScriptFont(run, property.slice(19) as TextFontScriptCode);
    else if (run) delete run[property as keyof typeof run];
  }
  return { ...preset, ...(run && { run }), ...(paragraph && { paragraph }), ...(frame && { frame }) };
}

function normalizePlaceholder(input: PlaceholderDefinitionInput, presets: Readonly<Record<string, TextStylePresetInput>>, context?: { owner: PresentationInheritanceSource; target: PresentationInheritanceTextTarget; sink: OriginSink }): { value: NormalizedPlaceholderDefinition; origins: TextStyleOrigins | undefined } {
  const preset = input.textStylePreset === undefined ? undefined : presets[input.textStylePreset];
  const fallback = normalizeTextFallbacks(preset, undefined, undefined, context ? presetSource(input.textStylePreset, context.owner) : undefined);
  const traces = traceStyles(context ? childSource(context.owner, "textStyle") : undefined, fallback.origins);
  const value = {
    key: input.key,
    kind: input.kind,
    ...(input.index === undefined ? {} : { index: input.index }),
    box: deepClone(input.box),
    textStyle: {
      frame: normalizeTextFrame(input.textStyle?.frame, fallback.frame, traces.frame),
      paragraph: normalizeTextParagraphStyle(input.textStyle?.paragraph, fallback.paragraph, traces.paragraph),
      run: normalizeTextRunStyle(input.textStyle?.run, fallback.run, traces.run),
    },
  };
  if (context) for (const component of ["frame", "paragraph", "run"] as const) traces[component]?.emit(context.target, `textStyle.${component}`, context.sink);
  return { value, origins: styleOrigins(traces) };
}

function normalizeCellStyle(cell: TableCellInput, presets: Readonly<Record<string, TextStylePresetInput>>, owner?: PresentationInheritanceSource, declarations?: StyleOriginMap) {
  const preset = presets[cell.textStylePreset ?? ""];
  const inputOrigins = owner ? new StyleInputTrace(property => childSource(presetSource(cell.textStylePreset ?? "", owner), `frame.${property}`), { ...declarations }) : undefined;
  if (cell.style?.margin === undefined && preset?.frame?.margin !== undefined) inputOrigins?.margin(preset.frame.margin);
  if (cell.style?.verticalAlign === undefined && preset?.frame?.verticalAlign !== undefined) inputOrigins?.record("verticalAlign");
  const trace = owner ? new StyleTrace(childSource(owner, "style"), "defaults.table.cell", undefined, inputOrigins?.origins) : undefined;
  const value = normalizeTableCellStyle({
    ...cell.style,
    ...(cell.style?.margin === undefined && preset?.frame?.margin !== undefined ? { margin: preset.frame.margin } : {}),
    ...(cell.style?.verticalAlign === undefined && preset?.frame?.verticalAlign !== undefined ? { verticalAlign: preset.frame.verticalAlign } : {}),
  }, trace);
  return { value, trace };
}

function boxFromElement(element: PresentationElement, placeholder?: NormalizedPlaceholderDefinition): Box {
  if (element.box !== undefined) {
    return {
      ...deepClone(element.box),
      ...(element.type === "text" && element.box.height === undefined ? { height: 0 } : {}),
    };
  }
  if (placeholder !== undefined) return deepClone(placeholder.box);
  if (element.type === "connector" && !("elementId" in element.start) && !("elementId" in element.end)) {
    let left = Math.min(element.start.x, element.end.x), right = Math.max(element.start.x, element.end.x);
    let top = Math.min(element.start.y, element.end.y), bottom = Math.max(element.start.y, element.end.y);
    for (const point of element.route ?? []) {
      left = Math.min(left, point.x); right = Math.max(right, point.x);
      top = Math.min(top, point.y); bottom = Math.max(bottom, point.y);
    }
    return { x: left, y: top, width: right - left, height: bottom - top };
  }
  return { x: 0, y: 0, width: 0, height: 0 };
}

function normalizedBase(element: PresentationElement, placeholder?: NormalizedPlaceholderDefinition): NormalizedElementBase {
  return {
    id: element.id,
    name: element.name ?? element.id,
    box: boxFromElement(element, placeholder),
    transform: {
      rotation: element.transform?.rotation ?? 0,
      flipH: element.transform?.flipH ?? false,
      flipV: element.transform?.flipV ?? false,
    },
    opacity: element.opacity ?? 1,
    hidden: element.hidden ?? false,
    accessibility: {
      ...(element.accessibility?.description !== undefined ? { description: element.accessibility.description } : {}),
      decorative: element.accessibility?.decorative ?? false,
    },
    ...(element.action !== undefined ? { action: deepClone(element.action) } : {}),
    ...(element.placeholderKey !== undefined ? { placeholderKey: element.placeholderKey } : {}),
  };
}

const CHART_PALETTE_ROLES: readonly ThemeColorRole[] = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
const CHART_MARKERS = ["diamond", "square", "circle", "triangle", "star", "x", "plus", "dash"] as const;

function ensureHexColor(color: string): string {
  return color.startsWith("#") ? color : `#${color}`;
}

function resolveChartColor(color: ColorValue | undefined, theme: NormalizedPresentationTheme, index: number): string {
  if (color !== undefined) {
    if (typeof color === "string") return ensureHexColor(color);
    return ensureHexColor(theme.colors[color.theme]);
  }
  const role = CHART_PALETTE_ROLES[index % CHART_PALETTE_ROLES.length] ?? "accent1";
  return ensureHexColor(theme.colors[role]);
}

function resolveChartPaint(input: PaintInput | undefined, fallback: NormalizedPaint, theme: NormalizedPresentationTheme): NormalizedPaint {
  const paint = normalizePaint(input, fallback);
  if (paint.type === "none") return paint;
  return { ...paint, color: resolveChartColor(paint.color, theme, 0) };
}

function resolveChartStroke(input: StrokeStyleInput | undefined, fallback: NormalizedStrokeStyle, theme: NormalizedPresentationTheme): NormalizedStrokeStyle {
  const stroke = normalizeStroke(input, fallback);
  return {
    ...stroke,
    paint: stroke.paint.type === "none"
      ? stroke.paint
      : { ...stroke.paint, color: resolveChartColor(stroke.paint.color, theme, 0) },
  };
}

function resolveChartTextStyle(
  input: ChartTextStyleInput | undefined,
  theme: NormalizedPresentationTheme,
  fallback?: NormalizedChartTextStyle,
  trace?: StyleTrace,
): NormalizedChartTextStyle {
  const fontFamily = selectStyle(trace, "fontFamily", input?.fontFamily, fallback?.fontFamily ?? { theme: "body" as const });
  return {
    fontFamily: typeof fontFamily === "string" ? fontFamily : theme.fonts[fontFamily.theme],
    fontSize: selectStyle(trace, "fontSize", input?.fontSize, fallback?.fontSize ?? 9),
    bold: selectStyle(trace, "bold", input?.bold, fallback?.bold ?? false),
    italic: selectStyle(trace, "italic", input?.italic, fallback?.italic ?? false),
    color: resolveChartColor(selectStyle(trace, "color", input?.color, fallback?.color ?? "#595959"), theme, 0),
  };
}

const AXIS_STROKE: NormalizedStrokeStyle = {
  paint: { type: "solid", color: "#7F7F7F", opacity: 1 },
  width: 0.75,
  dash: "solid",
  beginArrow: "none",
  endArrow: "none",
};

const GRIDLINE_STROKE: NormalizedStrokeStyle = {
  paint: { type: "solid", color: "#D9D9D9", opacity: 1 },
  width: 0.75,
  dash: "solid",
  beginArrow: "none",
  endArrow: "none",
};

function normalizeChartMarker(
  marker: ChartMarkerInput | undefined,
  seriesColor: string,
  seriesIndex: number,
  theme: NormalizedPresentationTheme,
): false | NormalizedChartMarkerStyle {
  if (marker === false) return false;
  const explicit = marker !== undefined && marker !== "auto" ? marker : undefined;
  const shape = explicit?.shape ?? CHART_MARKERS[seriesIndex % CHART_MARKERS.length] ?? "diamond";
  return {
    shape,
    size: explicit?.size ?? 6,
    fill: resolveChartPaint(explicit?.fill, { type: "solid", color: seriesColor, opacity: 1 }, theme),
    stroke: resolveChartStroke(explicit?.stroke, { ...AXIS_STROKE, paint: { type: "solid", color: seriesColor, opacity: 1 } }, theme),
  };
}

function normalizeElement(
  element: PresentationElement,
  placeholders: ReadonlyMap<string, NormalizedPlaceholderDefinition> = new Map(),
  presets: Readonly<Record<string, TextStylePresetInput>> = {},
  theme: NormalizedPresentationTheme = DEFAULT_THEME,
  context?: ElementOriginContext,
): NormalizedElement {
  const placeholder = element.placeholderKey === undefined ? undefined : placeholders.get(element.placeholderKey);
  const base = normalizedBase(element, placeholder);
  const owner: PresentationInheritanceSource | undefined = context ? { kind: "element", id: element.id, path: context.path } : undefined;
  const target: PresentationInheritanceTextTarget = { kind: "element", id: element.id };
  const placeholderOrigins = element.placeholderKey === undefined ? undefined : context?.placeholders.get(element.placeholderKey);
  if (element.type === "text") {
    const preset = element.textStylePreset === undefined ? undefined : presets[element.textStylePreset];
    const fallbacks = normalizeTextFallbacks(presetWithRunReset(preset, element.textStylePresetReset), placeholder, placeholderOrigins, owner ? presetSource(element.textStylePreset, owner) : undefined, typeof preset?.frame?.margin === "number", preset?.run?.scriptFontFamilies === null);
    const paragraphFallback = fallbacks.paragraph;
    const runFallback = fallbacks.run;
    const content = normalizeTextContent(element.content, paragraphFallback, runFallback, context && owner ? { owner, target, sink: context.sink, prefix: "content", fallback: fallbacks.origins } : undefined);
    const frameTrace = owner ? new StyleTrace(childSource(owner, "frame"), "defaults.text.frame", fallbacks.origins?.frame) : undefined;
    const frame = normalizeTextFrame(element.frame, fallbacks.frame, frameTrace);
    if (context && frameTrace) frameTrace.emit(target, "frame", context.sink);
    const box = element.box?.height === undefined && element.box !== undefined && placeholder === undefined
      ? { ...base.box, height: estimateTextHeight(base.box.width, content, frame) }
      : base.box;
    return {
      ...base,
      box,
      type: "text",
      content,
      plainText: content.map((paragraph) => paragraph.runs.map((run) => run.text).join("")).join("\n"),
      frame,
    };
  }
  if (element.type === "image") {
    return {
      ...base,
      type: "image",
      assetId: element.assetId,
      fit: element.fit ?? "stretch",
      crop: { left: 0, top: 0, right: 0, bottom: 0, ...(element.crop ?? {}) },
    };
  }
  if (element.type === "audio") return { ...base, type: "audio", audioAssetId: element.audioAssetId, ...(element.posterAssetId === undefined ? {} : { posterAssetId: element.posterAssetId }), ...(element.playbackStart === undefined ? {} : { playbackStart: element.playbackStart }), ...(element.playbackVolume === undefined ? {} : { playbackVolume: element.playbackVolume }) };
  if (element.type === "video") return { ...base, type: "video", videoAssetId: element.videoAssetId, ...(element.posterAssetId === undefined ? {} : { posterAssetId: element.posterAssetId }) };
  if (element.type === "shape") {
    const shapeTrace = owner ? new StyleTrace(childSource(owner, "style"), "defaults.shape.style") : undefined;
    const shapeStyle = normalizeShapeStyle(element.style, shapeTrace);
    if (context && shapeTrace) shapeTrace.emit(target, "style", context.sink);
    if (element.text === undefined) return { ...base, type: "shape", shape: element.shape, ...(element.slant === undefined ? {} : { slant: element.slant }), style: shapeStyle };
    const preset = element.text.textStylePreset === undefined ? undefined : presets[element.text.textStylePreset];
    const fallbacks = normalizeTextFallbacks(presetWithRunReset(preset, element.text.textStylePresetReset), placeholder, placeholderOrigins, owner ? presetSource(element.text.textStylePreset, owner) : undefined, typeof preset?.frame?.margin === "number", preset?.run?.scriptFontFamilies === null);
    const content = normalizeTextContent(element.text.content, fallbacks.paragraph, fallbacks.run, context && owner ? { owner, target, sink: context.sink, prefix: "text.content", fallback: fallbacks.origins } : undefined);
    const frameTrace = owner ? new StyleTrace(childSource(owner, "text.frame"), "defaults.text.frame", fallbacks.origins?.frame) : undefined;
    const frame = normalizeTextFrame(element.text.frame, fallbacks.frame, frameTrace);
    if (context && frameTrace) frameTrace.emit(target, "text.frame", context.sink);
    return {
      ...base,
      type: "shape",
      shape: element.shape,
      ...(element.slant === undefined ? {} : { slant: element.slant }),
      style: shapeStyle,
      text: {
        content,
        plainText: content.map((paragraph) => paragraph.runs.map((run) => run.text).join("")).join("\n"),
        frame,
      },
    };
  }
  if (element.type === "connector") {
    return {
      ...base,
      type: "connector",
      start: deepClone(element.start),
      end: deepClone(element.end),
      route: deepClone(element.route ?? []),
      ...(element.pathMode === undefined ? {} : { pathMode: element.pathMode }),
      style: normalizeStroke(element.style),
    };
  }
  if (element.type === "group") {
    return {
      ...base,
      type: "group",
      coordinateSize: deepClone(element.coordinateSize),
      children: element.children.map((child, index) => normalizeElement(child, new Map(), presets, theme, context ? { path: `${context.path}.children.${index}`, sink: context.sink, placeholders: new Map() } : undefined)),
    };
  }
  if (element.type === "chart") {
    const textStyle = (input: ChartTextStyleInput | undefined, prefix: string, fallback?: NormalizedChartTextStyle, origins?: StyleOriginMap) => {
      const trace = owner ? new StyleTrace(childSource(owner, prefix), "defaults.chart.textStyle", origins) : undefined;
      const value = resolveChartTextStyle(input, theme, fallback, trace);
      if (context && trace) trace.emit(target, prefix, context.sink);
      return { value, origins: trace?.origins };
    };
    const chartText = textStyle(element.style?.textStyle, "style.textStyle");
    const chartTextStyle = chartText.value;
    const titleOrigins = chartText.origins ? { ...chartText.origins,
      fontSize: { source: { kind: "default" as const, path: "defaults.chart.titleStyle.fontSize" } },
      bold: { source: { kind: "default" as const, path: "defaults.chart.titleStyle.bold" } },
    } : undefined;
    if (element.chartType === "scatter" || element.chartType === "bubble") {
      const valueAxis = (input: ValueAxisInput | undefined, prefix: string) => ({
        visible: input?.visible ?? true, labels: input?.labels ?? true,
        line: resolveChartStroke(input?.line, AXIS_STROKE, theme), majorTick: input?.majorTick ?? "outside" as const,
        majorGridlines: input?.majorGridlines === false || input?.majorGridlines === undefined ? false as const : resolveChartStroke(input.majorGridlines, GRIDLINE_STROKE, theme),
        labelStyle: textStyle(input?.labelStyle, `${prefix}.labelStyle`, chartTextStyle, chartText.origins).value,
        scale: input?.scale ?? "auto" as const,
      });
      const common = {
        ...base, type: "chart" as const,
        ...(element.title !== undefined ? { title: element.title } : {}),
        titleStyle: textStyle(element.titleStyle, "titleStyle", { ...chartTextStyle, fontSize: 12, bold: true }, titleOrigins).value,
        legend: { visible: element.legend?.visible ?? true, position: element.legend?.position ?? "right", textStyle: textStyle(element.legend?.textStyle, "legend.textStyle", chartTextStyle, chartText.origins).value },
        style: { chartArea: resolveChartPaint(element.style?.chartArea, { type: "none" }, theme), plotArea: resolveChartPaint(element.style?.plotArea, { type: "none" }, theme), textStyle: chartTextStyle },
        axes: { x: valueAxis(element.axes?.x, "axes.x"), y: valueAxis(element.axes?.y, "axes.y") },
      };
      if (element.chartType === "bubble") {
        const issue = bubbleDataIssue(element.series);
        if (issue) throw new RangeError(`Invalid bubble data: ${issue}.`);
        return { ...common, chartType: "bubble", series: element.series.map((item, index) => ({
          name: item.name, xValues: [...item.xValues], yValues: [...item.yValues], sizes: [...item.sizes],
          color: resolveChartColor(item.color, theme, index),
        })) };
      }
      const data = snapshotScatterData(element.series, element.blankDisplay);
      return { ...common, chartType: "scatter", blankDisplay: data.blankDisplay,
        series: element.series.map((item, index) => {
          const color = resolveChartColor(item.color, theme, index);
          return { name: item.name, xValues: [...data.series[index]!.xValues], yValues: [...data.series[index]!.yValues], color,
            line: item.line === undefined || item.line === false ? false : resolveChartStroke(item.line, { ...AXIS_STROKE, width: 1.5, paint: { type: "solid", color, opacity: 1 } }, theme),
            marker: normalizeChartMarker(item.marker, color, index, theme) };
        }),
      };
    }
    const common = {
      ...base,
      type: "chart" as const,
      blankDisplay: element.blankDisplay ?? (element.chartType === "area" && element.grouping !== undefined && element.grouping !== "standard" ? "zero" as const : element.series.some(series => series.values.includes(null)) ? "gap" as const : "zero" as const),
      categories: [...element.categories],
      ...(element.title !== undefined ? { title: element.title } : {}),
      titleStyle: textStyle(element.titleStyle, "titleStyle", { ...chartTextStyle, fontSize: 12, bold: true }, titleOrigins).value,
      legend: {
        visible: element.legend?.visible ?? true,
        position: element.legend?.position ?? "right",
        textStyle: textStyle(element.legend?.textStyle, "legend.textStyle", chartTextStyle, chartText.origins).value,
      },
      style: {
        chartArea: resolveChartPaint(element.style?.chartArea, { type: "none" }, theme),
        plotArea: resolveChartPaint(element.style?.plotArea, { type: "none" }, theme),
        textStyle: chartTextStyle,
      },
    };
    if (element.chartType === "pie" || element.chartType === "doughnut") {
      const series = element.series.map((item, index) => ({
        name: item.name,
        values: [...item.values],
        color: resolveChartColor(item.color, theme, index),
        pointColors: element.categories.map((_, pointIndex) => resolveChartColor(item.pointColors?.[pointIndex], theme, pointIndex)),
      }));
      if (element.chartType === "doughnut") return { ...common, chartType: "doughnut", holeSize: element.holeSize ?? 50, firstSliceAngle: element.firstSliceAngle ?? 0, series };
      return { ...common, chartType: "pie", series };
    }
    const category = element.axes?.category;
    const value = element.axes?.value;
    const axes: NormalizedChartAxes = {
      category: {
        visible: category?.visible ?? true,
        labels: category?.labels ?? true,
        line: resolveChartStroke(category?.line, AXIS_STROKE, theme),
        majorTick: category?.majorTick ?? ((element.chartType === "bar" || element.chartType === "combo") ? "none" : "outside"),
        majorGridlines: category?.majorGridlines === false || category?.majorGridlines === undefined
          ? false
          : resolveChartStroke(category.majorGridlines, GRIDLINE_STROKE, theme),
        labelStyle: textStyle(category?.labelStyle, "axes.category.labelStyle", chartTextStyle, chartText.origins).value,
      },
      value: {
        visible: value?.visible ?? true,
        labels: value?.labels ?? true,
        line: resolveChartStroke(value?.line, AXIS_STROKE, theme),
        majorTick: value?.majorTick ?? "outside",
        labelStyle: textStyle(value?.labelStyle, "axes.value.labelStyle", chartTextStyle, chartText.origins).value,
        scale: value?.scale ?? "auto",
        majorGridlines: value?.majorGridlines === false
          ? false
          : resolveChartStroke(value?.majorGridlines, GRIDLINE_STROKE, theme),
      },
    };
    if (element.chartType === "combo") {
      const secondary = element.axes?.secondaryValue;
      const secondaryValue = {
        visible: secondary?.visible ?? true, labels: secondary?.labels ?? true,
        line: resolveChartStroke(secondary?.line, AXIS_STROKE, theme), majorTick: secondary?.majorTick ?? "outside" as const,
        labelStyle: textStyle(secondary?.labelStyle, "axes.secondaryValue.labelStyle", chartTextStyle, chartText.origins).value,
        scale: secondary?.scale ?? "auto" as const,
        majorGridlines: secondary?.majorGridlines === false || secondary?.majorGridlines === undefined
          ? false as const : resolveChartStroke(secondary.majorGridlines, GRIDLINE_STROKE, theme),
      };
      return { ...common, chartType: "combo", axes: { ...axes, secondaryValue }, series: element.series.map((item, index) => {
        const color = resolveChartColor(item.color, theme, index);
        const data = { name: item.name, values: [...item.values], color };
        return item.kind === "bar" ? { ...data, kind: "bar" as const } : { ...data, kind: "line" as const,
          line: resolveChartStroke(item.line, { ...AXIS_STROKE, width: 1.5, paint: { type: "solid", color, opacity: 1 } }, theme),
          marker: normalizeChartMarker(item.marker, color, index, theme),
        };
      }) };
    }
    if (element.chartType === "bar") {
      return {
        ...common,
        chartType: "bar",
        orientation: element.orientation ?? "vertical",
        grouping: element.grouping ?? "clustered",
        seriesGap: element.seriesGap ?? 0,
        categoryGap: element.categoryGap ?? 150,
        axes,
        series: element.series.map((item, index) => ({
          name: item.name,
          values: [...item.values],
          color: resolveChartColor(item.color, theme, index),
        })),
      };
    }
    if (element.chartType === "area") return {
      ...common, chartType: "area", grouping: element.grouping ?? "standard", axes,
      series: element.series.map((item, index) => ({ name: item.name, values: [...item.values], color: resolveChartColor(item.color, theme, index) })),
    };
    return {
      ...common,
      chartType: "line",
      axes,
      series: element.series.map((item, index) => {
        const color = resolveChartColor(item.color, theme, index);
        return {
          name: item.name,
          values: [...item.values],
          color,
          line: resolveChartStroke(item.line, { ...AXIS_STROKE, width: 1.5, paint: { type: "solid", color, opacity: 1 } }, theme),
          marker: normalizeChartMarker(item.marker, color, index, theme),
        };
      }),
    };
  }
  if (element.type === "table") {
    const box = base.box;
    const styled = element.style === undefined ? undefined : tableStyleCells(element, presets, owner);
    return {
      ...base,
      type: "table",
      ...(styled === undefined ? {} : { styleSource: "materialized-regions" as const }),
      ...(element.nativeFirstRowStyle === undefined ? {} : { nativeFirstRowStyle: element.nativeFirstRowStyle }),
      ...(element.nativeTableStylePreset === undefined ? {} : { nativeTableStylePreset: element.nativeTableStylePreset }),
      ...(element.nativeTableStyleFlags === undefined ? {} : { nativeTableStyleFlags: { ...element.nativeTableStyleFlags } }),
      columns: [...element.columns],
      rows: element.rows.map((row, rowIndex) => ({
        height: row.height ?? (element.rows.length === 0 ? 0 : box.height / element.rows.length),
        cells: (styled?.cells[rowIndex] ?? row.cells).map((cell, cellIndex) => {
          const prefix = `rows.${rowIndex}.cells.${cellIndex}`;
          const declarations = styled?.origins?.[rowIndex]?.[cellIndex];
          const cellOwner = owner && childSource(owner, prefix);
          const preset = presets[cell.textStylePreset ?? ""];
          const fallbacks = normalizeTextFallbacks(preset, undefined, undefined, owner && presetSource(cell.textStylePreset ?? "", owner));
          const style = normalizeCellStyle(cell, presets, cellOwner, declarations?.style);
          if (context && style.trace) style.trace.emit(target, `${prefix}.style`, context.sink);
          return {
            content: normalizeTextContent(cell.content, fallbacks.paragraph, fallbacks.run, context && owner ? { owner, target, sink: context.sink, prefix: `${prefix}.content`, fallback: fallbacks.origins, declarations: declarations?.content } : undefined),
            ...(cell.headerScope === undefined ? {} : { headerScope: cell.headerScope }),
            rowSpan: cell.rowSpan ?? 1,
            colSpan: cell.colSpan ?? 1,
            style: style.value,
          };
        }),
      })),
    };
  }
  const _exhaustive: never = element;
  throw new Error(`Unknown element type: ${JSON.stringify(_exhaustive)}`);
}

function normalizeLayout(layout: SlideLayoutDefinition, presets: Readonly<Record<string, TextStylePresetInput>>, theme: NormalizedPresentationTheme, background: ResolvedBackground, master?: NormalizedSlideMaster, context?: { sink: OriginSink; path: string }): { value: NormalizedSlideLayout; placeholderOrigins: ReadonlyMap<string, TextStyleOrigins> } {
  const resolved = layout.placeholders.map((placeholder, index) => normalizePlaceholder(placeholder, presets, context ? { owner: { kind: "placeholder", id: layout.id, key: placeholder.key, path: `${context.path}.placeholders.${index}` }, target: { kind: "placeholder", id: layout.id, key: placeholder.key }, sink: context.sink } : undefined));
  const placeholders = resolved.map(item => item.value);
  const placeholderOrigins = new Map<string, TextStyleOrigins>();
  for (const item of resolved) if (item.origins) placeholderOrigins.set(item.value.key, item.origins);
  const placeholderMap = new Map(placeholders.map((placeholder) => [placeholder.key, placeholder]));
  const value = {
    id: layout.id,
    name: layout.name,
    ...(layout.showMasterShapes === undefined ? {} : { showMasterShapes: layout.showMasterShapes }),
    ...(master ? { masterId: master.id, theme: deepClone(theme), backgroundSource: backgroundSource(background) as "layout" | "master" | "theme" } : {}),
    background: background.value,
    elements: layout.elements.map((element, index) => normalizeElement(element, placeholderMap, presets, theme, context ? { path: `${context.path}.elements.${index}`, sink: context.sink, placeholders: placeholderOrigins } : undefined)),
    placeholders,
  };
  return { value, placeholderOrigins };
}

function defaultLayout(): NormalizedSlideLayout {
  return { id: DEFAULT_LAYOUT_ID, name: "Blank", background: deepClone(DEFAULT_BACKGROUND), elements: [], placeholders: [] };
}

function normalizeSlide(
  slide: PresentationSlideData,
  layouts: ReadonlyMap<string, NormalizedSlideLayout>,
  background: ResolvedBackground,
  presets: Readonly<Record<string, TextStylePresetInput>>,
  theme: NormalizedPresentationTheme,
  context?: ElementOriginContext,
): NormalizedPresentation["slides"][number] {
  const layoutId = slide.layoutId ?? DEFAULT_LAYOUT_ID;
  const layout = layouts.get(layoutId)!;
  const placeholderMap = new Map(layout.placeholders.map((placeholder) => [placeholder.key, placeholder]));
  return {
    id: slide.id,
    layoutId,
    ...(slide.showInheritedShapes === undefined ? {} : { showInheritedShapes: slide.showInheritedShapes }),
    background: background.value,
    backgroundSource: backgroundSource(background),
    elements: slide.elements.map((element, index) => normalizeElement(element, placeholderMap, presets, theme, context ? { ...context, path: `${context.path}.elements.${index}` } : undefined)),
    notes: normalizeTextContent(slide.notes ?? "", DEFAULT_TEXT_PARAGRAPH_STYLE, DEFAULT_TEXT_RUN_STYLE, context ? { owner: { kind: "slide", id: slide.id, path: context.path }, target: { kind: "notes", id: slide.id }, sink: context.sink, prefix: "notes", fallback: undefined } : undefined),
    ...(slide.transition !== undefined ? { transition: { ...deepClone(slide.transition), onClick: slide.transition.onClick ?? true } } : {}),
    ...(slide.readingOrder?.length ? { readingOrder: deepClone(slide.readingOrder) } : {}),
    ...(slide.eventTimeline !== undefined ? { eventTimeline: { schemaVersion: "pptkit-slide-event-timeline-v1", source: deepClone(slide.eventTimeline), plan: deepClone(compileSlideEventTimeline(slide.id, slide.elements, slide.eventTimeline)) } } : {}),
    hidden: slide.hidden,
    ...(slide.section !== undefined ? { section: slide.section } : {}),
    tags: [...slide.tags],
    customData: deepClone(slide.customData),
  };
}

function normalizeAsset(asset: PresentationData["assets"][number]): NormalizedAsset {
  return {
    id: asset.id,
    kind: asset.kind,
    source: deepClone(asset.source),
    ...(asset.mimeType !== undefined ? { mimeType: asset.mimeType } : {}),
    ...(asset.width !== undefined ? { width: asset.width } : {}),
    ...(asset.height !== undefined ? { height: asset.height } : {}),
    accessibility: {
      ...(asset.accessibility?.description !== undefined ? { description: asset.accessibility.description } : {}),
      decorative: asset.accessibility?.decorative ?? false,
    },
    ...(asset.dedupeKey !== undefined ? { dedupeKey: asset.dedupeKey } : {}),
  };
}

export class PresentationNormalizer {
  normalize(document: PresentationData, sink?: OriginSink): NormalizedPresentation {
    const theme = resolveTheme(document.theme, defaultTheme(), { kind: "document", id: document.id, path: "theme" });
    recordOrigins({ kind: "document", id: document.id }, theme, undefined, sink);
    const masterThemes = new Map<string, ResolvedTheme>(), masterBackgrounds = new Map<string, ResolvedBackground>();
    const normalizedMasters: NormalizedSlideMaster[] = (document.masters ?? []).map((master, index) => {
      const resolvedTheme = resolveTheme(master.theme, theme, { kind: "master", id: master.id, path: `masters.${index}.theme` });
      const background = resolveBackground(master.background, defaultBackground(), { kind: "master", id: master.id, path: `masters.${index}.background` }, sink !== undefined);
      masterThemes.set(master.id, resolvedTheme); masterBackgrounds.set(master.id, background);
      recordOrigins({ kind: "master", id: master.id }, resolvedTheme, background, sink);
      return { id: master.id, name: master.name, theme: resolvedTheme.value, ...(master.elements === undefined ? {} : { elements: master.elements.map((element, i) => normalizeElement(element, new Map(), document.textStylePresets, resolvedTheme.value, sink ? { sink, path: `masters.${index}.elements.${i}`, placeholders: new Map() } : undefined)) }), ...(master.background !== undefined ? { background: background.value } : {}) };
    });
    const masters = new Map(normalizedMasters.map(master => [master.id, master]));
    const layoutThemes = new Map<string, ResolvedTheme>(), layoutBackgrounds = new Map<string, ResolvedBackground>();
    const layoutPlaceholders = new Map<string, ReadonlyMap<string, TextStyleOrigins>>();
    const normalizedLayouts = document.layouts.map((layout, index) => {
      const master = layout.masterId === undefined ? undefined : masters.get(layout.masterId);
      const ownerTheme = layout.masterId === undefined ? theme : masterThemes.get(layout.masterId)!;
      const background = resolveBackground(layout.background, layout.masterId === undefined ? defaultBackground() : masterBackgrounds.get(layout.masterId)!, { kind: "layout", id: layout.id, path: `layouts.${index}.background` }, sink !== undefined);
      layoutThemes.set(layout.id, ownerTheme); layoutBackgrounds.set(layout.id, background);
      recordOrigins({ kind: "layout", id: layout.id }, ownerTheme, background, sink);
      const resolved = normalizeLayout(layout, document.textStylePresets, ownerTheme.value, background, master, sink ? { sink, path: `layouts.${index}` } : undefined);
      layoutPlaceholders.set(layout.id, resolved.placeholderOrigins);
      return resolved.value;
    });
    if (!normalizedLayouts.some((layout) => layout.id === DEFAULT_LAYOUT_ID)) {
      normalizedLayouts.unshift(defaultLayout()); layoutThemes.set(DEFAULT_LAYOUT_ID, theme); layoutBackgrounds.set(DEFAULT_LAYOUT_ID, defaultBackground());
      recordOrigins({ kind: "layout", id: DEFAULT_LAYOUT_ID }, theme, layoutBackgrounds.get(DEFAULT_LAYOUT_ID), sink);
    }
    const layouts = new Map(normalizedLayouts.map((layout) => [layout.id, layout]));
    return {
      irVersion: 2,
      id: document.id,
      ...(document.firstSlideNumber === undefined ? {} : { firstSlideNumber: document.firstSlideNumber }),
      metadata: normalizeMetadata(document),
      size: normalizeSize(document.size, DEFAULT_PRESENTATION_SIZE),
      theme: theme.value,
      ...(normalizedMasters.length ? { masters: normalizedMasters } : {}),
      assets: document.assets.map(normalizeAsset),
      layouts: normalizedLayouts,
      slides: document.slides.map((slide, index) => {
        const layoutId = slide.layoutId ?? DEFAULT_LAYOUT_ID, ownerTheme = layoutThemes.get(layoutId)!;
        const background = resolveBackground(slide.background, layoutBackgrounds.get(layoutId)!, { kind: "slide", id: slide.id, path: `slides.${index}.background` }, sink !== undefined);
        recordOrigins({ kind: "slide", id: slide.id }, ownerTheme, background, sink);
        return normalizeSlide(slide, layouts, background, document.textStylePresets, ownerTheme.value, sink ? { sink, path: `slides.${index}`, placeholders: layoutPlaceholders.get(layoutId) ?? new Map() } : undefined);
      }),
    };
  }
}
