import { captureSlideReadingOrder } from "../accessibility/reading-order.js";
import { bubbleDataIssue } from "../chart/bubble-data.js";
import { isTextStylePresetResetProperty } from "../text/run-properties.js";
import { inspectScatterData } from "../chart/scatter-data.js";
import { isDynamicDateUpdateFormat } from "../fields/date-format.js";
import { DEFAULT_LAYOUT_ID, DEFAULT_THEME } from "../constants/presentation.js";
import { compileSlideEventTimeline } from "../document/slide-event-timeline.js";
import { validateTableStyle } from "./table-style.js";
import { tabStopErrors } from "./tab-stops.js";
import { EventTimelineError } from "../types/event-timeline.js";
import { isCanonicalBase64 } from "../document/asset-source.js";
import type { PresentationAsset } from "../types/asset.js";
import type {
  ChartTextStyleInput,
  ChartMarkerInput,
  ConnectorEndpointInput,
  NormalizedElementAccessibility,
  PresentationElement,
  TextContentInput,
  ValueAxisScaleInput,
} from "../types/element.js";
import type { PresentationDiagnostic } from "../types/diagnostic.js";
import type { Box } from "../types/geometry.js";
import type { PresentationData, PresentationSlide, PresentationSlideData, SlideLayoutDefinition } from "../types/presentation.js";
import type { ColorValue, PaintInput, StrokeStyleInput, TextFrameStyleInput, TextParagraphStyleInput, TextRunStyleInput, TextStylePresetInput } from "../types/style.js";
import { duplicates } from "./collection.js";
import { isFiniteNumber, isValidBox, isValidPoint, isValidSize } from "./geometry.js";
import { isValidShapeFill, isValidBackground, isValidColor, isValidFontFamily, isValidOpacity, isValidPaint, scriptFontFamilyErrors } from "./style.js";

const CHART_TYPES = new Set<string>(["bar", "line", "area", "pie", "doughnut", "combo"]);
const CHART_MARKERS = new Set<string>(["diamond", "square", "circle", "triangle", "star", "x", "plus", "dash"]);
const CHART_TICKS = new Set<string>(["none", "inside", "outside", "cross"]);
const CHART_LEGEND_POSITIONS = new Set<string>(["right", "bottom"]);
const CHART_ORIENTATIONS = new Set<string>(["vertical", "horizontal"]);
const CHART_GROUPINGS = new Set<string>(["clustered", "stacked", "percentStacked"]);
const PLACEHOLDER_KINDS = new Set<string>(["title", "subtitle", "body", "image", "table", "footer", "slideNumber"]);

interface Context {
  diagnostics: PresentationDiagnostic[];
  slideIds: Set<string>;
  assetKinds: Map<string, PresentationAsset["kind"]>;
  globalElementIds: Set<string>;
  textStylePresetNames: Set<string>;
}

function add(context: Context, diagnostic: Omit<PresentationDiagnostic, "severity"> & { severity?: "error" | "warning" }): void {
  context.diagnostics.push({ severity: diagnostic.severity ?? "error", ...diagnostic });
}

function validatePaint(context: Context, paint: PaintInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (!isValidPaint(paint)) add(context, { code: "invalid-paint", message: "Paint must use a valid color and opacity between 0 and 1.", path, ...ids });
}

function validateTransition(context: Context, transition: PresentationSlide["transition"], path: string, ids: Partial<PresentationDiagnostic>): void {
  if (transition === undefined) return;
  if (!transition || typeof transition !== "object" || Array.isArray(transition)) {
    add(context, { code: "invalid-slide-transition", message: "Slide transition must be an object.", path, ...ids });
    return;
  }
  if (!["fade", "push", "wipe"].includes(transition.type)) add(context, { code: "invalid-slide-transition", message: "Supported slide transitions are fade, push and wipe.", path: `${path}.type`, ...ids });
  if (!Number.isInteger(transition.durationMs) || transition.durationMs < 0 || transition.durationMs > 4294967295) {
    add(context, { code: "invalid-slide-transition-duration", message: "Fade duration must be an unsigned 32-bit integer in milliseconds.", path: `${path}.durationMs`, ...ids });
  }
  if (transition.onClick !== undefined && typeof transition.onClick !== "boolean") {
    add(context, { code: "invalid-slide-transition-advance", message: "Slide transition onClick must be boolean.", path: `${path}.onClick`, ...ids });
  }
  if (transition.afterMs !== undefined && (!Number.isInteger(transition.afterMs) || transition.afterMs < 0 || transition.afterMs > 4294967295)) {
    add(context, { code: "invalid-slide-transition-advance", message: "Slide transition afterMs must be an unsigned 32-bit integer in milliseconds.", path: `${path}.afterMs`, ...ids });
  }
  if ((transition.type === "push" || transition.type === "wipe") && !["left", "right", "up", "down"].includes(transition.direction)) {
    add(context, { code: "invalid-slide-transition-direction", message: "Push/wipe direction must be left, right, up or down.", path: `${path}.direction`, ...ids });
  }
  for (const key of Object.keys(transition)) if (!["type", "durationMs", "onClick", "afterMs", ...(transition.type === "push" || transition.type === "wipe" ? ["direction"] : [])].includes(key)) {
    add(context, { code: "invalid-slide-transition", message: `Unknown slide transition field "${key}".`, path: `${path}.${key}`, ...ids });
  }
}

function validateStroke(context: Context, stroke: StrokeStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (stroke === undefined) return;
  validatePaint(context, stroke.paint, `${path}.paint`, ids);
  if (stroke.width !== undefined && (!isFiniteNumber(stroke.width) || stroke.width < 0)) {
    add(context, { code: "invalid-stroke-width", message: "Stroke width must be a non-negative finite number.", path: `${path}.width`, ...ids });
  }
}

function validateFrame(context: Context, frame: TextFrameStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (frame === undefined) return;
  if (frame.textDirection !== undefined && !["horizontal", "rotate90", "rotate270", "eastAsianVertical", "mongolianVertical", "stackedVertical", "stackedVerticalRtl"].includes(frame.textDirection)) add(context, { code: "invalid-frame-text-direction", message: "Unsupported text frame direction.", path: `${path}.textDirection`, ...ids });
  if (typeof frame.margin === "number" && (!isFiniteNumber(frame.margin) || frame.margin < 0)) {
    add(context, { code: "invalid-text-margin", message: "Text margin must be non-negative.", path: `${path}.margin`, ...ids });
  }
  if (typeof frame.margin === "object") {
    for (const [side, value] of Object.entries(frame.margin)) {
      if (!isFiniteNumber(value) || value < 0) add(context, { code: "invalid-text-margin", message: "Text margins must be non-negative finite numbers.", path: `${path}.margin.${side}`, ...ids });
    }
  }
  if (frame.autoFit?.mode === "shrink" && frame.autoFit.fontScale !== undefined && !isValidOpacity(frame.autoFit.fontScale)) {
    add(context, { code: "invalid-font-scale", message: "fontScale must be between 0 and 1.", path: `${path}.autoFit.fontScale`, ...ids });
  }
}

function validateParagraphStyle(context: Context, style: TextParagraphStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>, allowContinuation = false): void {
  if (style === undefined) return;
  if (style.direction !== undefined && style.direction !== "ltr" && style.direction !== "rtl") add(context, { code: "invalid-paragraph-direction", message: "direction must be ltr or rtl.", path: `${path}.direction`, ...ids });
  for (const key of ["indent", "hanging", "spaceBefore", "spaceAfter"] as const) {
    const value = style[key];
    if (value !== undefined && (!isFiniteNumber(value) || value < 0)) add(context, { code: "invalid-paragraph-style", message: `${key} must be a non-negative finite number.`, path: `${path}.${key}`, ...ids });
  }
  if (style.lineSpacing !== undefined && (!isFiniteNumber(style.lineSpacing) || style.lineSpacing <= 0)) add(context, { code: "invalid-line-spacing", message: "lineSpacing must be positive.", path: `${path}.lineSpacing`, ...ids });
  if (style.tabStops !== undefined) for (const error of tabStopErrors(style.tabStops)) add(context, { code: "invalid-tab-stop", message: error.message, path: `${path}.tabStops${error.index === undefined ? "" : `.${error.index}`}`, ...ids });
  if (style.bullet?.type === "number" && style.bullet.startAt !== undefined && (style.bullet.startAt === "continue" ? !allowContinuation : !Number.isSafeInteger(style.bullet.startAt) || typeof style.bullet.startAt !== "number" || style.bullet.startAt < 1)) add(context, { code: "invalid-number-start", message: "Numbered lists require a positive safe integer or an explicit in-content continuation.", path: `${path}.bullet.startAt`, ...ids });
}

function validateRunStyle(context: Context, style: TextRunStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (style === undefined) return;
  if (style.rightToLeft !== undefined && typeof style.rightToLeft !== "boolean") add(context, { code: "invalid-run-direction", message: "rightToLeft must be boolean.", path: `${path}.rightToLeft`, ...ids });
  for (const error of scriptFontFamilyErrors(style.scriptFontFamilies)) add(context, { code: "invalid-script-font-map", message: error.message, path: `${path}.scriptFontFamilies${error.key === undefined ? "" : `.${error.key}`}`, ...ids });
  for (const field of ["eastAsianFontFamily", "complexScriptFontFamily"] as const) {
    const value = style[field];
    if (value === undefined || value === null) continue;
    if (!isValidFontFamily(value)) add(context, { code: "invalid-font-family", message: "A script font family must be a nonempty name or heading/body theme reference.", path: `${path}.${field}`, ...ids });
  }
  if (style.fontSize !== undefined && (!isFiniteNumber(style.fontSize) || style.fontSize <= 0)) add(context, { code: "invalid-font-size", message: "fontSize must be positive.", path: `${path}.fontSize`, ...ids });
  if (style.color !== undefined && !isValidColor(style.color)) add(context, { code: "invalid-color", message: "Text color must be a six-digit RGB value or theme reference.", path: `${path}.color`, ...ids });
  if (style.colorOpacity !== undefined && !isValidOpacity(style.colorOpacity)) add(context, { code: "invalid-opacity", message: "Text color opacity must be between 0 and 1.", path: `${path}.colorOpacity`, ...ids });
}

function validateChartTextStyle(context: Context, style: ChartTextStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (style === undefined) return;
  if (style.fontSize !== undefined && (!isFiniteNumber(style.fontSize) || style.fontSize <= 0)) add(context, { code: "invalid-chart-text-style", message: "Chart fontSize must be positive.", path: `${path}.fontSize`, ...ids });
  if (style.color !== undefined && !isValidColor(style.color)) add(context, { code: "invalid-chart-text-style", message: "Chart text color must be valid.", path: `${path}.color`, ...ids });
}

function validateChartAxes(context: Context, axes: Extract<PresentationElement, { type: "chart"; chartType: "bar" | "line" }>["axes"], path: string, ids: Partial<PresentationDiagnostic>, valueKey = "value"): void {
  if (axes === undefined) return;
  for (const [name, axis] of [["category", axes.category], [valueKey, axes.value]] as const) {
    if (axis === undefined) continue;
    if (axis.visible !== undefined && typeof axis.visible !== "boolean") add(context, { code: "invalid-chart-axis", message: "Axis visible must be boolean.", path: `${path}.${name}.visible`, ...ids });
    if (axis.labels !== undefined && typeof axis.labels !== "boolean") add(context, { code: "invalid-chart-axis", message: "Axis labels must be boolean.", path: `${path}.${name}.labels`, ...ids });
    if (axis.majorTick !== undefined && !CHART_TICKS.has(axis.majorTick)) add(context, { code: "invalid-chart-axis", message: "Axis majorTick must be none, inside, outside, or cross.", path: `${path}.${name}.majorTick`, ...ids });
    validateStroke(context, axis.line, `${path}.${name}.line`, ids);
    if (axis.majorGridlines !== undefined && axis.majorGridlines !== false) validateStroke(context, axis.majorGridlines, `${path}.${name}.majorGridlines`, ids);
    validateChartTextStyle(context, axis.labelStyle, `${path}.${name}.labelStyle`, ids);
  }
  const scale = axes.value?.scale;
  if (scale !== undefined && scale !== "auto") {
    if (!isFiniteNumber(scale.min) || !isFiniteNumber(scale.max) || scale.min >= scale.max || !isFiniteNumber(scale.majorUnit) || scale.majorUnit <= 0) {
      add(context, { code: "invalid-chart-scale", message: "A fixed value-axis scale requires finite min < max and majorUnit > 0.", path: `${path}.${valueKey}.scale`, ...ids });
    }
  }
}

function warnOutsideFixedAxis(context: Context, scale: ValueAxisScaleInput | undefined, value: unknown, path: string, ids: Partial<PresentationDiagnostic>, stacked = false, percentage = false): void {
  if (scale && scale !== "auto" && isFiniteNumber(scale.min) && isFiniteNumber(scale.max) && scale.min < scale.max && isFiniteNumber(scale.majorUnit) && scale.majorUnit > 0 && isFiniteNumber(value)) {
    // Accumulated normalized fractions can stray a few ULPs past an exact bound.
    const tolerance = (bound: number): number => percentage ? Math.max(1, Math.abs(bound), Math.abs(value)) * Number.EPSILON * 16 : 0;
    if (!((value < scale.min && scale.min - value > tolerance(scale.min)) || (value > scale.max && value - scale.max > tolerance(scale.max)))) return;
    add(context, { code: "chart-value-outside-fixed-axis", severity: "warning", message: stacked ? "The cumulative stack endpoint lies outside the fixed value-axis bounds and may be clipped in the chart." : "The value lies outside the fixed value-axis bounds and may be clipped in the chart.", path, ...ids });
  }
}

function warnStackedValues(context: Context, chart: Extract<PresentationElement, { type: "chart"; chartType: "bar" | "area" }>, path: string, ids: Partial<PresentationDiagnostic>): void {
  const scale = chart.axes?.value?.scale;
  if (!scale || scale === "auto" || !isFiniteNumber(scale.min) || !isFiniteNumber(scale.max) || scale.min >= scale.max || !isFiniteNumber(scale.majorUnit) || scale.majorUnit <= 0) return;
  for (let i = 0; i < chart.categories.length; i++) {
    // A malformed column cannot yield a trustworthy cumulative position.
    if (chart.series.some(series => !series || !Array.isArray(series.values) || !(i in series.values) || (series.values[i] !== null && !isFiniteNumber(series.values[i])))) continue;
    const values = chart.series.map(series => series.values[i] ?? 0);
    let maxPositive = 0, maxNegative = 0, maxAbsolute = 0;
    if (chart.grouping === "percentStacked") for (const value of values) {
      maxAbsolute = Math.max(maxAbsolute, Math.abs(value));
      if (value > 0) maxPositive = Math.max(maxPositive, value);
      if (value < 0) maxNegative = Math.max(maxNegative, -value);
    }
    let totalPositive = 0, totalNegative = 0, totalAbsolute = 0;
    if (chart.grouping === "percentStacked") for (const value of values) {
      if (value > 0) totalPositive += value / maxPositive;
      if (value < 0) totalNegative += -value / maxNegative;
      if (maxAbsolute !== 0) totalAbsolute += Math.abs(value / maxAbsolute);
    }
    let positive = 0, negative = 0, area = 0;
    const endpoints: number[] = [];
    for (const [seriesIndex, raw] of values.entries()) {
      const value = chart.grouping !== "percentStacked" ? raw : chart.chartType === "area"
        ? (totalAbsolute === 0 ? 0 : (raw / maxAbsolute) / totalAbsolute)
        : raw > 0 ? (raw / maxPositive) / totalPositive * 100 : raw < 0 ? (raw / maxNegative) / totalNegative * 100 : 0;
      const endpoint = chart.chartType === "area" ? (area += value) : value >= 0 ? (positive += value) : (negative += value);
      if (!Number.isFinite(endpoint)) break;
      endpoints[seriesIndex] = endpoint;
    }
    if (endpoints.length !== values.length) continue;
    for (const [seriesIndex, endpoint] of endpoints.entries()) if (chart.series[seriesIndex]!.values[i] !== null) warnOutsideFixedAxis(context, scale, endpoint, `${path}.series.${seriesIndex}.values.${i}`, ids, true, chart.grouping === "percentStacked");
  }
}

function validateTextStylePresetReset(context: Context, value: unknown, path: string, identity: Partial<PresentationDiagnostic>): void {
  if (value !== undefined && (!Array.isArray(value) || Array.from(value).some(item => !isTextStylePresetResetProperty(item)) || new Set(value).size !== value.length)) {
    add(context, { code: "invalid-text-style-preset-reset", message: "Preset resets must be a unique list of supported text properties.", path, ...identity });
  }
}

function validateTextStylePresetRef(context: Context, name: string | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (name !== undefined && !context.textStylePresetNames.has(name)) {
    add(context, { code: "missing-text-style-preset", message: `Unknown text style preset "${name}".`, path, ...ids });
  }
}

function validateTextStylePreset(context: Context, preset: TextStylePresetInput, path: string): void {
  validateFrame(context, preset.frame, `${path}.frame`, {});
  validateParagraphStyle(context, preset.paragraph, `${path}.paragraph`, {});
  validateRunStyle(context, preset.run, `${path}.run`, {});
}

function validateAction(context: Context, action: PresentationElement["action"], path: string, ids: Partial<PresentationDiagnostic>): void {
  if (action === undefined) return;
  if (action.type === "url" && action.url.trim() === "") add(context, { code: "invalid-action-url", message: "URL actions require a non-empty URL.", path, ...ids });
  if (action.type === "slide" && !context.slideIds.has(action.slideId)) add(context, { code: "missing-action-slide", message: `Action references missing slide "${action.slideId}".`, path, ...ids });
}

function validateText(context: Context, content: TextContentInput, path: string, ids: Partial<PresentationDiagnostic>, decorative = false): void {
  if (typeof content === "string") return;
  if (!Array.isArray(content)) {
    add(context, { code: "invalid-text-content", message: "Text content must be a string or paragraph array.", path, ...ids });
    return;
  }
  let previousNumber: { style: string; value: number } | undefined;
  content.forEach((paragraph, paragraphIndex) => {
    validateParagraphStyle(context, paragraph.style, `${path}.${paragraphIndex}.style`, ids, true);
    const bullet = paragraph.style?.bullet;
    if (bullet?.type === "number") {
      const style = bullet.style ?? "arabicPeriod";
      if (bullet.startAt === "continue") {
        if (!previousNumber || previousNumber.style !== style || !Number.isSafeInteger(previousNumber.value + 1)) {
          add(context, { code: "invalid-number-continuation", message: "Continuation requires an immediately preceding explicit numbered paragraph of the same style and an available next value.", path: `${path}.${paragraphIndex}.style.bullet.startAt`, ...ids });
          previousNumber = undefined;
        } else previousNumber = { style, value: previousNumber.value + 1 };
      } else {
        const start = bullet.startAt ?? 1;
        previousNumber = typeof start === "number" && Number.isSafeInteger(start) && start > 0 ? { style, value: start } : undefined;
      }
    } else previousNumber = undefined;
    if (!Array.isArray(paragraph.runs)) add(context, { code: "invalid-text-runs", message: "Paragraph runs must be an array.", path: `${path}.${paragraphIndex}.runs`, ...ids });
    else paragraph.runs.forEach((run, runIndex) => {
      if (typeof run.text !== "string") add(context, { code: "invalid-run-text", message: "Text run content must be a string.", path: `${path}.${paragraphIndex}.runs.${runIndex}.text`, ...ids });
      if (run.field !== undefined && (!run.field || typeof run.field.id !== "string" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(run.field.id) || (run.field.kind === "slide-number" ? Object.keys(run.field).some(key => key !== "id" && key !== "kind") : run.field.kind === "dynamic-date" ? !isDynamicDateUpdateFormat(run.field.updateFormat) || Object.keys(run.field).some(key => key !== "id" && key !== "kind" && key !== "updateFormat") : true))) add(context, { code: "invalid-run-field", message: "A dynamic text field requires a UUID and a supported slide-number or explicit date-update format.", path: `${path}.${paragraphIndex}.runs.${runIndex}.field`, ...ids });
      validateRunStyle(context, run.style, `${path}.${paragraphIndex}.runs.${runIndex}.style`, ids);
      validateAction(context, run.action, `${path}.${paragraphIndex}.runs.${runIndex}.action`, ids);
      if (decorative && run.action !== undefined) add(context, { code: "decorative-action", severity: "warning", message: "An action is inside a decorative object; its accessibility depends on the target renderer.", path: `${path}.${paragraphIndex}.runs.${runIndex}.action`, ...ids });
    });
  });
}

function collectScopeElementIds(elements: readonly PresentationElement[], destination: Set<string>): void {
  for (const element of elements) {
    destination.add(element.id);
  }
}

function endpointReference(endpoint: ConnectorEndpointInput): string | undefined {
  return "elementId" in endpoint ? endpoint.elementId : undefined;
}

function validateElement(
  context: Context,
  element: PresentationElement,
  path: string,
  scopeIds: Set<string>,
  placeholderBoxes: ReadonlyMap<string, Box>,
  ids: Partial<PresentationDiagnostic>,
  decorativeAncestor = false,
): void {
  const identity = { ...ids, elementId: element.id };
  const decorative = decorativeAncestor || element.accessibility?.decorative === true;
  if (context.globalElementIds.has(element.id)) add(context, { code: "duplicate-element-id", message: `Duplicate element id "${element.id}".`, path: `${path}.id`, ...identity });
  context.globalElementIds.add(element.id);
  if (element.box === undefined) {
    if (element.type !== "connector" && (element.placeholderKey === undefined || !placeholderBoxes.has(element.placeholderKey))) add(context, { code: "missing-element-box", message: "Elements require a box or a valid placeholder binding.", path: `${path}.box`, ...identity });
  } else if (element.type === "text" && element.box.height === undefined) {
    if (!isFiniteNumber(element.box.x) || !isFiniteNumber(element.box.y) || !isFiniteNumber(element.box.width) || element.box.width < 0) {
      add(context, { code: "invalid-text-auto-size-box", message: "Auto-sized text boxes require finite x, y, and non-negative width values.", path: `${path}.box`, ...identity });
    }
  } else if (!isValidBox(element.box)) add(context, { code: "invalid-element-box", message: "Element box must contain finite coordinates and non-negative dimensions.", path: `${path}.box`, ...identity });
  if (element.placeholderKey !== undefined && !placeholderBoxes.has(element.placeholderKey)) add(context, { code: "missing-placeholder", message: `Unknown placeholder key "${element.placeholderKey}".`, path: `${path}.placeholderKey`, ...identity });
  if (element.opacity !== undefined && !isValidOpacity(element.opacity)) add(context, { code: "invalid-opacity", message: "Element opacity must be between 0 and 1.", path: `${path}.opacity`, ...identity });
  if (element.transform?.rotation !== undefined && !isFiniteNumber(element.transform.rotation)) add(context, { code: "invalid-rotation", message: "Rotation must be finite.", path: `${path}.transform.rotation`, ...identity });
  validateAction(context, element.action, `${path}.action`, identity);
  if (decorative && element.action !== undefined) add(context, { code: "decorative-action", severity: "warning", message: "An action is on or inside a decorative object; its accessibility depends on the target renderer.", path: `${path}.action`, ...identity });

  if (element.type === "text") {
    validateText(context, element.content, `${path}.content`, identity, decorative);
    validateFrame(context, element.frame, `${path}.frame`, identity);
    validateTextStylePresetRef(context, element.textStylePreset, `${path}.textStylePreset`, identity);
    validateTextStylePresetReset(context, element.textStylePresetReset, `${path}.textStylePresetReset`, identity);
  } else if (element.type === "image") {
    if (context.assetKinds.get(element.assetId) !== "image") add(context, { code: "missing-image-asset", message: `Image references missing or non-image asset "${element.assetId}".`, path: `${path}.assetId`, assetId: element.assetId, ...identity });
    const validFit = element.fit === undefined || ["stretch", "contain", "cover", "crop"].includes(element.fit);
    if (!validFit) add(context, { code: "invalid-image-fit", message: "Image fit must be stretch, contain, cover or crop.", path: `${path}.fit`, ...identity });
    const crop = { left: 0, top: 0, right: 0, bottom: 0, ...element.crop };
    for (const [side, value] of Object.entries(crop)) if (!isValidOpacity(value)) add(context, { code: "invalid-image-crop", message: "Crop values must be between 0 and 1.", path: `${path}.crop.${side}`, ...identity });
    if (crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1) add(context, { code: "invalid-image-crop", message: "Opposing crop values must leave a visible area.", path: `${path}.crop`, ...identity });
    if (validFit && element.fit !== "crop" && [crop.left, crop.top, crop.right, crop.bottom].every(isValidOpacity) && crop.left + crop.right < 1 && crop.top + crop.bottom < 1 && [crop.left, crop.top, crop.right, crop.bottom].some(value => value > 0)) {
      add(context, { code: "image-crop-ignored", severity: "warning", message: "Authored crop fractions are used only with fit crop; other fit modes ignore or replace them.", path: `${path}.crop`, ...identity });
    }
  } else if (element.type === "audio") {
    if (element.playbackStart !== undefined && element.playbackStart !== "click-sequence" && element.playbackStart !== "object-click" && element.playbackStart !== "automatic") add(context, { code: "invalid-audio-playback-start", message: "Audio playbackStart must be click-sequence, object-click, or automatic.", path: `${path}.playbackStart`, ...identity });
    if (element.playbackVolume !== undefined && element.playbackVolume !== "low" && element.playbackVolume !== "medium" && element.playbackVolume !== "high" && element.playbackVolume !== "mute") add(context, { code: "invalid-audio-playback-volume", message: "Audio playbackVolume must be low, medium, high, or mute.", path: `${path}.playbackVolume`, ...identity });
    if (element.loop !== undefined && typeof element.loop !== "boolean") add(context, { code: "invalid-audio-loop", message: "Audio loop must be a boolean.", path: `${path}.loop`, ...identity });
    if (element.rewindAfterPlayback !== undefined && typeof element.rewindAfterPlayback !== "boolean") add(context, { code: "invalid-audio-rewind", message: "Audio rewindAfterPlayback must be a boolean.", path: `${path}.rewindAfterPlayback`, ...identity });
    for (const key of ["trimStartMs", "trimEndMs"] as const) if (element[key] !== undefined && (!Number.isSafeInteger(element[key]) || element[key]! < 0 || element[key]! > 2147483647)) add(context, { code: "invalid-audio-trim", message: `Audio ${key} must be a nonnegative integer no greater than 2147483647 milliseconds.`, path: `${path}.${key}`, ...identity });
    if (context.assetKinds.get(element.audioAssetId) !== "audio") add(context, { code: "missing-audio-asset", message: `Audio references a missing or non-audio asset "${element.audioAssetId}".`, path: `${path}.audioAssetId`, assetId: element.audioAssetId, ...identity });
    if (element.posterAssetId !== undefined && context.assetKinds.get(element.posterAssetId) !== "image") add(context, { code: "missing-audio-poster", message: `Audio poster references a missing or non-image asset "${element.posterAssetId}".`, path: `${path}.posterAssetId`, assetId: element.posterAssetId, ...identity });
  } else if (element.type === "video") {
    if (element.playbackStart !== undefined && element.playbackStart !== "click-sequence" && element.playbackStart !== "object-click" && element.playbackStart !== "automatic") add(context, { code: "invalid-video-playback-start", message: "Video playbackStart must be click-sequence, object-click, or automatic.", path: `${path}.playbackStart`, ...identity });
    if (context.assetKinds.get(element.videoAssetId) !== "video") add(context, { code: "missing-video-asset", message: `Video references a missing or non-video asset "${element.videoAssetId}".`, path: `${path}.videoAssetId`, assetId: element.videoAssetId, ...identity });
    if (element.posterAssetId !== undefined && context.assetKinds.get(element.posterAssetId) !== "image") add(context, { code: "missing-video-poster", message: `Video poster references a missing or non-image asset "${element.posterAssetId}".`, path: `${path}.posterAssetId`, assetId: element.posterAssetId, ...identity });
  } else if (element.type === "shape") {
    if (!["rect", "roundRect", "ellipse", "triangle", "diamond", "arrow", "chevron", "parallelogram"].includes(element.shape)) add(context, { code: "invalid-shape-kind", message: "Shape kind is not supported.", path: `${path}.shape`, ...identity });
    if (element.slant !== undefined && (element.shape !== "parallelogram" || !isValidOpacity(element.slant))) add(context, { code: "invalid-shape-slant", message: "Slant is available only on parallelograms and must be between 0 and 1.", path: `${path}.slant`, ...identity });
    if (!isValidShapeFill(element.style?.fill)) add(context, { code: "invalid-paint", message: "Shape fill must be a valid solid, none, linear gradient or supported pattern paint.", path: `${path}.style.fill`, ...identity });
    validateStroke(context, element.style?.stroke, `${path}.style.stroke`, identity);
    const shadow = element.style?.shadow;
    if (shadow !== undefined && (!shadow || typeof shadow !== "object" || !isValidColor(shadow.color) || (shadow.opacity !== undefined && !isValidOpacity(shadow.opacity)) || !Number.isFinite(shadow.blur) || shadow.blur < 0 || shadow.blur > 1000 || !Number.isFinite(shadow.distance) || shadow.distance < 0 || shadow.distance > 1000 || !Number.isFinite(shadow.direction) || shadow.direction < 0 || shadow.direction >= 360)) add(context, { code: "invalid-shape-shadow", message: "Shadow requires a valid color, opacity, blur and distance in 0–1000 pt, and direction in [0, 360).", path: `${path}.style.shadow`, ...identity });
    if (element.text !== undefined) {
      validateText(context, element.text.content, `${path}.text.content`, identity, decorative);
      validateFrame(context, element.text.frame, `${path}.text.frame`, identity);
      validateTextStylePresetRef(context, element.text.textStylePreset, `${path}.text.textStylePreset`, identity);
      validateTextStylePresetReset(context, element.text.textStylePresetReset, `${path}.text.textStylePresetReset`, identity);
    }
  } else if (element.type === "connector") {
    for (const [key, endpoint] of [["start", element.start], ["end", element.end]] as const) {
      const reference = endpointReference(endpoint);
      if (reference !== undefined && !scopeIds.has(reference)) add(context, { code: "missing-connector-target", message: `Connector references missing element "${reference}".`, path: `${path}.${key}`, ...identity });
      if (reference === undefined && !isValidPoint(endpoint as { x: number; y: number })) add(context, { code: "invalid-connector-point", message: "Connector points must contain finite coordinates.", path: `${path}.${key}`, ...identity });
    }
    element.route?.forEach((point, index) => { if (!isValidPoint(point)) add(context, { code: "invalid-connector-point", message: "Connector route points must contain finite coordinates.", path: `${path}.route.${index}`, ...identity }); });
    if (element.pathMode !== undefined) {
      const start = element.start, end = element.end;
      const points = start && end && !("elementId" in start) && !("elementId" in end);
      const closed = points && start.x === end.x && start.y === end.y;
      if (!(["open-round", "closed-round"] as unknown[]).includes(element.pathMode) || !points || (element.pathMode === "closed-round" && (!closed || (element.route?.length ?? 0) < 2)) || (element.style?.beginArrow !== undefined && element.style.beginArrow !== "none") || (element.style?.endArrow !== undefined && element.style.endArrow !== "none")) add(context, { code: "invalid-connector-path-mode", message: "Round custom paths require free endpoints and no arrows; closed paths end at their start and have at least two route points.", path: `${path}.pathMode`, ...identity });
    }
    validateStroke(context, element.style, `${path}.style`, identity);
  } else if (element.type === "group") {
    if (!isValidSize(element.coordinateSize) || element.coordinateSize.width <= 0 || element.coordinateSize.height <= 0) add(context, { code: "invalid-group-coordinate-size", message: "Group coordinateSize must be positive.", path: `${path}.coordinateSize`, ...identity });
    const childIds = new Set<string>();
    collectScopeElementIds(element.children, childIds);
    element.children.forEach((child, index) => validateElement(context, child, `${path}.children.${index}`, childIds, new Map(), ids, decorative));
  } else if (element.type === "chart") {
    if (element.chartType === "scatter" || element.chartType === "bubble") {
      if (element.chartType === "bubble") {
        const issue = bubbleDataIssue(element.series);
        if (issue) add(context, { code: "invalid-bubble-data", message: `Bubble data ${issue}.`, path: `${path}.series`, ...identity });
        if (Array.isArray(element.series)) for (const [i, item] of element.series.entries()) {
          if (item?.color !== undefined && !isValidColor(item.color)) add(context, { code: "invalid-chart-color", message: "Chart series color must be a valid color.", path: `${path}.series.${i}.color`, ...identity });
        }
      } else {
        for (const issue of inspectScatterData(element.series)) {
          const location = `${path}.series${issue.seriesIndex === undefined ? "" : `.${issue.seriesIndex}`}${issue.coordinate === undefined ? "" : `.${issue.coordinate}`}${issue.pointIndex === undefined ? "" : `.${issue.pointIndex}`}`;
          add(context, { code: issue.code, message: "Scatter series require paired arrays of finite coordinates or explicit null blanks.", path: location, ...identity });
        }
        if (element.blankDisplay !== undefined && !["gap", "zero", "span"].includes(element.blankDisplay)) add(context, { code: "invalid-chart-blank-display", message: "Scatter blank display must be gap, zero, or span.", path: `${path}.blankDisplay`, ...identity });
        for (const [i, series] of element.series.entries()) {
          if (!series) continue;
          if (typeof series.name !== "string" || !series.name) add(context, { code: "chart-series-name", message: "Chart series name must be a non-empty string.", path: `${path}.series.${i}.name`, ...identity });
          if (series.color !== undefined && !isValidColor(series.color)) add(context, { code: "invalid-chart-color", message: "Chart series color must be a valid color.", path: `${path}.series.${i}.color`, ...identity });
          if (series.line !== false) validateStroke(context, series.line, `${path}.series.${i}.line`, identity);
          const marker = series.marker;
          if (marker !== undefined && marker !== false && marker !== "auto") {
            if (!marker || typeof marker !== "object") add(context, { code: "invalid-chart-marker", message: "Scatter marker must be auto, false, or a marker style.", path: `${path}.series.${i}.marker`, ...identity });
            else {
              if (!CHART_MARKERS.has(marker.shape) || (marker.size !== undefined && (!isFiniteNumber(marker.size) || marker.size < 2 || marker.size > 72))) add(context, { code: "invalid-chart-marker", message: "Scatter marker shape or size is invalid.", path: `${path}.series.${i}.marker`, ...identity });
              validatePaint(context, marker.fill, `${path}.series.${i}.marker.fill`, identity); validateStroke(context, marker.stroke, `${path}.series.${i}.marker.stroke`, identity);
            }
          }
        }
      }
      for (const [name, axis] of [["x", element.axes?.x], ["y", element.axes?.y]] as const) {
        if (!axis) continue;
        const axisPath = `${path}.axes.${name}`;
        if (axis.visible !== undefined && typeof axis.visible !== "boolean") add(context, { code: "invalid-chart-axis", message: "Axis visible must be boolean.", path: `${axisPath}.visible`, ...identity });
        if (axis.labels !== undefined && typeof axis.labels !== "boolean") add(context, { code: "invalid-chart-axis", message: "Axis labels must be boolean.", path: `${axisPath}.labels`, ...identity });
        if (axis.majorTick !== undefined && !CHART_TICKS.has(axis.majorTick)) add(context, { code: "invalid-chart-axis", message: "Axis majorTick must be none, inside, outside, or cross.", path: `${axisPath}.majorTick`, ...identity });
        validateStroke(context, axis.line, `${axisPath}.line`, identity);
        if (axis.majorGridlines !== undefined && axis.majorGridlines !== false) validateStroke(context, axis.majorGridlines, `${axisPath}.majorGridlines`, identity);
        validateChartTextStyle(context, axis.labelStyle, `${axisPath}.labelStyle`, identity);
        const scale = axis.scale;
        if (scale !== undefined && scale !== "auto" && (!isFiniteNumber(scale.min) || !isFiniteNumber(scale.max) || scale.min >= scale.max || !isFiniteNumber(scale.majorUnit) || scale.majorUnit <= 0)) add(context, { code: "invalid-chart-scale", message: "A fixed value-axis scale requires finite min < max and majorUnit > 0.", path: `${axisPath}.scale`, ...identity });
      }
      if (Array.isArray(element.series)) for (const [seriesIndex, series] of element.series.entries()) {
        if (!series || !Array.isArray(series.xValues) || !Array.isArray(series.yValues) || series.xValues.length !== series.yValues.length) continue;
        for (let index = 0; index < series.xValues.length; index++) {
          const x = series.xValues[index], y = series.yValues[index];
          if (!isFiniteNumber(x) || !isFiniteNumber(y)) continue;
          if (element.chartType === "bubble" && (!("sizes" in series) || !Array.isArray(series.sizes) || series.sizes.length !== series.xValues.length || !isFiniteNumber(series.sizes[index]) || series.sizes[index]! <= 0)) continue;
          warnOutsideFixedAxis(context, element.axes?.x?.scale, x, `${path}.series.${seriesIndex}.xValues.${index}`, identity);
          warnOutsideFixedAxis(context, element.axes?.y?.scale, y, `${path}.series.${seriesIndex}.yValues.${index}`, identity);
        }
      }
      if (element.title !== undefined && element.title.length === 0) add(context, { code: "invalid-chart-title", message: "Chart title must be a non-empty string.", path: `${path}.title`, ...identity });
      if (element.legend?.visible !== undefined && typeof element.legend.visible !== "boolean") add(context, { code: "invalid-chart-legend", message: "Legend visible must be boolean.", path: `${path}.legend.visible`, ...identity });
      if (element.legend?.position !== undefined && !CHART_LEGEND_POSITIONS.has(element.legend.position)) add(context, { code: "invalid-chart-legend", message: "Legend position must be right or bottom.", path: `${path}.legend.position`, ...identity });
      validateChartTextStyle(context, element.titleStyle, `${path}.titleStyle`, identity); validateChartTextStyle(context, element.legend?.textStyle, `${path}.legend.textStyle`, identity); validateChartTextStyle(context, element.style?.textStyle, `${path}.style.textStyle`, identity);
      validatePaint(context, element.style?.chartArea, `${path}.style.chartArea`, identity); validatePaint(context, element.style?.plotArea, `${path}.style.plotArea`, identity);
    } else {
    if (!CHART_TYPES.has(element.chartType)) add(context, { code: "invalid-chart-type", message: "Unknown categorical chart type.", path: `${path}.chartType`, ...identity });
    if (element.blankDisplay !== undefined && (!(["gap", "zero", "span"] as readonly string[]).includes(element.blankDisplay) || (element.blankDisplay === "span" && element.chartType !== "line"))) add(context, { code: "invalid-chart-blank-display", message: "Blank display must be gap or zero; span is supported only for line charts.", path: `${path}.blankDisplay`, ...identity });
    if (element.categories.length === 0 || Array.from(element.categories).some((category) => typeof category !== "string" || category.length === 0)) add(context, { code: "empty-chart-categories", message: "Chart categories must be a non-empty array of non-empty strings.", path: `${path}.categories`, ...identity });
    if (element.series.length === 0) add(context, { code: "empty-chart-series", message: "Chart series must be a non-empty array.", path: `${path}.series`, ...identity });
    if (element.chartType === "combo") {
      let sawBar = false, sawLine = false, invalidOrder = false;
      for (const series of element.series) {
        if (series?.kind === "bar" && !sawLine) sawBar = true;
        else if (series?.kind === "line") sawLine = true;
        else invalidOrder = true;
      }
      if (!sawBar || !sawLine || invalidOrder) add(context, { code: "invalid-combo-series-order", message: "Combo charts require one or more bar series followed by one or more line series.", path: `${path}.series`, ...identity });
    }
    // Stacked and percent-stacked charts require aggregation, not raw values.
    for (const [seriesIndex, series] of element.series.entries()) {
      if (series === undefined) {
        add(context, { code: "invalid-chart-series", message: "Chart series entries must be present.", path: `${path}.series.${seriesIndex}`, ...identity });
        continue;
      }
      if (series.name.length === 0) add(context, { code: "chart-series-name", message: "Chart series name must be a non-empty string.", path: `${path}.series.${seriesIndex}.name`, ...identity });
      if (series.values.length !== element.categories.length) add(context, { code: "chart-values-length", message: "Chart series values length must match categories length.", path: `${path}.series.${seriesIndex}.values`, ...identity });
      const fixedScale = element.chartType === "line" || (element.chartType === "bar" && (element.grouping === undefined || element.grouping === "clustered")) || (element.chartType === "area" && (element.grouping === undefined || element.grouping === "standard")) ? element.axes?.value?.scale : element.chartType === "combo" && "kind" in series ? (series.kind === "bar" ? element.axes?.value?.scale : series.kind === "line" ? element.axes?.secondaryValue?.scale : undefined) : undefined;
      for (const [valueIndex, value] of series.values.entries()) {
        if (value !== null && !isFiniteNumber(value)) add(context, { code: "chart-values-finite", message: "Chart series values must be finite numbers or explicit null blanks.", path: `${path}.series.${seriesIndex}.values.${valueIndex}`, ...identity });
        if ((element.chartType === "pie" || element.chartType === "doughnut") && isFiniteNumber(value) && value < 0) add(context, { code: "pie-negative-value", message: "Pie and doughnut chart values must be non-negative.", path: `${path}.series.${seriesIndex}.values.${valueIndex}`, ...identity });
        warnOutsideFixedAxis(context, fixedScale, value, `${path}.series.${seriesIndex}.values.${valueIndex}`, identity);
      }
      if (series.color !== undefined && !isValidColor(series.color)) add(context, { code: "invalid-chart-color", message: "Chart series color must be a valid color.", path: `${path}.series.${seriesIndex}.color`, ...identity });
      const marker = (series as { marker?: ChartMarkerInput }).marker;
      const lineSeries = element.chartType === "line" || (element.chartType === "combo" && "kind" in series && series.kind === "line");
      if (!lineSeries && ("marker" in series || "line" in series)) add(context, { code: "invalid-chart-series-field", message: "line and marker are only valid for line charts.", path: `${path}.series.${seriesIndex}`, ...identity });
      if (lineSeries) validateStroke(context, (series as { line?: StrokeStyleInput }).line, `${path}.series.${seriesIndex}.line`, identity);
      if (lineSeries && marker !== undefined && marker !== false && marker !== "auto") {
        if (!CHART_MARKERS.has(marker.shape)) add(context, { code: "invalid-chart-marker", message: "Unknown chart marker shape.", path: `${path}.series.${seriesIndex}.marker.shape`, ...identity });
        if (marker.size !== undefined && (!isFiniteNumber(marker.size) || marker.size < 2 || marker.size > 72)) add(context, { code: "invalid-chart-marker", message: "Chart marker size must be between 2 and 72 points.", path: `${path}.series.${seriesIndex}.marker.size`, ...identity });
        validatePaint(context, marker.fill, `${path}.series.${seriesIndex}.marker.fill`, identity);
        validateStroke(context, marker.stroke, `${path}.series.${seriesIndex}.marker.stroke`, identity);
      }
      const pointColors = (series as { pointColors?: ColorValue[] }).pointColors;
      if (element.chartType !== "pie" && element.chartType !== "doughnut" && "pointColors" in series) add(context, { code: "invalid-chart-series-field", message: "pointColors is only valid for pie and doughnut charts.", path: `${path}.series.${seriesIndex}.pointColors`, ...identity });
      if ((element.chartType === "pie" || element.chartType === "doughnut") && pointColors !== undefined) {
        if (pointColors.length !== element.categories.length) add(context, { code: "chart-point-colors-length", message: "Pie pointColors length must match categories length.", path: `${path}.series.${seriesIndex}.pointColors`, ...identity });
        for (const [colorIndex, color] of pointColors.entries()) { if (!isValidColor(color)) add(context, { code: "invalid-chart-color", message: "Pie point color must be valid.", path: `${path}.series.${seriesIndex}.pointColors.${colorIndex}`, ...identity }); }
      }
    }
    if (element.chartType === "pie" && element.series.length !== 1) add(context, { code: "pie-single-series", message: "Pie charts must have exactly one series.", path: `${path}.series`, ...identity });
    if (element.title !== undefined && element.title.length === 0) add(context, { code: "invalid-chart-title", message: "Chart title must be a non-empty string.", path: `${path}.title`, ...identity });
    validateChartTextStyle(context, element.titleStyle, `${path}.titleStyle`, identity);
    validateChartTextStyle(context, element.legend?.textStyle, `${path}.legend.textStyle`, identity);
    validateChartTextStyle(context, element.style?.textStyle, `${path}.style.textStyle`, identity);
    validatePaint(context, element.style?.chartArea, `${path}.style.chartArea`, identity);
    validatePaint(context, element.style?.plotArea, `${path}.style.plotArea`, identity);
    if (element.legend?.visible !== undefined && typeof element.legend.visible !== "boolean") add(context, { code: "invalid-chart-legend", message: "Legend visible must be boolean.", path: `${path}.legend.visible`, ...identity });
    if (element.legend?.position !== undefined && !CHART_LEGEND_POSITIONS.has(element.legend.position)) add(context, { code: "invalid-chart-legend", message: "Legend position must be right or bottom.", path: `${path}.legend.position`, ...identity });
    if (element.chartType === "pie" || element.chartType === "doughnut") {
      if ((element as { axes?: unknown }).axes !== undefined) add(context, { code: "invalid-chart-axis", message: "Pie and doughnut charts do not support axes.", path: `${path}.axes`, ...identity });
    } else {
      validateChartAxes(context, element.axes, `${path}.axes`, identity);
      if (element.chartType === "combo" && element.axes?.secondaryValue !== undefined) {
        validateChartAxes(context, { value: element.axes.secondaryValue }, `${path}.axes`, identity, "secondaryValue");
      }
    }
    if (element.chartType !== "doughnut" && ((element as { holeSize?: unknown }).holeSize !== undefined || (element as { firstSliceAngle?: unknown }).firstSliceAngle !== undefined)) add(context, { code: "invalid-chart-geometry", message: "holeSize and firstSliceAngle are supported only for doughnut charts.", path, ...identity });
    if ((element.chartType === "bar" || element.chartType === "area") && (element.grouping === "stacked" || element.grouping === "percentStacked") && (element.chartType !== "area" || element.blankDisplay === undefined || element.blankDisplay === "zero")) warnStackedValues(context, element, path, identity);
    if (element.chartType === "bar" && element.grouping === "stacked") for (let i = 0; i < element.categories.length; i++) {
      let positive = 0, negative = 0;
      for (const series of element.series) {
        const value = series?.values[i] ?? 0;
        if (!isFiniteNumber(value)) continue;
        if (value >= 0) positive += value; else negative += value;
        if (!Number.isFinite(positive) || !Number.isFinite(negative)) {
          add(context, { code: "bar-stack-overflow", message: "Positive and negative bar stack cumulative values must remain finite.", path: `${path}.series`, ...identity });
          break;
        }
      }
    }
    if (element.chartType === "area") {
      if (element.grouping !== undefined && !["standard", "stacked", "percentStacked"].includes(element.grouping)) add(context, { code: "invalid-chart-grouping", message: "Area grouping must be standard, stacked, or percentStacked.", path: `${path}.grouping`, ...identity });
      if (element.grouping !== undefined && element.grouping !== "standard" && element.blankDisplay !== undefined && element.blankDisplay !== "zero") add(context, { code: "invalid-chart-blank-display", message: "Stacked area charts support zero blank display only.", path: `${path}.blankDisplay`, ...identity });
      if (element.grouping === "stacked") for (let i = 0; i < element.categories.length; i++) {
        let sum = 0; for (const series of element.series) { sum += series?.values[i] ?? 0; if (!Number.isFinite(sum)) { add(context, { code: "area-stack-overflow", message: "Area stack cumulative values must remain finite.", path: `${path}.series`, ...identity }); break; } }
      }
    }
    if (element.chartType === "doughnut") {
      if (element.holeSize !== undefined && (!Number.isInteger(element.holeSize) || element.holeSize < 1 || element.holeSize > 90)) add(context, { code: "invalid-chart-hole-size", message: "Doughnut holeSize must be an integer from 1 to 90.", path: `${path}.holeSize`, ...identity });
      if (element.firstSliceAngle !== undefined && (!Number.isInteger(element.firstSliceAngle) || element.firstSliceAngle < 0 || element.firstSliceAngle > 360)) add(context, { code: "invalid-chart-first-slice-angle", message: "Doughnut firstSliceAngle must be an integer from 0 to 360.", path: `${path}.firstSliceAngle`, ...identity });
    }
    if (element.chartType === "bar") {
      if (element.orientation !== undefined && !CHART_ORIENTATIONS.has(element.orientation)) add(context, { code: "invalid-chart-orientation", message: "Bar orientation must be vertical or horizontal.", path: `${path}.orientation`, ...identity });
      if (element.grouping !== undefined && !CHART_GROUPINGS.has(element.grouping)) add(context, { code: "invalid-chart-grouping", message: "Bar grouping must be clustered, stacked, or percentStacked.", path: `${path}.grouping`, ...identity });
      if (element.seriesGap !== undefined && (!isFiniteNumber(element.seriesGap) || element.seriesGap < 0 || element.seriesGap > 100)) add(context, { code: "invalid-chart-gap", message: "seriesGap must be between 0 and 100 percent.", path: `${path}.seriesGap`, ...identity });
      if (element.categoryGap !== undefined && (!isFiniteNumber(element.categoryGap) || element.categoryGap < 0 || element.categoryGap > 500)) add(context, { code: "invalid-chart-gap", message: "categoryGap must be between 0 and 500 percent.", path: `${path}.categoryGap`, ...identity });
    }
    }
  } else if (element.type === "table") {
    if (element.nativeTableStylePreset !== undefined && element.nativeTableStylePreset !== "medium-2-accent-1" && element.nativeTableStylePreset !== "medium-2-accent-2") {
      add(context, { code: "invalid-native-table-style-preset", message: "Unsupported native table style preset.", path: `${path}.nativeTableStylePreset`, ...identity });
    }
    if (element.nativeTableStyleFlags !== undefined) {
      const flags = element.nativeTableStyleFlags;
      const allowed = ["firstColumn", "lastColumn", "lastRow", "bandedRows", "bandedColumns"];
      if (!flags || typeof flags !== "object" || Array.isArray(flags) || Object.entries(flags).some(([key, value]) => !allowed.includes(key) || typeof value !== "boolean")) {
        add(context, { code: "invalid-native-table-style-flags", message: "Native table style flags must be an object of supported booleans.", path: `${path}.nativeTableStyleFlags`, ...identity });
      }
    }
    if (element.nativeFirstRowStyle !== undefined && typeof element.nativeFirstRowStyle !== "boolean") add(context, { code: "invalid-native-first-row-style", message: "Native first-row table style must be a boolean.", path: `${path}.nativeFirstRowStyle`, ...identity });
    if (element.style !== undefined) context.diagnostics.push(...validateTableStyle(element, path, identity));
    if (element.columns.length === 0 || element.columns.some((width) => !isFiniteNumber(width) || width <= 0)) add(context, { code: "invalid-table-columns", message: "Table columns must contain positive widths.", path: `${path}.columns`, ...identity });
    element.rows.forEach((row, rowIndex) => {
      if (row.height !== undefined && (!isFiniteNumber(row.height) || row.height <= 0)) add(context, { code: "invalid-table-row-height", message: "Table row height must be positive.", path: `${path}.rows.${rowIndex}.height`, ...identity });
      const span = row.cells.reduce((total, cell) => total + (cell.colSpan ?? 1), 0);
      if (span > element.columns.length) add(context, { code: "table-row-overflow", message: "Table row spans more columns than the table defines.", path: `${path}.rows.${rowIndex}`, ...identity });
      row.cells.forEach((cell, cellIndex) => {
        if (cell.headerScope !== undefined && cell.headerScope !== "row" && cell.headerScope !== "column") add(context, { code: "invalid-table-header-scope", message: 'Table header scope must be "row" or "column".', path: `${path}.rows.${rowIndex}.cells.${cellIndex}.headerScope`, ...identity });
        if (!Number.isInteger(cell.rowSpan ?? 1) || (cell.rowSpan ?? 1) < 1 || !Number.isInteger(cell.colSpan ?? 1) || (cell.colSpan ?? 1) < 1) add(context, { code: "invalid-table-span", message: "rowSpan and colSpan must be positive integers.", path: `${path}.rows.${rowIndex}.cells.${cellIndex}`, ...identity });
        validateText(context, cell.content, `${path}.rows.${rowIndex}.cells.${cellIndex}.content`, identity, decorative);
        validateTextStylePresetRef(context, cell.textStylePreset, `${path}.rows.${rowIndex}.cells.${cellIndex}.textStylePreset`, identity);
        if (cell.style?.textDirection !== undefined && !["horizontal", "rotate90", "rotate270", "eastAsianVertical", "mongolianVertical", "stackedVertical", "stackedVerticalRtl"].includes(cell.style.textDirection)) add(context, { code: "invalid-table-text-direction", message: "Unsupported table cell text direction.", path: `${path}.rows.${rowIndex}.cells.${cellIndex}.style.textDirection`, ...identity });
        validatePaint(context, cell.style?.fill, `${path}.rows.${rowIndex}.cells.${cellIndex}.style.fill`, identity);
        validateStroke(context, cell.style?.stroke, `${path}.rows.${rowIndex}.cells.${cellIndex}.style.stroke`, identity);
        if (cell.style?.borders !== undefined) {
          const at = `${path}.rows.${rowIndex}.cells.${cellIndex}.style.borders`, borders = cell.style.borders;
          if (!borders || typeof borders !== "object" || Array.isArray(borders) || Object.keys(borders).some(side => !["left", "right", "top", "bottom"].includes(side))) {
            add(context, { code: "invalid-table-borders", message: "Table borders require a record of left/right/top/bottom stroke overrides.", path: at, ...identity });
          } else for (const side of ["left", "right", "top", "bottom"] as const) {
            const stroke = borders[side];
            if (stroke === undefined) continue;
            if (!stroke || typeof stroke !== "object" || Array.isArray(stroke)) add(context, { code: "invalid-table-borders", message: "A border override must be a stroke record.", path: `${at}.${side}`, ...identity });
            else {
              validateStroke(context, stroke, `${at}.${side}`, identity);
              if (Object.keys(stroke).some(key => !["paint", "width", "dash", "beginArrow", "endArrow"].includes(key)) || (stroke.dash !== undefined && !["solid", "dash", "dot", "dashDot"].includes(stroke.dash)) || (stroke.beginArrow !== undefined && stroke.beginArrow !== "none") || (stroke.endArrow !== undefined && stroke.endArrow !== "none")) add(context, { code: "invalid-table-borders", message: "Table border overrides support known stroke fields and finite line dashes without arrowheads.", path: `${at}.${side}`, ...identity });
            }
          }
        }
      });
    });
  } else {
    const _exhaustive: never = element;
    void _exhaustive;
  }
}

function validateLayout(context: Context, layout: SlideLayoutDefinition, index: number, masterElements: readonly PresentationElement[] = []): void {
  const identity = { layoutId: layout.id };
  if (layout.showMasterShapes !== undefined && typeof layout.showMasterShapes !== "boolean") add(context, { code: "invalid-artwork-visibility", message: "showMasterShapes must be boolean.", path: `layouts.${index}.showMasterShapes`, ...identity });
  const keys = layout.placeholders.map((placeholder) => placeholder.key).filter(key => typeof key === "string" && key.trim().length > 0);
  for (const key of duplicates(keys)) add(context, { code: "duplicate-placeholder-key", message: `Duplicate placeholder key "${key}".`, path: `layouts.${index}.placeholders`, ...identity });
  const indices = new Set<number>();
  const boxes = new Map(layout.placeholders.map((placeholder) => [placeholder.key, placeholder.box]));
  layout.placeholders.forEach((placeholder, placeholderIndex) => {
    if (typeof placeholder.key !== "string" || !placeholder.key.trim()) add(context, { code: "invalid-placeholder-key", message: "Placeholder key must be a nonempty string.", path: `layouts.${index}.placeholders.${placeholderIndex}.key`, ...identity });
    if (!PLACEHOLDER_KINDS.has(placeholder.kind)) add(context, { code: "invalid-placeholder-kind", message: "Placeholder kind is unsupported.", path: `layouts.${index}.placeholders.${placeholderIndex}.kind`, ...identity });
    if (placeholder.index !== undefined) {
      const path = `layouts.${index}.placeholders.${placeholderIndex}.index`;
      if (!Number.isInteger(placeholder.index) || placeholder.index < 0 || placeholder.index > 4294967295) add(context, { code: "invalid-placeholder-index", message: "Placeholder index must be an unsigned 32-bit integer.", path, ...identity });
      else if (indices.has(placeholder.index)) add(context, { code: "duplicate-placeholder-index", message: `Duplicate placeholder index "${placeholder.index}".`, path, ...identity });
      else indices.add(placeholder.index);
    }
    if (!isValidBox(placeholder.box)) add(context, { code: "invalid-placeholder-box", message: "Placeholder box is invalid.", path: `layouts.${index}.placeholders.${placeholderIndex}.box`, ...identity });
    validateFrame(context, placeholder.textStyle?.frame, `layouts.${index}.placeholders.${placeholderIndex}.textStyle.frame`, identity);
    validateParagraphStyle(context, placeholder.textStyle?.paragraph, `layouts.${index}.placeholders.${placeholderIndex}.textStyle.paragraph`, identity);
    validateRunStyle(context, placeholder.textStyle?.run, `layouts.${index}.placeholders.${placeholderIndex}.textStyle.run`, identity);
    validateTextStylePresetRef(context, placeholder.textStylePreset, `layouts.${index}.placeholders.${placeholderIndex}.textStylePreset`, identity);
  });
  if (!isValidBackground(layout.background)) add(context, { code: "invalid-paint", message: "Background must be a valid paint or linear gradient.", path: `layouts.${index}.background`, ...identity });
  const scopeIds = new Set<string>();
  collectScopeElementIds(layout.elements, scopeIds);
  collectScopeElementIds(masterElements, scopeIds);
  layout.elements.forEach((element, elementIndex) => validateElement(context, element, `layouts.${index}.elements.${elementIndex}`, scopeIds, boxes, identity));
}

function validateAsset(context: Context, asset: PresentationAsset, index: number): void {
  if (asset.source.type === "bytes" && (asset.source.encoding !== "base64" || !isCanonicalBase64(asset.source.value))) add(context, { code: "invalid-asset-bytes", message: "Byte asset source must contain canonical Base64.", path: `assets.${index}.source`, assetId: asset.id });
  if (typeof asset.source.value !== "string" || asset.source.value.trim() === "") add(context, { code: "invalid-asset-source", message: "Asset source value cannot be empty.", path: `assets.${index}.source.value`, assetId: asset.id });
  if (asset.kind === "audio" && (asset.width !== undefined || asset.height !== undefined)) add(context, { code: "invalid-audio-dimensions", message: "Audio assets cannot have image dimensions.", path: `assets.${index}`, assetId: asset.id });
  for (const key of ["width", "height"] as const) {
    const value = asset[key];
    if (value !== undefined && (!isFiniteNumber(value) || value <= 0)) add(context, { code: "invalid-asset-dimension", message: `Asset ${key} must be positive.`, path: `assets.${index}.${key}`, assetId: asset.id });
  }
}

export function validatePresentation(document: Pick<PresentationData, "slides" | "assets" | "textStylePresets" | "layouts" | "masters" | "theme" | "firstSlideNumber">): PresentationDiagnostic[] {
  const diagnostics: PresentationDiagnostic[] = [];
  const slideIds = new Set(document.slides.map((slide) => slide.id));
  const textStylePresetNames = new Set(Object.keys(document.textStylePresets));
  const context: Context = { diagnostics, slideIds, assetKinds: new Map(document.assets.map(asset => [asset.id, asset.kind])), globalElementIds: new Set(), textStylePresetNames };
  if (document.firstSlideNumber !== undefined && (!Number.isSafeInteger(document.firstSlideNumber) || document.firstSlideNumber < 0)) add(context, { code: "invalid-first-slide-number", message: "First slide number must be a nonnegative safe integer.", path: "firstSlideNumber" });

  for (const id of duplicates(document.slides.map((slide) => slide.id))) add(context, { code: "duplicate-slide-id", message: `Duplicate slide id "${id}".`, path: "slides" });
  for (const id of duplicates(document.assets.map((asset) => asset.id))) add(context, { code: "duplicate-asset-id", message: `Duplicate asset id "${id}".`, path: "assets", assetId: id });
  for (const id of duplicates(document.layouts.map((layout) => layout.id))) add(context, { code: "duplicate-layout-id", message: `Duplicate layout id "${id}".`, path: "layouts", layoutId: id });

  for (const [name, preset] of Object.entries(document.textStylePresets)) {
    if (name.trim() === "") add(context, { code: "invalid-text-style-preset-name", message: "Text style preset names must be non-empty.", path: "textStylePresets" });
    validateTextStylePreset(context, preset, `textStylePresets.${name}`);
  }

  for (const [role, color] of Object.entries(document.theme.colors ?? {})) if (!/^#?[0-9a-f]{6}$/i.test(color)) add(context, { code: "invalid-theme-color", message: `Theme color "${role}" must be six-digit RGB.`, path: `theme.colors.${role}` });
  document.assets.forEach((asset, index) => validateAsset(context, asset, index));
  document.layouts.forEach((layout, index) => validateLayout(context, layout, index, (document.masters ?? []).find(master => master.id === layout.masterId)?.elements));

  const masters = document.masters ?? [];
  for (const id of duplicates(masters.map(master => master.id))) add(context, { code: "duplicate-master-id", message: `Duplicate master id "${id}".`, path: "masters" });
  masters.forEach((master, index) => {
    const path = `masters.${index}`, start = diagnostics.length;
    if (Object.keys(master).some(key => !["id", "name", "theme", "background", "elements"].includes(key))) add(context, { code: "invalid-master-definition", message: "Unknown master declaration field.", path });
    if (typeof master.id !== "string" || master.id.trim() === "") add(context, { code: "invalid-master-id", message: "Master id must be a nonempty string.", path: `${path}.id` });
    if (typeof master.name !== "string" || master.name.trim() === "") add(context, { code: "invalid-master-name", message: "Master name must be a nonempty string.", path: `${path}.name` });
    if (master.theme !== undefined && (!master.theme || typeof master.theme !== "object" || Array.isArray(master.theme) || Object.keys(master.theme).some(key => !["name", "colors", "fonts"].includes(key)))) add(context, { code: "invalid-master-theme", message: "Master theme must contain name, colors and fonts only.", path: `${path}.theme` });
    if (master.theme?.name !== undefined && (typeof master.theme.name !== "string" || master.theme.name.trim() === "")) add(context, { code: "invalid-theme-name", message: "Master theme name must be nonempty.", path: `${path}.theme.name` });
    for (const component of ["colors", "fonts"] as const) {
      const value = master.theme?.[component];
      if (value !== undefined && (!value || typeof value !== "object" || Array.isArray(value))) add(context, { code: "invalid-master-theme", message: "Master theme components must be objects.", path: `${path}.theme.${component}` });
    }
    for (const [role, color] of Object.entries(master.theme?.colors ?? {})) if (!Object.hasOwn(DEFAULT_THEME.colors, role) || typeof color !== "string" || !/^#?[0-9a-f]{6}$/i.test(color)) add(context, { code: "invalid-theme-color", message: "Master theme colors must be six-digit RGB.", path: `${path}.theme.colors.${role}` });
    for (const [role, font] of Object.entries(master.theme?.fonts ?? {})) if (!["heading", "body"].includes(role) || typeof font !== "string" || font.trim() === "") add(context, { code: "invalid-font-family", message: "Master theme font must be a nonempty name.", path: `${path}.theme.fonts.${role}` });
    if (!isValidBackground(master.background)) add(context, { code: "invalid-paint", message: "Background must be a valid paint or linear gradient.", path: `${path}.background` });
    if (master.elements !== undefined && !Array.isArray(master.elements)) add(context, { code: "invalid-master-elements", message: "Master elements must be an array.", path: `${path}.elements` });
    else {
      const scopeIds = new Set<string>(); collectScopeElementIds(master.elements ?? [], scopeIds);
      (master.elements ?? []).forEach((element, i) => validateElement(context, element, `${path}.elements.${i}`, scopeIds, new Map(), { masterId: master.id }));
    }
    for (let i = start; i < diagnostics.length; i++) diagnostics[i] = { ...diagnostics[i]!, masterId: master.id };
  });
  const masterIds = new Set(masters.map(master => master.id));
  document.layouts.forEach((layout, index) => { if (layout.masterId !== undefined && !masterIds.has(layout.masterId)) add(context, { code: "missing-layout-master", message: `Layout references missing master "${layout.masterId}".`, path: `layouts.${index}.masterId`, layoutId: layout.id }); });
  const layoutMap = new Map(document.layouts.map((layout) => [layout.id, layout]));
  document.slides.forEach((slide: PresentationSlideData, slideIndex) => {
    const identity = { slideId: slide.id };
    const layout = layoutMap.get(slide.layoutId ?? DEFAULT_LAYOUT_ID);
    if (slide.layoutId !== undefined && layout === undefined) add(context, { code: "missing-slide-layout", message: `Slide references missing layout "${slide.layoutId}".`, path: `slides.${slideIndex}.layoutId`, ...identity });
    if (slide.showInheritedShapes !== undefined && typeof slide.showInheritedShapes !== "boolean") add(context, { code: "invalid-artwork-visibility", message: "showInheritedShapes must be boolean.", path: `slides.${slideIndex}.showInheritedShapes`, ...identity });
    if (!isValidBackground(slide.background)) add(context, { code: "invalid-paint", message: "Background must be a valid paint or linear gradient.", path: `slides.${slideIndex}.background`, ...identity });
    if (slide.notes !== undefined) validateText(context, slide.notes, `slides.${slideIndex}.notes`, identity);
    validateTransition(context, slide.transition, `slides.${slideIndex}.transition`, identity);
    if (slide.readingOrder !== undefined) {
      try { captureSlideReadingOrder(slide.elements, slide.readingOrder); }
      catch (error) { add(context, { code: "invalid-reading-order", message: error instanceof Error ? error.message : "Invalid reading order.", path: `slides.${slideIndex}.readingOrder`, ...identity }); }
    }
    if (slide.eventTimeline !== undefined) {
      try { compileSlideEventTimeline(slide.id, slide.elements, slide.eventTimeline); }
      catch (error) {
        if (!(error instanceof EventTimelineError)) throw error;
        add(context, { code: error.code, message: error.message, path: `slides.${slideIndex}.eventTimeline${error.path.replaceAll("/", ".")}`, ...identity });
      }
    }
    const placeholderBoxes = new Map((layout?.placeholders ?? []).map((placeholder) => [placeholder.key, placeholder.box]));
    const scopeIds = new Set<string>();
    collectScopeElementIds(slide.elements, scopeIds);
    if (layout !== undefined) { collectScopeElementIds(layout.elements, scopeIds); collectScopeElementIds(masters.find(master => master.id === layout.masterId)?.elements ?? [], scopeIds); }
    slide.elements.forEach((element, elementIndex) => validateElement(context, element, `slides.${slideIndex}.elements.${elementIndex}`, scopeIds, placeholderBoxes, identity));
  });
  return diagnostics;
}
