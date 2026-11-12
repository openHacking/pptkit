import {
  DEFAULT_SHAPE_STYLE,
  DEFAULT_STROKE,
  DEFAULT_TABLE_CELL_STYLE,
  DEFAULT_TEXT_FRAME_STYLE,
  DEFAULT_TEXT_PARAGRAPH_STYLE,
  DEFAULT_TEXT_RUN_STYLE,
} from "../constants/presentation.js";
import type { Insets } from "../types/geometry.js";
import type {
  NormalizedPaint,
  NormalizedShapeStyle,
  NormalizedStrokeStyle,
  NormalizedTableCellStyle,
  NormalizedTextAutoFit,
  NormalizedTextBullet,
  NormalizedTextFrameStyle,
  NormalizedTextParagraphStyle,
  NormalizedTextRunStyle,
  PaintInput,
  ShapeStyleInput,
  StrokeStyleInput,
  TableCellStyleInput,
  TextAutoFitInput,
  TextBulletInput,
  TextFrameStyleInput,
  TextParagraphStyleInput,
  TextRunStyleInput,
} from "../types/style.js";
import { deepClone } from "../utils/clone.js";
import { selectStyle, type StyleTrace } from "./style-origins.js";
import type { StyleInputTrace } from "./style-input-origins.js";
import type { StyleOriginMap } from "./style-origins.js";
import { TEXT_FONT_SCRIPT_CODES } from "../text/font-scripts.js";

const POWERPOINT_DEFAULT_LIST_INDENT = 27;

export function normalizePaint(input: PaintInput | undefined, fallback: NormalizedPaint = { type: "none" }, trace?: StyleTrace, prefix = "", fallbackOrigins?: StyleOriginMap): NormalizedPaint {
  const choose = <T>(key: string, local: T | undefined, value: T) => selectStyle(trace, `${prefix}${key}`, local, value, `${prefix}${key}`, fallbackOrigins?.[key]);
  if (input === undefined) {
    choose("type", undefined, fallback.type);
    if (fallback.type === "solid") { choose("color", undefined, fallback.color); choose("opacity", undefined, fallback.opacity); }
    return deepClone(fallback);
  }
  choose("type", input.type, fallback.type);
  if (input.type === "none") return { type: "none" };
  return { type: "solid", color: deepClone(choose("color", input.color, input.color)), opacity: selectStyle(trace, `${prefix}opacity`, input.opacity, 1, `${prefix}opacity`, trace?.builtin(`${prefix}opacity`)) };
}

export function normalizeStroke(input: StrokeStyleInput | undefined, fallback: NormalizedStrokeStyle = DEFAULT_STROKE, trace?: StyleTrace, prefix = "", fallbackOrigins?: StyleOriginMap): NormalizedStrokeStyle {
  const choose = <T>(key: string, local: T | undefined, value: T) => selectStyle(trace, `${prefix}${key}`, local, value, `${prefix}${key}`, fallbackOrigins?.[key]);
  const paintOrigins = fallbackOrigins && Object.fromEntries(Object.entries(fallbackOrigins).filter(([key]) => key.startsWith("paint.")).map(([key, value]) => [key.slice(6), value]));
  return {
    paint: normalizePaint(input?.paint, fallback.paint, trace, `${prefix}paint.`, paintOrigins),
    width: choose("width", input?.width, fallback.width),
    dash: choose("dash", input?.dash, fallback.dash),
    beginArrow: choose("beginArrow", input?.beginArrow, fallback.beginArrow),
    endArrow: choose("endArrow", input?.endArrow, fallback.endArrow),
  };
}

function normalizeInsets(input: number | Partial<Insets> | undefined, fallback: Insets, trace?: StyleTrace): Insets {
  const side = (key: keyof Insets) => selectStyle(trace, `margin.${key}`, typeof input === "number" ? input : input?.[key], fallback[key], typeof input === "number" ? "margin" : `margin.${key}`);
  return { top: side("top"), right: side("right"), bottom: side("bottom"), left: side("left") };
}

function normalizeAutoFit(input: TextAutoFitInput | undefined, fallback: NormalizedTextAutoFit, trace?: StyleTrace): NormalizedTextAutoFit {
  if (input === undefined) {
    selectStyle(trace, "autoFit.mode", undefined, fallback.mode);
    if (fallback.mode === "shrink") selectStyle(trace, "autoFit.fontScale", undefined, fallback.fontScale);
    return deepClone(fallback);
  }
  const mode = selectStyle(trace, "autoFit.mode", input.mode, fallback.mode);
  if (mode === "shrink" && input.mode === "shrink") return { mode, fontScale: selectStyle(trace, "autoFit.fontScale", input.fontScale, 1, "autoFit.fontScale", trace?.builtin("autoFit.fontScale")) };
  return { mode } as NormalizedTextAutoFit;
}

export function normalizeTextFrame(input?: TextFrameStyleInput, fallback: NormalizedTextFrameStyle = DEFAULT_TEXT_FRAME_STYLE, trace?: StyleTrace): NormalizedTextFrameStyle {
  const textDirection = selectStyle(trace, "textDirection", input?.textDirection, fallback.textDirection);
  return {
    ...(textDirection === undefined ? {} : { textDirection }),
    margin: normalizeInsets(input?.margin, fallback.margin, trace),
    verticalAlign: selectStyle(trace, "verticalAlign", input?.verticalAlign, fallback.verticalAlign),
    wrap: selectStyle(trace, "wrap", input?.wrap, fallback.wrap),
    autoFit: normalizeAutoFit(input?.autoFit, fallback.autoFit, trace),
  };
}

function normalizeBullet(input: TextBulletInput | undefined, fallback: NormalizedTextBullet, trace?: StyleTrace): NormalizedTextBullet {
  if (input === undefined) {
    selectStyle(trace, "bullet.type", undefined, fallback.type);
    if (fallback.type === "bullet") selectStyle(trace, "bullet.character", undefined, fallback.character);
    if (fallback.type === "number") { selectStyle(trace, "bullet.style", undefined, fallback.style); selectStyle(trace, "bullet.startAt", undefined, fallback.startAt); }
    return deepClone(fallback);
  }
  selectStyle(trace, "bullet.type", input.type, fallback.type);
  if (input.type === "none") return { type: "none" };
  if (input.type === "bullet") return { type: "bullet", character: selectStyle(trace, "bullet.character", input.character, "•", "bullet.character", trace?.builtin("bullet.character")) };
  return { type: "number", style: selectStyle(trace, "bullet.style", input.style, "arabicPeriod", "bullet.style", trace?.builtin("bullet.style")), startAt: selectStyle(trace, "bullet.startAt", typeof input.startAt === "number" ? input.startAt : undefined, 1, "bullet.startAt", trace?.builtin("bullet.startAt")) };
}

export function normalizeTextParagraphStyle(input?: TextParagraphStyleInput, fallback: NormalizedTextParagraphStyle = DEFAULT_TEXT_PARAGRAPH_STYLE, trace?: StyleTrace): NormalizedTextParagraphStyle {
  const bullet = normalizeBullet(input?.bullet, fallback.bullet, trace);
  const tabStops = selectStyle(trace, "tabStops", input?.tabStops, fallback.tabStops);
  const direction = selectStyle(trace, "direction", input?.direction, fallback.direction);
  const enablesList = input?.bullet !== undefined && input.bullet.type !== "none" && fallback.bullet.type === "none";
  const listOrigin = trace ? { source: { kind: "default" as const, path: "defaults.text.paragraph.listIndent" } } : undefined;
  return {
    align: selectStyle(trace, "align", input?.align, fallback.align),
    ...(direction === undefined ? {} : { direction }),
    indent: selectStyle(trace, "indent", input?.indent, enablesList ? POWERPOINT_DEFAULT_LIST_INDENT : fallback.indent, "indent", enablesList ? listOrigin : undefined),
    hanging: selectStyle(trace, "hanging", input?.hanging, enablesList ? POWERPOINT_DEFAULT_LIST_INDENT : fallback.hanging, "hanging", enablesList ? listOrigin : undefined),
    lineSpacing: selectStyle(trace, "lineSpacing", input?.lineSpacing, fallback.lineSpacing),
    spaceBefore: selectStyle(trace, "spaceBefore", input?.spaceBefore, fallback.spaceBefore),
    spaceAfter: selectStyle(trace, "spaceAfter", input?.spaceAfter, fallback.spaceAfter),
    bullet,
    ...(tabStops === undefined ? {} : { tabStops: deepClone(tabStops) }),
  };
}

export function normalizeTextRunStyle(input?: TextRunStyleInput, fallback: NormalizedTextRunStyle = DEFAULT_TEXT_RUN_STYLE, trace?: StyleTrace): NormalizedTextRunStyle {
  const colorOpacity = selectStyle(trace, "colorOpacity", input?.colorOpacity, fallback.colorOpacity ?? 1);
  const rightToLeft = selectStyle(trace, "rightToLeft", input?.rightToLeft, fallback.rightToLeft);
  const optionalFont = (property: "eastAsianFontFamily" | "complexScriptFontFamily") => {
    if (input?.[property] === null) { trace?.clear(property); return undefined; }
    return selectStyle(trace, property, input?.[property], fallback[property]);
  };
  const eastAsian = optionalFont("eastAsianFontFamily"), complexScript = optionalFont("complexScriptFontFamily");
  const map = input?.scriptFontFamilies;
  const scriptFontFamilies = map === undefined && fallback.scriptFontFamilies === undefined && trace === undefined ? undefined : Object.fromEntries(TEXT_FONT_SCRIPT_CODES.flatMap(code => {
    const property = `scriptFontFamilies.${code}`;
    if (map === null || map?.[code] === null) { trace?.clear(property, map === null ? "scriptFontFamilies" : property); return []; }
    const value = selectStyle(trace, property, map?.[code], fallback.scriptFontFamilies?.[code]);
    return value === undefined ? [] : [[code, deepClone(value)]];
  }));
  return {
    fontFamily: deepClone(selectStyle(trace, "fontFamily", input?.fontFamily, fallback.fontFamily)),
    ...(eastAsian === undefined ? {} : { eastAsianFontFamily: deepClone(eastAsian) }),
    ...(complexScript === undefined ? {} : { complexScriptFontFamily: deepClone(complexScript) }),
    ...(scriptFontFamilies && Object.keys(scriptFontFamilies).length ? { scriptFontFamilies } : {}),
    ...(rightToLeft === undefined ? {} : { rightToLeft }),
    fontSize: selectStyle(trace, "fontSize", input?.fontSize, fallback.fontSize),
    bold: selectStyle(trace, "bold", input?.bold, fallback.bold),
    italic: selectStyle(trace, "italic", input?.italic, fallback.italic),
    underline: selectStyle(trace, "underline", input?.underline, fallback.underline),
    strike: selectStyle(trace, "strike", input?.strike, fallback.strike),
    color: deepClone(selectStyle(trace, "color", input?.color, fallback.color)),
    ...(colorOpacity === 1 ? {} : { colorOpacity }),
    language: selectStyle(trace, "language", input?.language, fallback.language),
  };
}

/** Retain whole-map reset barriers while combining regional/preset/local input. */
export function mergeTextRunStyleInput(base: TextRunStyleInput | undefined, patch: TextRunStyleInput | undefined, trace?: StyleInputTrace): TextRunStyleInput {
  const fields = Object.entries(patch ?? {}).filter(([, value]) => value !== undefined);
  for (const [key, value] of fields) if (key !== "scriptFontFamilies") trace?.record(key, key, value === null);
  if (patch?.scriptFontFamilies === null) for (const code of TEXT_FONT_SCRIPT_CODES) trace?.record(`scriptFontFamilies.${code}`, "scriptFontFamilies", true);
  else for (const [code, value] of Object.entries(patch?.scriptFontFamilies ?? {})) if (value !== undefined) trace?.record(`scriptFontFamilies.${code}`, `scriptFontFamilies.${code}`, value === null);
  const result = { ...base, ...Object.fromEntries(fields) };
  if (patch?.scriptFontFamilies !== undefined && patch.scriptFontFamilies !== null) {
    const inherited = base?.scriptFontFamilies === null ? Object.fromEntries(TEXT_FONT_SCRIPT_CODES.map(code => [code, null])) : base?.scriptFontFamilies;
    result.scriptFontFamilies = { ...inherited, ...Object.fromEntries(Object.entries(patch.scriptFontFamilies).filter(([, value]) => value !== undefined)) };
  }
  return result;
}

export function mergeTextParagraphStyleInput(base: TextParagraphStyleInput | undefined, patch: TextParagraphStyleInput | undefined, trace?: StyleInputTrace): TextParagraphStyleInput {
  const fields = Object.entries(patch ?? {}).filter(([, value]) => value !== undefined);
  for (const [key, value] of fields) {
    if (key === "bullet") trace?.replaceObject(key, value as object);
    else trace?.record(key);
  }
  return { ...base, ...Object.fromEntries(fields) };
}

export function normalizeShapeStyle(input?: ShapeStyleInput, trace?: StyleTrace): NormalizedShapeStyle {
  const fill = input?.fill;
  const shadow = input?.shadow === undefined ? {} : { shadow: {
    color: deepClone(selectStyle(trace, "shadow.color", input.shadow.color, input.shadow.color)),
    opacity: selectStyle(trace, "shadow.opacity", input.shadow.opacity, 1),
    blur: selectStyle(trace, "shadow.blur", input.shadow.blur, input.shadow.blur),
    distance: selectStyle(trace, "shadow.distance", input.shadow.distance, input.shadow.distance),
    direction: selectStyle(trace, "shadow.direction", input.shadow.direction, input.shadow.direction),
  } };
  if (fill?.type === "pattern") {
    selectStyle(trace, "fill.type", fill.type, fill.type);
    const pattern = selectStyle(trace, "fill.pattern", fill.pattern, fill.pattern);
    const foreground = deepClone(selectStyle(trace, "fill.foreground", fill.foreground, fill.foreground));
    const background = deepClone(selectStyle(trace, "fill.background", fill.background, fill.background));
    const foregroundOpacity = selectStyle(trace, "fill.foregroundOpacity", fill.foregroundOpacity, 1);
    const backgroundOpacity = selectStyle(trace, "fill.backgroundOpacity", fill.backgroundOpacity, 1);
    return { fill: { type: "pattern", pattern, foreground, background, foregroundOpacity, backgroundOpacity }, stroke: normalizeStroke(input?.stroke, DEFAULT_SHAPE_STYLE.stroke, trace, "stroke."), ...shadow };
  }
  if (fill?.type === "linearGradient") {
    selectStyle(trace, "fill.type", fill.type, fill.type);
    const angle = selectStyle(trace, "fill.angle", fill.angle, fill.angle);
    const scaled = selectStyle(trace, "fill.scaled", fill.scaled, false, "fill.scaled", { source: { kind: "default", path: "defaults.shape.fill.linearGradient.scaled" } });
    selectStyle(trace, "fill.stops", fill.stops, fill.stops);
    const stops = fill.stops.map((stop, index) => ({
      position: selectStyle(trace, `fill.stops.${index}.position`, stop.position, stop.position),
      color: deepClone(selectStyle(trace, `fill.stops.${index}.color`, stop.color, stop.color)),
      opacity: selectStyle(trace, `fill.stops.${index}.opacity`, stop.opacity, 1, `fill.stops.${index}.opacity`, { source: { kind: "default", path: "defaults.shape.fill.linearGradient.stop.opacity" } }),
    }));
    return { fill: { type: "linearGradient", angle, scaled, stops }, stroke: normalizeStroke(input?.stroke, DEFAULT_SHAPE_STYLE.stroke, trace, "stroke."), ...shadow };
  }
  return {
    fill: normalizePaint(fill, DEFAULT_SHAPE_STYLE.fill.type === "linearGradient" || DEFAULT_SHAPE_STYLE.fill.type === "pattern" ? { type: "none" } : DEFAULT_SHAPE_STYLE.fill, trace, "fill."),
    stroke: normalizeStroke(input?.stroke, DEFAULT_SHAPE_STYLE.stroke, trace, "stroke."),
    ...shadow,
  };
}

export function normalizeTableCellStyle(input?: TableCellStyleInput, trace?: StyleTrace): NormalizedTableCellStyle {
  const stroke = normalizeStroke(input?.stroke, DEFAULT_TABLE_CELL_STYLE.stroke, trace, "stroke.");
  const strokeOrigins = trace && Object.fromEntries(Object.entries(trace.origins).filter(([key]) => key.startsWith("stroke.")).map(([key, value]) => [key.slice(7), value]));
  const borders = input?.borders === undefined ? undefined : Object.fromEntries(
    (["left", "right", "top", "bottom"] as const).flatMap(side => input.borders?.[side] === undefined ? [] : [[side, normalizeStroke(input.borders[side], stroke, trace, `borders.${side}.`, strokeOrigins)]]),
  );
  return {
    fill: normalizePaint(input?.fill, DEFAULT_TABLE_CELL_STYLE.fill, trace, "fill."),
    stroke,
    ...(borders === undefined ? {} : { borders }),
    margin: normalizeInsets(input?.margin, DEFAULT_TABLE_CELL_STYLE.margin, trace),
    verticalAlign: selectStyle(trace, "verticalAlign", input?.verticalAlign, DEFAULT_TABLE_CELL_STYLE.verticalAlign),
    ...(input?.textDirection === undefined ? {} : { textDirection: selectStyle(trace, "textDirection", input.textDirection, input.textDirection) }),
  };
}
