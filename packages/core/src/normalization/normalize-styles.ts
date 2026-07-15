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

const POWERPOINT_DEFAULT_LIST_INDENT = 27;

export function normalizePaint(input: PaintInput | undefined, fallback: NormalizedPaint = { type: "none" }): NormalizedPaint {
  if (input === undefined) return deepClone(fallback);
  if (input.type === "none") return { type: "none" };
  return { type: "solid", color: deepClone(input.color), opacity: input.opacity ?? 1 };
}

export function normalizeStroke(input: StrokeStyleInput | undefined, fallback: NormalizedStrokeStyle = DEFAULT_STROKE): NormalizedStrokeStyle {
  return {
    paint: normalizePaint(input?.paint, fallback.paint),
    width: input?.width ?? fallback.width,
    dash: input?.dash ?? fallback.dash,
    beginArrow: input?.beginArrow ?? fallback.beginArrow,
    endArrow: input?.endArrow ?? fallback.endArrow,
  };
}

function normalizeInsets(input: number | Partial<Insets> | undefined, fallback: Insets): Insets {
  if (typeof input === "number") return { top: input, right: input, bottom: input, left: input };
  return {
    top: input?.top ?? fallback.top,
    right: input?.right ?? fallback.right,
    bottom: input?.bottom ?? fallback.bottom,
    left: input?.left ?? fallback.left,
  };
}

function normalizeAutoFit(input: TextAutoFitInput | undefined, fallback: NormalizedTextAutoFit): NormalizedTextAutoFit {
  if (input === undefined) return deepClone(fallback);
  if (input.mode === "shrink") return { mode: "shrink", fontScale: input.fontScale ?? 1 };
  return { mode: input.mode };
}

export function normalizeTextFrame(input?: TextFrameStyleInput, fallback: NormalizedTextFrameStyle = DEFAULT_TEXT_FRAME_STYLE): NormalizedTextFrameStyle {
  return {
    margin: normalizeInsets(input?.margin, fallback.margin),
    verticalAlign: input?.verticalAlign ?? fallback.verticalAlign,
    wrap: input?.wrap ?? fallback.wrap,
    autoFit: normalizeAutoFit(input?.autoFit, fallback.autoFit),
  };
}

function normalizeBullet(input: TextBulletInput | undefined, fallback: NormalizedTextBullet): NormalizedTextBullet {
  if (input === undefined) return deepClone(fallback);
  if (input.type === "none") return { type: "none" };
  if (input.type === "bullet") return { type: "bullet", character: input.character ?? "•" };
  return { type: "number", style: input.style ?? "arabicPeriod", startAt: input.startAt ?? 1 };
}

export function normalizeTextParagraphStyle(input?: TextParagraphStyleInput, fallback: NormalizedTextParagraphStyle = DEFAULT_TEXT_PARAGRAPH_STYLE): NormalizedTextParagraphStyle {
  const bullet = normalizeBullet(input?.bullet, fallback.bullet);
  const enablesList = input?.bullet !== undefined && input.bullet.type !== "none" && fallback.bullet.type === "none";
  const defaultIndent = enablesList ? POWERPOINT_DEFAULT_LIST_INDENT : fallback.indent;
  const defaultHanging = enablesList ? POWERPOINT_DEFAULT_LIST_INDENT : fallback.hanging;
  return {
    align: input?.align ?? fallback.align,
    indent: input?.indent ?? defaultIndent,
    hanging: input?.hanging ?? defaultHanging,
    lineSpacing: input?.lineSpacing ?? fallback.lineSpacing,
    spaceBefore: input?.spaceBefore ?? fallback.spaceBefore,
    spaceAfter: input?.spaceAfter ?? fallback.spaceAfter,
    bullet,
  };
}

export function normalizeTextRunStyle(input?: TextRunStyleInput, fallback: NormalizedTextRunStyle = DEFAULT_TEXT_RUN_STYLE): NormalizedTextRunStyle {
  return {
    fontFamily: deepClone(input?.fontFamily ?? fallback.fontFamily),
    fontSize: input?.fontSize ?? fallback.fontSize,
    bold: input?.bold ?? fallback.bold,
    italic: input?.italic ?? fallback.italic,
    underline: input?.underline ?? fallback.underline,
    strike: input?.strike ?? fallback.strike,
    color: deepClone(input?.color ?? fallback.color),
    language: input?.language ?? fallback.language,
  };
}

export function normalizeShapeStyle(input?: ShapeStyleInput): NormalizedShapeStyle {
  return {
    fill: normalizePaint(input?.fill, DEFAULT_SHAPE_STYLE.fill),
    stroke: normalizeStroke(input?.stroke, DEFAULT_SHAPE_STYLE.stroke),
  };
}

export function normalizeTableCellStyle(input?: TableCellStyleInput): NormalizedTableCellStyle {
  return {
    fill: normalizePaint(input?.fill, DEFAULT_TABLE_CELL_STYLE.fill),
    stroke: normalizeStroke(input?.stroke, DEFAULT_TABLE_CELL_STYLE.stroke),
    margin: normalizeInsets(input?.margin, DEFAULT_TABLE_CELL_STYLE.margin),
    verticalAlign: input?.verticalAlign ?? DEFAULT_TABLE_CELL_STYLE.verticalAlign,
  };
}
