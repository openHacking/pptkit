import type { Insets } from "./geometry.js";
import type { TextFontScriptCode } from "../text/font-scripts.js";
export type { TextFontScriptCode } from "../text/font-scripts.js";

export type ThemeColorRole =
  | "background1"
  | "background2"
  | "text1"
  | "text2"
  | "accent1"
  | "accent2"
  | "accent3"
  | "accent4"
  | "accent5"
  | "accent6"
  | "hyperlink"
  | "followedHyperlink";

export type ThemeFontRole = "heading" | "body";

export type ColorValue = string | { theme: ThemeColorRole };
export type FontFamilyValue = string | { theme: ThemeFontRole };
export type ScriptFontFamiliesInput = Partial<Record<TextFontScriptCode, FontFamilyValue | null>>;
export type NormalizedScriptFontFamilies = Partial<Record<TextFontScriptCode, FontFamilyValue>>;

export type PaintInput =
  | { type: "none" }
  | { type: "solid"; color: ColorValue; opacity?: number };

export type NormalizedPaint =
  | { type: "none" }
  | { type: "solid"; color: ColorValue; opacity: number };

/** Linear gradient for slide backgrounds and shape fills. Angle is clockwise from the positive x axis. */
export type BackgroundPaintInput = PaintInput | { type: "linearGradient"; angle: number; scaled?: boolean; stops: readonly { position: number; color: ColorValue; opacity?: number }[] };
export type NormalizedBackgroundPaint = NormalizedPaint | { type: "linearGradient"; angle: number; scaled: boolean; stops: { position: number; color: ColorValue; opacity: number }[] };

/** Bounded shape-only preset pattern fill. */
export type ShapePattern = "horizontal" | "vertical" | "cross" | "downwardDiagonal" | "upwardDiagonal" | "diagonalCross" | "percent5" | "percent10" | "percent20" | "percent25" | "percent30" | "percent40" | "percent50" | "percent60" | "percent70" | "percent75" | "percent80" | "percent90";
export type ShapeFillPaintInput = BackgroundPaintInput | { type: "pattern"; pattern: ShapePattern; foreground: ColorValue; background: ColorValue; foregroundOpacity?: number; backgroundOpacity?: number };
export type NormalizedShapeFillPaint = NormalizedBackgroundPaint | { type: "pattern"; pattern: ShapePattern; foreground: ColorValue; background: ColorValue; foregroundOpacity: number; backgroundOpacity: number };

export type LineDash = "solid" | "dash" | "dot" | "dashDot";
export type ArrowType = "none" | "arrow" | "triangle" | "stealth" | "diamond" | "oval";

export interface StrokeStyleInput {
  paint?: PaintInput;
  width?: number;
  dash?: LineDash;
  beginArrow?: ArrowType;
  endArrow?: ArrowType;
}

export interface NormalizedStrokeStyle {
  paint: NormalizedPaint;
  width: number;
  dash: LineDash;
  beginArrow: ArrowType;
  endArrow: ArrowType;
}

export interface TransformInput {
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export interface NormalizedTransform {
  rotation: number;
  flipH: boolean;
  flipV: boolean;
}

export type TextAutoFitInput =
  | { mode: "none" }
  | { mode: "shrink"; fontScale?: number }
  | { mode: "resize" };

export type NormalizedTextAutoFit =
  | { mode: "none" }
  | { mode: "shrink"; fontScale: number }
  | { mode: "resize" };

export interface TextFrameStyleInput {
  /** Writing direction of the text inside its frame, independent of shape rotation. */
  textDirection?: "horizontal" | "rotate90" | "rotate270" | "eastAsianVertical" | "mongolianVertical" | "stackedVertical" | "stackedVerticalRtl";
  margin?: number | Partial<Insets>;
  verticalAlign?: "top" | "middle" | "bottom";
  wrap?: boolean;
  autoFit?: TextAutoFitInput;
}

export interface NormalizedTextFrameStyle {
  /** Writing direction of the text inside its frame, independent of shape rotation. */
  textDirection?: "horizontal" | "rotate90" | "rotate270" | "eastAsianVertical" | "mongolianVertical" | "stackedVertical" | "stackedVerticalRtl";
  margin: Insets;
  verticalAlign: "top" | "middle" | "bottom";
  wrap: boolean;
  autoFit: NormalizedTextAutoFit;
}

export type TextBulletInput =
  | { type: "none" }
  | { type: "bullet"; character?: string }
  | { type: "number"; style?: "arabicPeriod" | "arabicParen" | "alphaLowerPeriod" | "alphaUpperPeriod"; startAt?: number | "continue" };

export type NormalizedTextBullet =
  | { type: "none" }
  | { type: "bullet"; character: string }
  | { type: "number"; style: "arabicPeriod" | "arabicParen" | "alphaLowerPeriod" | "alphaUpperPeriod"; startAt: number };

/** Position is in points relative to the paragraph's left margin. */
export interface TextTabStop {
  position: number;
  align: "left" | "center" | "right" | "decimal";
}

export interface TextParagraphStyleInput {
  align?: "left" | "center" | "right" | "justify";
  /** Explicit paragraph base direction; omitted retains the inherited/default behavior. */
  direction?: "ltr" | "rtl";
  indent?: number;
  hanging?: number;
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  bullet?: TextBulletInput;
  tabStops?: TextTabStop[];
}

export interface NormalizedTextParagraphStyle {
  align: "left" | "center" | "right" | "justify";
  direction?: "ltr" | "rtl";
  indent: number;
  hanging: number;
  lineSpacing: number;
  spaceBefore: number;
  spaceAfter: number;
  bullet: NormalizedTextBullet;
  tabStops?: TextTabStop[];
}

export interface TextRunStyleInput {
  fontFamily?: FontFamilyValue;
  /** Explicit East Asian family; undefined inherits, null resets to fontFamily. */
  eastAsianFontFamily?: FontFamilyValue | null;
  /** Explicit complex-script family; undefined inherits, null resets to fontFamily. */
  complexScriptFontFamily?: FontFamilyValue | null;
  /** Per-script overrides. Undefined inherits; null map/key clears inherited values. */
  scriptFontFamilies?: ScriptFontFamiliesInput | null;
  /** DrawingML run-level RTL characteristics; distinct from paragraph direction. */
  rightToLeft?: boolean;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: ColorValue;
  /** Text paint alpha, independent of element opacity; 0 is transparent. */
  colorOpacity?: number;
  language?: string;
}

/** Named-preset properties bypassed in favor of placeholder inheritance. */
export type TextStylePresetResetProperty = keyof TextRunStyleInput | `scriptFontFamilies.${TextFontScriptCode}` | `paragraph.${keyof TextParagraphStyleInput}` | `frame.${keyof TextFrameStyleInput}` | `frame.margin.${keyof Insets}`;

export interface TextStylePresetInput {
  frame?: TextFrameStyleInput;
  paragraph?: TextParagraphStyleInput;
  run?: TextRunStyleInput;
}

export type TextStylePresetMap = Readonly<Record<string, TextStylePresetInput>>;

export interface NormalizedTextRunStyle {
  fontFamily: FontFamilyValue;
  /** Absent slots use fontFamily. No host font substitution is stored here. */
  eastAsianFontFamily?: FontFamilyValue;
  complexScriptFontFamily?: FontFamilyValue;
  scriptFontFamilies?: NormalizedScriptFontFamilies;
  rightToLeft?: boolean;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: ColorValue;
  /** Absent means fully opaque, preserving existing IR v2 snapshots. */
  colorOpacity?: number;
  language: string;
}

export interface ShapeStyleInput {
  fill?: ShapeFillPaintInput;
  stroke?: StrokeStyleInput;
  /** Outer shadow; distance and blur radius are points, direction is clockwise degrees. */
  shadow?: { color: ColorValue; opacity?: number; blur: number; distance: number; direction: number };
}

export interface NormalizedShapeStyle {
  fill: NormalizedShapeFillPaint;
  stroke: NormalizedStrokeStyle;
  shadow?: { color: ColorValue; opacity: number; blur: number; distance: number; direction: number };
}

export type TableCellEdge = "left" | "right" | "top" | "bottom";

export interface TableCellStyleInput {
  fill?: PaintInput;
  stroke?: StrokeStyleInput;
  /** Direct per-edge overrides inherit unspecified stroke properties from the uniform stroke. */
  borders?: Partial<Record<TableCellEdge, StrokeStyleInput>>;
  margin?: number | Partial<Insets>;
  verticalAlign?: "top" | "middle" | "bottom";
  /** Horizontal text or clockwise/counterclockwise quarter-turn in the cell. */
  textDirection?: "horizontal" | "rotate90" | "rotate270" | "eastAsianVertical" | "mongolianVertical" | "stackedVertical" | "stackedVerticalRtl";
}

export interface NormalizedTableCellStyle {
  fill: NormalizedPaint;
  stroke: NormalizedStrokeStyle;
  /** Presence opts the whole table into deterministic shared-border resolution. */
  borders?: Partial<Record<TableCellEdge, NormalizedStrokeStyle>>;
  margin: Insets;
  verticalAlign: "top" | "middle" | "bottom";
  textDirection?: "horizontal" | "rotate90" | "rotate270" | "eastAsianVertical" | "mongolianVertical" | "stackedVertical" | "stackedVerticalRtl";
}

/** A format-independent table region; explicit cell text still takes precedence. */
export interface TableRegionStyleInput extends TableCellStyleInput {
  text?: Pick<TextStylePresetInput, "paragraph" | "run">;
}

export interface TableBandStyleInput {
  /** Positive physical row/column count per band; defaults to 1. */
  size?: number;
  first?: TableRegionStyleInput;
  second?: TableRegionStyleInput;
}

/** Resolved at normalization; no native table-style identifiers enter Core. */
export interface TableStyleInput {
  wholeTable?: TableRegionStyleInput;
  bandedColumns?: TableBandStyleInput;
  bandedRows?: TableBandStyleInput;
  firstColumn?: TableRegionStyleInput;
  lastColumn?: TableRegionStyleInput;
  firstRow?: TableRegionStyleInput;
  lastRow?: TableRegionStyleInput;
  topLeft?: TableRegionStyleInput;
  topRight?: TableRegionStyleInput;
  bottomLeft?: TableRegionStyleInput;
  bottomRight?: TableRegionStyleInput;
  borders?: {
    outer?: StrokeStyleInput;
    innerHorizontal?: StrokeStyleInput;
    innerVertical?: StrokeStyleInput;
  };
}
