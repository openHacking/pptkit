import type { Insets } from "./geometry.js";

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

export type PaintInput =
  | { type: "none" }
  | { type: "solid"; color: ColorValue; opacity?: number };

export type NormalizedPaint =
  | { type: "none" }
  | { type: "solid"; color: ColorValue; opacity: number };

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
  margin?: number | Partial<Insets>;
  verticalAlign?: "top" | "middle" | "bottom";
  wrap?: boolean;
  autoFit?: TextAutoFitInput;
}

export interface NormalizedTextFrameStyle {
  margin: Insets;
  verticalAlign: "top" | "middle" | "bottom";
  wrap: boolean;
  autoFit: NormalizedTextAutoFit;
}

export type TextBulletInput =
  | { type: "none" }
  | { type: "bullet"; character?: string }
  | { type: "number"; style?: "arabicPeriod" | "arabicParen" | "alphaLowerPeriod" | "alphaUpperPeriod"; startAt?: number };

export type NormalizedTextBullet =
  | { type: "none" }
  | { type: "bullet"; character: string }
  | { type: "number"; style: "arabicPeriod" | "arabicParen" | "alphaLowerPeriod" | "alphaUpperPeriod"; startAt: number };

export interface TextParagraphStyleInput {
  align?: "left" | "center" | "right" | "justify";
  indent?: number;
  hanging?: number;
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  bullet?: TextBulletInput;
}

export interface NormalizedTextParagraphStyle {
  align: "left" | "center" | "right" | "justify";
  indent: number;
  hanging: number;
  lineSpacing: number;
  spaceBefore: number;
  spaceAfter: number;
  bullet: NormalizedTextBullet;
}

export interface TextRunStyleInput {
  fontFamily?: FontFamilyValue;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: ColorValue;
  language?: string;
}

export interface NormalizedTextRunStyle {
  fontFamily: FontFamilyValue;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: ColorValue;
  language: string;
}

export interface ShapeStyleInput {
  fill?: PaintInput;
  stroke?: StrokeStyleInput;
}

export interface NormalizedShapeStyle {
  fill: NormalizedPaint;
  stroke: NormalizedStrokeStyle;
}

export interface TableCellStyleInput {
  fill?: PaintInput;
  stroke?: StrokeStyleInput;
  margin?: number | Partial<Insets>;
  verticalAlign?: "top" | "middle" | "bottom";
}

export interface NormalizedTableCellStyle {
  fill: NormalizedPaint;
  stroke: NormalizedStrokeStyle;
  margin: Insets;
  verticalAlign: "top" | "middle" | "bottom";
}
