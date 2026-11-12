import type { PptxSourceLocation } from "./source.js";
import type { TextAutoFitInput, TextBulletInput, TextTabStop, PaintInput, StrokeStyleInput, TextFontScriptCode } from "@pptkit/core";

export type PptxPlaceholderLevel = "slide" | "layout" | "master";
export type PptxPlaceholderDimension = "x" | "y" | "width" | "height";

export interface PptxPlaceholderMatch {
  readonly rule: "index" | "type";
  readonly key: number | string;
  readonly status: "matched" | "missing" | "ambiguous" | "unsupported";
  readonly part?: string;
  /** Count only; ambiguous candidate lists are not expanded for every child. */
  readonly candidateCount: number;
  readonly source?: PptxSourceLocation;
  readonly reason?: string;
}

export interface PptxInheritedDimension {
  /** Original EMUs, before conversion to Core points. Zero is an explicit value. */
  readonly valueEmu: number;
  readonly source: PptxSourceLocation;
  readonly level: PptxPlaceholderLevel;
}

export interface PptxPlaceholderInfo {
  readonly source: PptxSourceLocation;
  readonly declaration: PptxSourceLocation;
  readonly level: PptxPlaceholderLevel;
  readonly type: string;
  readonly index: number;
  readonly explicitType: boolean;
  readonly explicitIndex: boolean;
  readonly grouped: boolean;
  /** Masters have no placeholder parent. */
  readonly parent?: PptxPlaceholderMatch;
  /** Resolved dimensions only; this is not a complete shape/style cascade. */
  readonly dimensions: Readonly<Partial<Record<PptxPlaceholderDimension, PptxInheritedDimension>>>;
  readonly geometryStatus: "resolved" | "incomplete" | "unsupported";
  readonly reason?: string;
}

/** Source provenance owned by Parser; not a Core template or source save contract. */
export interface PptxPlaceholderInheritance {
  /** Each declaration in a visited slide/layout/master part appears once. */
  readonly placeholders: readonly PptxPlaceholderInfo[];
}

export type PptxTextStyleLevel = "notes-master-text-style" | "notes-master-placeholder" | "notes-direct" | "table-style" | "presentation-default" | "master-text-style" | "master-placeholder" | "layout-placeholder" | "slide-list-style" | "paragraph" | "run" | "end-paragraph" | "slide-frame";
export type PptxTextStyleProperty =
  | "frame.margin.left" | "frame.margin.right" | "frame.margin.top" | "frame.margin.bottom"
  | "frame.textDirection" | "frame.verticalAlign" | "frame.wrap" | "frame.autoFit"
  | "paragraph.align" | "paragraph.direction" | "paragraph.indent" | "paragraph.hanging" | "paragraph.lineSpacing"
  | "paragraph.spaceBefore" | "paragraph.spaceAfter" | "paragraph.bullet" | "paragraph.tabStops"
  | "run.fontFamily" | "run.eastAsianFontFamily" | "run.complexScriptFontFamily" | "run.fontSize" | "run.bold" | "run.italic" | "run.underline"
  | `run.scriptFontFamilies.${TextFontScriptCode}`
  | "run.strike" | "run.color" | "run.colorOpacity" | "run.language" | "run.rightToLeft";
export type PptxTextStyleValue = string | number | boolean | Readonly<TextAutoFitInput> | Readonly<TextBulletInput> | readonly Readonly<TextTabStop>[];
export interface PptxTextStylePropertyInfo {
  readonly property: PptxTextStyleProperty;
  readonly source: PptxSourceLocation;
  /** Local destination body, paragraph or run; never a parent prompt's text. */
  readonly target: PptxSourceLocation;
  readonly level: PptxTextStyleLevel;
  readonly status: "mapped" | "unsupported";
  readonly value?: PptxTextStyleValue;
  /** Theme palette/font or color-map dependencies used to derive the value. */
  readonly resolvedFrom?: readonly PptxSourceLocation[];
  readonly reason?: string;
}
export interface PptxTextStyleInheritance {
  /** Selected declarations only; absent properties still use projection defaults. */
  readonly properties: readonly PptxTextStylePropertyInfo[];
}


export interface PptxTableStyleSelection {
  readonly flags: Readonly<Record<"firstRow" | "lastRow" | "firstCol" | "lastCol" | "bandRow" | "bandCol", boolean>>;
  readonly source: PptxSourceLocation;
  readonly styleId?: string;
  readonly styleName?: string;
  readonly definition?: PptxSourceLocation;
  readonly status: "inline" | "resolved" | "missing" | "unsupported";
  readonly reason?: string;
}
export interface PptxTableStylePropertyInfo {
  readonly level?: PptxTextStyleLevel;
  readonly property: "cell.fill" | `cell.border.${"left" | "right" | "top" | "bottom"}` | PptxTextStyleProperty;
  readonly source: PptxSourceLocation;
  readonly target: PptxSourceLocation;
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly region?: string;
  readonly status: "mapped" | "unsupported";
  readonly value?: PptxTextStyleValue | Readonly<PaintInput> | Readonly<StrokeStyleInput>;
  readonly resolvedFrom?: readonly PptxSourceLocation[];
  readonly reason?: string;
}
export interface PptxTableStyleInheritance {
  readonly profile: "drawingml-direct-regions-v1";
  readonly tables: readonly PptxTableStyleSelection[];
  readonly properties: readonly PptxTableStylePropertyInfo[];
}
