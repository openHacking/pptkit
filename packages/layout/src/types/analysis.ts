import type { IntrinsicSizingRule } from "./intrinsic.js";
import type { TablePaginationRule } from "./pagination.js";
import type { TableSizingRule } from "./table.js";
import type { TableBorderLimits } from "./table-borders.js";
import type { LayoutComposition } from "./composition.js";
import type { Box, DateTimeCacheStyle, NormalizedTextParagraphStyle, NormalizedTextRunStyle, PresentationDiagnostic } from "@pptkit/core";

type Immutable<T> = T extends object ? { readonly [K in keyof T]: Immutable<T[K]> } : T;

export interface FontRequest {
  readonly family: string;
  readonly fontSize: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly language: string;
  readonly text: string;
}

export interface ResolvedFont {
  readonly family: string;
}

/** Resolve a requested font, including glyph coverage, without hidden global state. */
export type FontResolver = (request: FontRequest) => ResolvedFont | undefined;

export interface MeasurementRun {
  readonly text: string;
  /** Script slots have already been selected into concrete measurement segments. */
  readonly style: Immutable<Omit<NormalizedTextRunStyle, "fontFamily" | "eastAsianFontFamily" | "complexScriptFontFamily" | "scriptFontFamilies"> & { fontFamily: string }>;
}

export interface MeasurementParagraph {
  readonly runs: readonly MeasurementRun[];
  readonly style: Immutable<NormalizedTextParagraphStyle>;
}

export interface TextMeasurementRequest {
  readonly paragraphs: readonly MeasurementParagraph[];
  /** Content width in logical points, excluding frame margins. May be zero. */
  readonly availableWidth: number;
  readonly wrap: boolean;
}

export interface TextMeasurement {
  /** Content bounds, including paragraph spacing/indentation, before frame alignment. */
  readonly bounds: Readonly<Box>;
}

/** Synchronous provider. Load fonts/resources before resolving Layout. */
export type TextMeasurer = (request: TextMeasurementRequest) => TextMeasurement;

/** East Asian mixed vertical text; coordinates run down inline and leftward across columns. */
export interface EastAsianVerticalMeasurementRequest {
  readonly paragraphs: readonly MeasurementParagraph[];
  readonly availableInlineSize: number;
  readonly availableBlockSize: number;
  readonly wrap: boolean;
}

/** Returns logical bounds: x/width are inline, y/height are block, before frame alignment. */
export type EastAsianVerticalMeasurer = (request: EastAsianVerticalMeasurementRequest) => TextMeasurement;

export type VerticalTextDirection = "eastAsianVertical" | "mongolianVertical" | "stackedVertical" | "stackedVerticalRtl";

/** Logical coordinates follow the mode's inline and column progression. */
export interface VerticalTextMeasurementRequest extends EastAsianVerticalMeasurementRequest {
  readonly textDirection: VerticalTextDirection;
}
export type VerticalTextMeasurer = (request: VerticalTextMeasurementRequest) => TextMeasurement;

export interface TextMeasurementOptions {
  measure: TextMeasurer;
  resolveFont?: FontResolver;
  measureEastAsianVertical?: EastAsianVerticalMeasurer;
  measureVertical?: VerticalTextMeasurer;
}

export interface LayoutDiagnosticOptions {
  textOverflow?: boolean;
  slideBounds?: boolean;
  /** Estimate shape shadow bounds using the SVG preview's three-sigma blur envelope. */
  shadowBounds?: boolean;
  /** Geometric tolerance in points. Default 0.01. */
  tolerance?: number;
  /** Suppress only slide-boundary warnings for these elements and their descendants. */
  allowedOffSlideElementIds?: readonly string[];
  includeHidden?: boolean;
}

export interface LayoutOptions {
  /** Refresh explicit dynamic field caches after pagination using caller-supplied context. */
  fieldRefresh?: { readonly firstPageNumber: number; readonly includeHidden: boolean; readonly instantMs?: number; readonly locale?: string; readonly timeZone?: string; readonly dateStyle?: DateTimeCacheStyle; readonly timeStyle?: DateTimeCacheStyle };
  /** Shared limits across all tables opting into per-edge border resolution. */
  tableBorderLimits?: TableBorderLimits;
  tablePagination?: readonly TablePaginationRule[];
  intrinsicSizing?: readonly IntrinsicSizingRule[];
  tableSizing?: readonly TableSizingRule[];
  composition?: readonly LayoutComposition[];
  textMeasurement?: TextMeasurementOptions;
  diagnostics?: LayoutDiagnosticOptions;
}

export interface LayoutDiagnostic extends PresentationDiagnostic {
  bounds?: Box;
  available?: Box;
}

export interface LayoutTextMeasurement {
  elementId: string;
  path: string;
  slideId?: string;
  layoutId?: string;
  masterId?: string;
  /** Element-local frame area, excluding margins. */
  contentBox: Box;
  /** Element-local measured content after frame margins and vertical alignment. */
  bounds: Box;
}

export interface LayoutAnalysis {
  measurements: LayoutTextMeasurement[];
  diagnostics: LayoutDiagnostic[];
}
