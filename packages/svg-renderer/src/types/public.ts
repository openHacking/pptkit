import type { NormalizedAsset } from "@pptkit/core";
import type { LayoutOptions } from "@pptkit/layout";

export interface SvgRenderedSlide {
  slideId: string;
  index: number;
  hidden: boolean;
  svg: string;
}

export interface SvgRenderWarning {
  code: string;
  message: string;
  slideId?: string;
  elementId?: string;
  assetId?: string;
  layoutId?: string;
  masterId?: string;
  path?: string;
  /** Declared families and styles emitted for this object, not verified installed faces. */
  fontDependencies?: { family: string; bold: boolean; italic: boolean }[];
}

export type SvgAssetResolver = (
  asset: NormalizedAsset,
  context?: { readonly signal?: AbortSignal },
) => string | undefined | Promise<string | undefined>;

/** A permutation of a visible slide's or group's visible, non-decorative direct children. */
export interface SvgSlideReadingOrder {
  readonly slideId: string;
  readonly groupId?: string;
  readonly elementIds: readonly string[];
}

/** Caller-supplied, licensed font program with an explicit CSS family and face style. */
export interface SvgEmbeddedFont {
  family: string;
  weight: 400 | 700;
  italic: boolean;
  format: "woff2" | "woff" | "truetype" | "opentype";
  base64: string;
}

export interface SvgRenderOptions {
  /** Cooperative cancellation, including while awaiting an asset resolver. */
  signal?: AbortSignal;
  /** Maximum cumulative UTF-8 bytes of rendered SVG strings, excluding warnings and layout. */
  maxSvgBytes?: number;
  /** Zero-based slide positions, rendered in the supplied order; omitted renders every slide. */
  slideIndices?: readonly number[];
  /** Distinguishes independently rendered SVGs placed in the same DOM. */
  idNamespace?: string;
  /** Experimental SVG accessibility order, separate from drawing and keyboard focus. */
  readingOrder?: readonly SvgSlideReadingOrder[];
  resolveAsset?: SvgAssetResolver;
  /** Opt in to a bounded, non-validating APNG declaration scan for byte assets. */
  pngAnimationDiagnostics?: "off" | "warn";
  /** Report selected dependencies that can prevent standalone SVG viewing. */
  portabilityDiagnostics?: "off" | "warn";
  /** Opt into native SVG for eligible mixed-run and shape text. */
  mixedTextMode?: "browser" | "native-no-wrap" | "native";
  /** Opt into approximate native SVG tables where their content is supported. */
  tableMode?: "browser" | "native";
  /** Embed supplied font programs into each SVG; no filesystem or network is accessed. */
  embeddedFonts?: readonly SvgEmbeddedFont[];
  /** Omit the slide background rectangle for transparent export. */
  backgroundMode?: "slide" | "transparent";
  /** Link slide actions between selected SVGs embedded together in one document. */
  slideActionMode?: "metadata" | "document-links";
}

export class SvgRenderLimitError extends Error {
  readonly code = "svg-limit-exceeded";
  constructor(readonly limit: number, readonly slideIndex: number) {
    super("Rendered SVG bytes exceed maxSvgBytes.");
    this.name = "SvgRenderLimitError";
  }
}

export class SvgRenderAbortError extends Error {
  readonly code = "operation-aborted";
  constructor() { super("SVG rendering was aborted."); this.name = "SvgRenderAbortError"; }
}

export interface SvgPresentationRenderOptions extends SvgRenderOptions {
  layout?: LayoutOptions;
}

export interface SvgRenderResult {
  width: number;
  height: number;
  slides: SvgRenderedSlide[];
  warnings: SvgRenderWarning[];
  status: "rendered" | "rendered-with-warnings";
}
