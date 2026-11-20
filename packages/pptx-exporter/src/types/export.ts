import type { NormalizedAsset } from "@pptkit/core";
import type { LayoutOptions } from "@pptkit/layout";

export interface ExportWarning {
  code: string;
  message: string;
  slideId?: string;
  elementIndex?: number;
  assetId?: string;
  elementId?: string;
  stageId?: string;
  layoutId?: string;
  masterId?: string;
  path?: string;
}

export type PptxAssetResolverRequest = Readonly<Omit<NormalizedAsset, "source" | "accessibility">> & { readonly source: Readonly<NormalizedAsset["source"]>; readonly accessibility: Readonly<NormalizedAsset["accessibility"]> };
export interface ResolvedPptxAsset {
  bytes: Uint8Array;
  mimeType?: string;
  /** Resolver reports that animated input was flattened to its first frame. */
  animationFallback?: "first-frame";
  /** Resolver reports that EXIF-oriented JPEG pixels were normalized. */
  orientationNormalization?: "pixels";
  /** Resolver asserts it rasterized a vector source; fidelity is caller-owned. */
  vectorRasterization?: "emf" | "wmf" | "svg";
  /** SVG bytes paired with the PNG/JPEG bytes above; correspondence is caller-owned. */
  svgPair?: Uint8Array;
}
export interface PptxAssetResolverContext { readonly signal?: AbortSignal }
/** Undefined explicitly omits an asset; no implicit fallback loading follows. */
export type PptxAssetResolver = (asset: PptxAssetResolverRequest, context: PptxAssetResolverContext) => ResolvedPptxAsset | undefined | Promise<ResolvedPptxAsset | undefined>;

export interface GeneratePptxOptions {
  /** Experimental byte-backed PCM WAVE with visible top-level audio objects. */
  audioMode?: "wave-pcm-v1";
  /** Experimental byte-backed MP4 framing and movie-header profile; codecs are not validated. */
  videoMode?: "iso-bmff-v1";
  layout?: LayoutOptions;
  /** Native presentation starting slide number (0–9999); independent of cached field refresh. */
  firstSlideNumber?: number;
  /** Opt in to live native placeholder run inheritance; may change shared text metrics. */
  nativePlaceholderText?: "materialize" | "inherit";
  /** Experimental native straight connectors; unsupported cases retain paths with diagnostics. */
  connectorMode?: "path" | "native";
  /** Native event timing is unavailable. Only explicit omit permits static output. */
  eventTimelineMode?: "reject" | "omit";
  /** Warn about GIF playback or omit GIFs with invalid bounded image data. */
  gifDiagnostics?: "off" | "warn" | "reject-lzw";
  /** Warn or omit a PNG/APNG asset when its bounded structure cannot be verified. */
  pngDiagnostics?: "off" | "warn" | "reject-structure" | "reject-scanlines";
  /** Opt in to pre-scan EXIF orientation diagnostics without changing image bytes. */
  jpegOrientationDiagnostics?: "off" | "warn";
  /** Opt in to bounded pre-scan JPEG ICC chunk and color-fidelity diagnostics. */
  jpegIccDiagnostics?: "off" | "warn";
  /** Require the bounded standalone Huffman DCT JPEG structure before packaging. */
  jpegValidation?: "structure";
  /** Share identical resolved media parts within this export; Core asset identities stay separate. */
  deduplicateResolvedAssets?: boolean;
  /** Require supported PNG/JPEG/GIF signatures for built-in URL and Node path loads. */
  externalImageValidation?: "signature";
  resolveAsset?: PptxAssetResolver;
  signal?: AbortSignal;
  /** Optional positive byte cap per asset; includes default URL/path loading. */
  maxAssetBytes?: number;
}

export interface GeneratePptxResult {
  bytes: Uint8Array;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "generated" | "generated-with-warnings";
}

export interface WritePptxOptions extends GeneratePptxOptions {
  output: string;
}

export interface WritePptxResult {
  output: string;
  slideCount: number;
  byteLength: number;
  warnings: ExportWarning[];
  status: "written" | "written-with-warnings";
}

/** Rejects unsupported authored timing before assets or files are produced. */
export class PptxExportError extends Error {
  readonly code = "event-timeline-export-unsupported";
  readonly diagnostics: readonly Readonly<ExportWarning>[];
  constructor(diagnostics: readonly ExportWarning[]) {
    super("Native event timeline output is not supported; explicitly select eventTimelineMode: omit for static output with diagnostics.");
    this.name = "PptxExportError";
    this.diagnostics = Object.freeze(diagnostics.map(diagnostic => Object.freeze({ ...diagnostic })));
  }
}
