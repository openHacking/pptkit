import type { ImageElementInput, PresentationAssetInput } from "@pptkit/core";

/** Caller-supplied rasterization; no SVG parsing, fetching or rasterization is performed. */
export interface SvgRasterFallbackInput {
  readonly sourceId: string;
  readonly assetId: string;
  readonly bytes: Uint8Array;
  readonly mimeType: "image/png" | "image/jpeg";
  /** Intended viewport in Core points; the image is stretched to this frame. */
  readonly width: number;
  readonly height: number;
}

export interface SvgRasterFallback {
  readonly status: "raster-fallback";
  readonly sourceId: string;
  readonly asset: PresentationAssetInput;
  readonly element: ImageElementInput;
  readonly issue: { readonly code: "rasterized"; readonly message: string };
}

/** The caller is responsible for raster trust, source identity and visual fidelity. */
export function createSvgRasterFallback(input: SvgRasterFallbackInput): SvgRasterFallback {
  if (!input || typeof input.sourceId !== "string" || !input.sourceId.trim() || input.sourceId.length > 1024 ||
      typeof input.assetId !== "string" || !input.assetId.trim() || input.assetId.length > 1024 ||
      !(input.bytes instanceof Uint8Array) || input.bytes.length === 0 || input.bytes.length > 20_000_000 ||
      (input.mimeType !== "image/png" && input.mimeType !== "image/jpeg") ||
      !Number.isFinite(input.width) || input.width <= 0 || input.width > 1_000_000 ||
      !Number.isFinite(input.height) || input.height <= 0 || input.height > 1_000_000) {
    throw new TypeError("Invalid SVG raster fallback input.");
  }
  const asset: PresentationAssetInput = { id: input.assetId, kind: "image", mimeType: input.mimeType, source: { type: "bytes", value: new Uint8Array(input.bytes), encoding: "raw" } };
  const element: ImageElementInput = { type: "image", assetId: input.assetId, box: { x: 0, y: 0, width: input.width, height: input.height }, fit: "stretch" };
  return { status: "raster-fallback", sourceId: input.sourceId, asset, element, issue: { code: "rasterized", message: "Caller-supplied raster replaces the source; individual SVG objects are not editable and fidelity is not verified." } };
}
