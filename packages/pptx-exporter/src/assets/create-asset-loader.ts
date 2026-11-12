import type { PptxAssetResolver } from "../types/export.js";
import type { AssetLoader } from "../types/internal.js";
import { withAbort } from "./limits.js";
import { inspectJpegExifOrientation } from "./inspect-jpeg-exif.js";
import { loadImageBytes, packageImageMime } from "./load-bytes.js";
import { captureImageInput } from "./capture-image-input.js";

export function createAssetLoader(loadAsset: AssetLoader, resolveAsset: PptxAssetResolver | undefined, signal: AbortSignal | undefined, maxAssetBytes: number | undefined): AssetLoader {
  const resolverContext = Object.freeze({ ...(signal ? { signal } : {}) });
  const resolvedLoader: AssetLoader = (asset, context) => withAbort((async () => {
    signal?.throwIfAborted();
    if (!resolveAsset) return loadAsset(asset, context);
    const request = structuredClone(asset);
    Object.freeze(request.source); Object.freeze(request.accessibility); Object.freeze(request);
    const resolved = await resolveAsset(request, resolverContext);
    signal?.throwIfAborted();
    if (!resolved) throw new Error("The asset resolver returned no asset; default loading is disabled when a resolver is supplied.");
    const { bytes, mimeType, animationFallback, orientationNormalization, vectorRasterization, svgPair } = resolved;
    if (animationFallback !== undefined && animationFallback !== "first-frame") throw new TypeError("Resolver animationFallback must be first-frame when provided.");
    if (orientationNormalization !== undefined && orientationNormalization !== "pixels") throw new TypeError("Resolver orientationNormalization must be pixels when provided.");
    if (orientationNormalization !== undefined && animationFallback !== undefined) throw new TypeError("Orientation normalization and animation fallback declarations cannot be combined.");
    if (orientationNormalization !== undefined && asset.mimeType !== undefined && asset.mimeType.trim().toLowerCase() !== "image/jpeg") throw new TypeError("Orientation normalization requires a JPEG source MIME when declared.");
    if (vectorRasterization !== undefined && vectorRasterization !== "emf" && vectorRasterization !== "wmf" && vectorRasterization !== "svg") throw new TypeError("Resolver vectorRasterization must be emf, wmf or svg when provided.");
    if (vectorRasterization !== undefined && (orientationNormalization !== undefined || animationFallback !== undefined)) throw new TypeError("Vector rasterization cannot be combined with other conversion declarations.");
    if (vectorRasterization !== undefined && asset.mimeType?.trim().toLowerCase() !== (vectorRasterization === "svg" ? "image/svg+xml" : `image/${vectorRasterization}`)) throw new TypeError("Vector rasterization requires a matching declared source MIME.");
    if (svgPair !== undefined && (orientationNormalization !== undefined || animationFallback !== undefined || vectorRasterization !== undefined)) throw new TypeError("SVG pairing cannot be combined with conversion declarations.");
    if (svgPair !== undefined && asset.mimeType?.trim().toLowerCase() !== "image/svg+xml") throw new TypeError("SVG pairing requires a declared image/svg+xml source MIME.");
    if (!(bytes instanceof Uint8Array)) throw new TypeError("Asset resolver bytes must be Uint8Array.");
    const converted = orientationNormalization !== undefined || vectorRasterization !== undefined || svgPair !== undefined;
    const loaded = loadImageBytes(bytes, converted ? mimeType : mimeType ?? asset.mimeType, maxAssetBytes);
    if (!converted && mimeType !== undefined && asset.mimeType !== undefined && packageImageMime(mimeType).trim().toLowerCase() !== packageImageMime(asset.mimeType).trim().toLowerCase()) throw new Error("Resolver MIME differs from declared asset MIME.");
    if (vectorRasterization !== undefined && loaded.mimeType !== "image/png" && loaded.mimeType !== "image/jpeg") throw new TypeError("Vector rasterization output must be PNG or JPEG.");
    if (svgPair !== undefined && loaded.mimeType !== "image/png" && loaded.mimeType !== "image/jpeg") throw new TypeError("SVG fallback must be PNG or JPEG.");
    const capturedSvg = svgPair === undefined ? undefined : captureImageInput(svgPair, "SVG pair", (maxAssetBytes ?? Number.MAX_SAFE_INTEGER) - loaded.data.byteLength, signal, "remaining maxAssetBytes");
    if (capturedSvg !== undefined && capturedSvg.byteLength === 0) throw new TypeError("SVG pair must not be empty.");
    if (animationFallback !== undefined && loaded.mimeType !== "image/png" && loaded.mimeType !== "image/jpeg") throw new TypeError("Resolver first-frame fallback must use PNG or JPEG output.");
    if (orientationNormalization !== undefined) {
      if (loaded.mimeType !== "image/png" && loaded.mimeType !== "image/jpeg") throw new TypeError("Normalized orientation output must be PNG or JPEG.");
      if (loaded.mimeType === "image/jpeg") {
        const orientation = inspectJpegExifOrientation(loaded.data, { ...(signal ? { signal } : {}) }).orientation;
        if (orientation !== undefined && orientation !== 1) throw new TypeError("Normalized JPEG retains a non-default EXIF orientation.");
      }
    }
    return { ...loaded, ...(animationFallback === undefined ? {} : { animationFallback }), ...(orientationNormalization === undefined ? {} : { orientationNormalization }), ...(vectorRasterization === undefined ? {} : { vectorRasterization }), ...(capturedSvg === undefined ? {} : { svgPair: capturedSvg }) };
  })(), signal);
  return resolvedLoader;
}
