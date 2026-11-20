export { orientJpegCrop } from "./assets/orient-jpeg-crop.js";
export { inspectJpegExifOrientation } from "./assets/inspect-jpeg-exif.js";
export { inspectJpegIcc } from "./assets/inspect-jpeg-icc.js";
export { extractJpegIccProfile } from "./assets/inspect-jpeg-icc.js";
export type { ExtractedJpegIcc } from "./assets/inspect-jpeg-icc.js";
export { inspectPngIcc } from "./assets/inspect-png-icc.js";
export { inspectIccProfile } from "./assets/inspect-icc-profile.js";
export type { IccProfileInspection, IccProfileInspectionOptions, IccProfileTag } from "./assets/inspect-icc-profile.js";
export type { PngIccInspection, PngIccInspectionOptions } from "./assets/inspect-png-icc.js";
export type { JpegIccInspection, JpegIccInspectionOptions } from "./assets/inspect-jpeg-icc.js";
export type { JpegExifOrientationInspection, JpegExifInspectionOptions } from "./assets/inspect-jpeg-exif.js";
export { inspectGif } from "./assets/inspect-gif.js";
export { inspectPng } from "./assets/inspect-png.js";
export type { PngInspection, PngInspectionOptions } from "./assets/inspect-png.js";
export { inspectApng } from "./assets/inspect-apng.js";
export type { ApngFrameInfo, ApngInspection, ApngInspectionOptions } from "./assets/inspect-apng.js";
export type { GifFrameInfo, GifInspection, GifInspectionOptions } from "./assets/inspect-gif.js";
export { PptxExportError } from "./types/export.js";
export { createHttpMediaResolver } from "./adapters/http-media-resolver.js";
export type { HttpMediaResolverOptions } from "./adapters/http-media-resolver.js";
import type { PresentationDocument } from "@pptkit/core";
import { loadAsset } from "./assets/load-asset.js";
import { generatePptxWith } from "./generation/generate-pptx.js";
import type { GeneratePptxOptions, GeneratePptxResult } from "./types/export.js";

export type { PptxAssetResolverRequest, PptxAssetResolver, PptxAssetResolverContext, PptxMediaResolver, ResolvedPptxAsset, ExportWarning, GeneratePptxOptions, GeneratePptxResult } from "./types/export.js";

export async function generatePptx(document: PresentationDocument, options?: GeneratePptxOptions): Promise<GeneratePptxResult> {
  return generatePptxWith(document, loadAsset, options);
}
