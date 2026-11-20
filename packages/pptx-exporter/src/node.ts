export { orientJpegCrop } from "./assets/orient-jpeg-crop.js";
export { inspectJpegExifOrientation } from "./assets/inspect-jpeg-exif.js";
export { inspectJpegIcc } from "./assets/inspect-jpeg-icc.js";
export { extractJpegIccProfile } from "./assets/inspect-jpeg-icc.js";
export type { ExtractedJpegIcc } from "./assets/inspect-jpeg-icc.js";
export { inspectPngIcc } from "./assets/inspect-png-icc.js";
export { inspectIccProfile } from "./assets/inspect-icc-profile.js";
export type { IccProfileInspection, IccProfileInspectionOptions, IccProfileTag } from "./assets/inspect-icc-profile.js";
export { extractPngIccProfile } from "./adapters/node-png-icc.js";
export type { ExtractPngIccOptions, ExtractedPngIcc } from "./adapters/node-png-icc.js";
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
export { createFileMediaResolver } from "./adapters/node-file-media-resolver.js";
export type { FileMediaResolverOptions } from "./adapters/node-file-media-resolver.js";
import type { PresentationDocument } from "@pptkit/core";
import { loadNodeAsset } from "./assets/load-asset.node.js";
import { generatePptxWith } from "./generation/generate-pptx.js";
import { writeOutput } from "./output/write-output.js";
import type { GeneratePptxOptions, GeneratePptxResult, WritePptxOptions, WritePptxResult } from "./types/export.js";

export type { PptxAssetResolverRequest, PptxAssetResolver, PptxAssetResolverContext, PptxMediaResolver, ResolvedPptxAsset, ExportWarning, GeneratePptxOptions, GeneratePptxResult, WritePptxOptions, WritePptxResult } from "./types/export.js";

export async function generatePptx(document: PresentationDocument, options?: GeneratePptxOptions): Promise<GeneratePptxResult> {
  return generatePptxWith(document, loadNodeAsset, options);
}

export async function writePptx(document: PresentationDocument, options: WritePptxOptions): Promise<WritePptxResult> {
  const generated = await generatePptx(document, options);
  const written = await writeOutput(options.output, generated.bytes);
  return {
    output: written.output,
    slideCount: generated.slideCount,
    byteLength: written.byteLength,
    warnings: generated.warnings,
    status: generated.warnings.length === 0 ? "written" : "written-with-warnings",
  };
}
