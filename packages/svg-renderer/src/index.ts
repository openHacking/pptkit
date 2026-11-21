import { normalizePresentation, type PresentationDocument } from "@pptkit/core";
import { resolveNormalizedLayout, type LayoutResult } from "@pptkit/layout";
import { renderResolvedLayout } from "./rendering/render.js";
import type { SvgPresentationRenderOptions, SvgRenderOptions, SvgRenderResult } from "./types/public.js";

export type {
  SvgAssetResolver,
  SvgRenderedSlide,
  SvgSlideReadingOrder,
  SvgEmbeddedFont,
  SvgRenderOptions,
  SvgPresentationRenderOptions,
  SvgRenderResult,
  SvgRenderWarning,
} from "./types/public.js";
export { SvgRenderLimitError, SvgRenderAbortError } from "./types/public.js";
export { getSvgSelectionFrames, getSvgSlideSelectionFrames, getSvgSlideSelectionGeometry, hitTestSvgSelectionFrames, hitTestSvgSelectionGeometry, hitTestSvgSelectionOutlines, hitTestSvgConnectorStrokes, hitTestSvgConnectorMarkers, hitTestSvgShapeStrokes, hitTestSvgSelectionOutlinesAndStrokes } from "./rendering/selection.js";
export type { SvgSelectionFrame, SvgOwnedSelectionFrame, SvgSelectionOwner, SvgSelectionConnector, SvgSelectionGeometry, SvgSelectionGeometryOptions, SvgShapeOutline, SvgConnectorStrokeGeometry, SvgConnectorMarkerGeometry, SvgRoundedShapeStrokeGeometry } from "./rendering/selection.js";
export { createSvgViewportMapping, createSvgAffineMapping } from "./rendering/viewport.js";
export type { SvgViewportMapping, SvgAffineMapping, SvgAffineMatrix } from "./rendering/viewport.js";
export { captureSvgPreviewSnapshot, planSvgPreviewInvalidation } from "./rendering/invalidation.js";
export type { SvgPreviewSnapshotOptions, SvgPreviewSnapshot, SvgPreviewInvalidation } from "./rendering/invalidation.js";
export { renderIncrementalSvg } from "./rendering/incremental.js";
export type { IncrementalSvgRenderOptions, IncrementalSvgCache, IncrementalSvgResult } from "./rendering/incremental.js";
export { renderLayoutToBrowserRaster, BrowserRasterError } from "./rendering/browser-raster.js";
export type { BrowserRasterOptions, BrowserRasterResult, BrowserRasterSlide, BrowserRasterFontCoverageRequest, BrowserRasterStaticImageRequest } from "./rendering/browser-raster.js";
import { checkRenderAbort } from "./rendering/abort.js";

export async function renderPresentationToSvg(
  document: PresentationDocument,
  options?: SvgPresentationRenderOptions,
): Promise<SvgRenderResult> {
  checkRenderAbort(options?.signal);
  return renderResolvedLayout(resolveNormalizedLayout(normalizePresentation(document), options?.layout), options);
}

export async function renderLayoutToSvg(
  layout: LayoutResult,
  options?: SvgRenderOptions,
): Promise<SvgRenderResult> {
  return renderResolvedLayout(layout, options);
}
