import { normalizePresentation, type PresentationDocument } from "@pptkit/core";
import { resolveNormalizedLayout, type LayoutResult } from "@pptkit/layout";
import { renderResolvedLayout } from "./rendering/render.js";
import type { SvgRenderOptions, SvgRenderResult } from "./types/public.js";

export type {
  SvgAssetResolver,
  SvgRenderedSlide,
  SvgRenderOptions,
  SvgRenderResult,
  SvgRenderWarning,
} from "./types/public.js";

export async function renderPresentationToSvg(
  document: PresentationDocument,
  options?: SvgRenderOptions,
): Promise<SvgRenderResult> {
  return renderResolvedLayout(resolveNormalizedLayout(normalizePresentation(document)), options);
}

export async function renderLayoutToSvg(
  layout: LayoutResult,
  options?: SvgRenderOptions,
): Promise<SvgRenderResult> {
  return renderResolvedLayout(layout, options);
}
