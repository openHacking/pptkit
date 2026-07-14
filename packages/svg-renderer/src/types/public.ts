import type { NormalizedAsset } from "@pptkit/core";

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
}

export type SvgAssetResolver = (
  asset: NormalizedAsset,
) => string | undefined | Promise<string | undefined>;

export interface SvgRenderOptions {
  resolveAsset?: SvgAssetResolver;
}

export interface SvgRenderResult {
  width: number;
  height: number;
  slides: SvgRenderedSlide[];
  warnings: SvgRenderWarning[];
  status: "rendered" | "rendered-with-warnings";
}
