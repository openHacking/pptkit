import type { PresentationAsset } from "../types/asset.js";
import type { PresentationSlide } from "../types/presentation.js";

export function assertUniqueAssetIds(assets: PresentationAsset[]): Set<string> {
  const assetIds = new Set<string>();
  for (const asset of assets) {
    if (assetIds.has(asset.id)) throw new Error(`Duplicate asset id "${asset.id}" detected.`);
    assetIds.add(asset.id);
  }
  return assetIds;
}

export function assertUniqueSlideIds(slides: PresentationSlide[]): void {
  const slideIds = new Set<string>();
  for (const slide of slides) {
    if (slideIds.has(slide.id)) throw new Error(`Duplicate slide id "${slide.id}" detected.`);
    slideIds.add(slide.id);
  }
}
