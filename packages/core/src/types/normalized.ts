import type { NormalizedAsset } from "./asset.js";
import type { NormalizedElement } from "./element.js";
import type { PresentationSize } from "./geometry.js";

export interface NormalizedSlide {
  id: string;
  background?: string;
  elements: NormalizedElement[];
}

export interface NormalizedPresentation {
  id: string;
  title?: string;
  size: PresentationSize;
  slides: NormalizedSlide[];
  assets: NormalizedAsset[];
}
