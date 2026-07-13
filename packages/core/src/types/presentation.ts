import type { PresentationAsset, PresentationAssetInput } from "./asset.js";
import type { PresentationElementInput } from "./element.js";
import type { PresentationSize } from "./geometry.js";

export interface PresentationSlideInput {
  id?: string;
  elements?: PresentationElementInput[];
}

export interface PresentationInit {
  id?: string;
  title?: string;
  size?: Partial<PresentationSize>;
}

export interface PresentationSlide {
  id: string;
  elements: PresentationElementInput[];
}

export interface PresentationDocument {
  id: string;
  title?: string;
  size: PresentationSize;
  slides: PresentationSlide[];
  assets: PresentationAsset[];
  addSlide(input?: PresentationSlideInput): PresentationSlide;
  registerAsset(input: PresentationAssetInput): PresentationAsset;
  getAsset(assetId: string): PresentationAsset | undefined;
}
