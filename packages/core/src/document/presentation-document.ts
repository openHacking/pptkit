import { DEFAULT_PRESENTATION_SIZE } from "../constants/presentation.js";
import type { PresentationAsset, PresentationAssetInput } from "../types/asset.js";
import type { PresentationSize } from "../types/geometry.js";
import type { PresentationDocument, PresentationInit, PresentationSlide, PresentationSlideInput } from "../types/presentation.js";
import { cloneSlide } from "../utils/clone.js";
import { createDocumentId, createSlideId } from "../utils/id.js";
import { assertUniqueSlideIds } from "../validation/collection.js";
import { normalizeSize } from "../validation/geometry.js";
import { AssetRegistry } from "./asset-registry.js";

export class PresentationDocumentImpl implements PresentationDocument {
  readonly id: string;
  readonly size: PresentationSize;
  readonly slides: PresentationSlide[] = [];
  readonly assets: PresentationAsset[];
  title?: string;
  private readonly assetRegistry = new AssetRegistry();

  constructor(init: PresentationInit = {}) {
    this.id = init.id ?? createDocumentId();
    this.size = normalizeSize(init.size, DEFAULT_PRESENTATION_SIZE);
    this.assets = this.assetRegistry.assets;
    if (init.title !== undefined) this.title = init.title;
  }

  addSlide(input: PresentationSlideInput = {}): PresentationSlide {
    const slideId = input.id ?? createSlideId(this.slides.length);
    assertUniqueSlideIds([...this.slides, { id: slideId, elements: [] }]);
    const slide = cloneSlide({ id: slideId, elements: input.elements ?? [] });
    this.slides.push(slide);
    return slide;
  }

  registerAsset(input: PresentationAssetInput): PresentationAsset {
    return this.assetRegistry.register(input);
  }

  getAsset(assetId: string): PresentationAsset | undefined {
    return this.assetRegistry.get(assetId);
  }
}
