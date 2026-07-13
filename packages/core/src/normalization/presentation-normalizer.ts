import type { PresentationAsset } from "../types/asset.js";
import type { NormalizedElement } from "../types/element.js";
import type { NormalizedPresentation } from "../types/normalized.js";
import type { PresentationDocument, PresentationSlide } from "../types/presentation.js";
import { cloneAsset, cloneBox, cloneShapeStyle, cloneTextStyle } from "../utils/clone.js";
import { assertUniqueAssetIds, assertUniqueSlideIds } from "../validation/collection.js";
import { assertBox, assertOptionalDimension, normalizeSize } from "../validation/geometry.js";
import { assertTextStyle } from "../validation/style.js";
import { DEFAULT_PRESENTATION_SIZE } from "../constants/presentation.js";

function normalizeAsset(asset: PresentationAsset): PresentationAsset {
  assertOptionalDimension(asset.width, `Asset "${asset.id}" width`);
  assertOptionalDimension(asset.height, `Asset "${asset.id}" height`);
  return cloneAsset(asset);
}

function normalizeElement(element: PresentationSlide["elements"][number], assetIds: Set<string>, slideId: string, elementIndex: number): NormalizedElement {
  const context = `Slide "${slideId}" element ${elementIndex + 1}`;
  assertBox(element.box, `${context} box`);
  if (element.type === "text") {
    assertTextStyle(element.style, `${context} text style`);
    return { type: "text", text: element.text, box: cloneBox(element.box), style: cloneTextStyle(element.style) };
  }
  if (element.type === "image") {
    if (!assetIds.has(element.assetId)) {
      throw new Error(`Slide "${slideId}" image element references missing asset "${element.assetId}".`);
    }
    return {
      type: "image", assetId: element.assetId, box: cloneBox(element.box),
      ...(element.altText !== undefined ? { altText: element.altText } : {}),
    };
  }
  return { type: "shape", shape: element.shape, box: cloneBox(element.box), style: cloneShapeStyle(element.style) };
}

export class PresentationNormalizer {
  normalize(document: PresentationDocument): NormalizedPresentation {
    const assetIds = assertUniqueAssetIds(document.assets);
    assertUniqueSlideIds(document.slides);
    return {
      id: document.id,
      ...(document.title !== undefined ? { title: document.title } : {}),
      size: normalizeSize(document.size, DEFAULT_PRESENTATION_SIZE),
      assets: document.assets.map(normalizeAsset),
      slides: document.slides.map((slide) => ({
        id: slide.id,
        ...(slide.background !== undefined ? { background: slide.background } : {}),
        elements: slide.elements.map((element, index) => normalizeElement(element, assetIds, slide.id, index)),
      })),
    };
  }
}
