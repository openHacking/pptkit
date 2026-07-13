import type { PresentationAsset, PresentationAssetInput, PresentationAssetSource } from "../types/asset.js";
import type { PresentationElementInput } from "../types/element.js";
import type { Box } from "../types/geometry.js";
import type { PresentationSlide, PresentationSlideInput } from "../types/presentation.js";
import type { ShapeStyle, TextStyle } from "../types/style.js";

export function cloneBox(box: Box): Box {
  return { x: box.x, y: box.y, width: box.width, height: box.height };
}

export function cloneAssetSource(source: PresentationAssetSource): PresentationAssetSource {
  return { type: source.type, value: source.value };
}

export function cloneTextStyle(style: TextStyle | undefined): TextStyle {
  return style === undefined ? {} : { ...style };
}

export function cloneShapeStyle(style: ShapeStyle | undefined): ShapeStyle {
  return style === undefined ? {} : { ...style };
}

export function cloneAsset(input: PresentationAssetInput | PresentationAsset): PresentationAsset {
  return {
    id: input.id ?? "",
    kind: input.kind,
    source: cloneAssetSource(input.source),
    ...(input.mimeType !== undefined ? { mimeType: input.mimeType } : {}),
    ...(input.width !== undefined ? { width: input.width } : {}),
    ...(input.height !== undefined ? { height: input.height } : {}),
    ...(input.altText !== undefined ? { altText: input.altText } : {}),
    ...(input.dedupeKey !== undefined ? { dedupeKey: input.dedupeKey } : {}),
  };
}

export function cloneElement(element: PresentationElementInput): PresentationElementInput {
  if (element.type === "text") {
    return {
      type: "text", text: element.text, box: cloneBox(element.box),
      ...(element.style !== undefined ? { style: cloneTextStyle(element.style) } : {}),
    };
  }
  if (element.type === "image") {
    return {
      type: "image", assetId: element.assetId, box: cloneBox(element.box),
      ...(element.altText !== undefined ? { altText: element.altText } : {}),
    };
  }
  return {
    type: "shape", shape: element.shape, box: cloneBox(element.box),
    ...(element.style !== undefined ? { style: cloneShapeStyle(element.style) } : {}),
  };
}

export function cloneSlide(input: PresentationSlideInput | PresentationSlide): PresentationSlide {
  return { id: input.id ?? "", elements: (input.elements ?? []).map(cloneElement) };
}
