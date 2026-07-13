import type { NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import type { LayoutResult } from "../types/layout.js";

function cloneElement(element: NormalizedElement): NormalizedElement {
  return {
    ...element,
    box: { ...element.box },
    ...(element.type === "text" || element.type === "shape" ? { style: { ...element.style } } : {}),
  } as NormalizedElement;
}

export function resolveNormalizedLayout(normalized: NormalizedPresentation): LayoutResult {
  return {
    size: { ...normalized.size },
    slides: normalized.slides.map((slide) => ({
      id: slide.id,
      elements: slide.elements.map(cloneElement),
    })),
    slideCount: normalized.slides.length,
    status: "resolved",
  };
}
