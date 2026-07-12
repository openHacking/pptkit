import type { PresentationDocument } from "@pptkit/core";

export interface LayoutResult {
  slideCount: number;
  status: "placeholder";
}

export function resolveLayout(document: PresentationDocument): LayoutResult {
  return {
    slideCount: document.slides.length,
    status: "placeholder",
  };
}

