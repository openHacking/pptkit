import type { NormalizedElement, PresentationSize } from "@pptkit/core";

export interface LayoutSlide {
  id: string;
  elements: NormalizedElement[];
}

export interface LayoutResult {
  size: PresentationSize;
  slides: LayoutSlide[];
  slideCount: number;
  status: "resolved";
}
