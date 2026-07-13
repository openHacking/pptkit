import { normalizePresentation, type PresentationDocument } from "@pptkit/core";

export interface LayoutResult {
  slideCount: number;
  status: "placeholder";
}

export function resolveLayout(document: PresentationDocument): LayoutResult {
  const normalized = normalizePresentation(document);

  return {
    slideCount: normalized.slides.length,
    status: "placeholder",
  };
}
