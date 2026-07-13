import type {
  NormalizedAsset,
  NormalizedConnectorElement,
  NormalizedElement,
  NormalizedGroupElement,
  NormalizedImageElement,
  NormalizedPaint,
  NormalizedPlaceholderDefinition,
  NormalizedPresentationMetadata,
  NormalizedPresentationTheme,
  NormalizedTextParagraph,
  Point,
  PresentationSize,
  JsonValue,
} from "@pptkit/core";

export interface LayoutConnectorElement extends Omit<NormalizedConnectorElement, "start" | "end"> {
  start: Point;
  end: Point;
}

export interface LayoutGroupElement extends Omit<NormalizedGroupElement, "children"> {
  children: LayoutElement[];
}

export type LayoutElement =
  | Exclude<NormalizedElement, NormalizedConnectorElement | NormalizedGroupElement>
  | LayoutConnectorElement
  | LayoutGroupElement;

export interface LayoutSlideLayout {
  id: string;
  name: string;
  background: NormalizedPaint;
  elements: LayoutElement[];
  placeholders: NormalizedPlaceholderDefinition[];
}

export interface LayoutSlide {
  id: string;
  layoutId: string;
  background: NormalizedPaint;
  backgroundSource: "slide" | "layout" | "theme";
  elements: LayoutElement[];
  notes: NormalizedTextParagraph[];
  hidden: boolean;
  section?: string;
  tags: string[];
  customData: Record<string, JsonValue>;
}

export interface LayoutResult {
  size: PresentationSize;
  metadata: NormalizedPresentationMetadata;
  theme: NormalizedPresentationTheme;
  assets: NormalizedAsset[];
  layouts: LayoutSlideLayout[];
  slides: LayoutSlide[];
  slideCount: number;
  status: "resolved";
}
