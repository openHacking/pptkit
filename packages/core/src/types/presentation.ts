import type { NormalizedAsset, PresentationAsset, PresentationAssetInput } from "./asset.js";
import type {
  NormalizedElement,
  NormalizedPlaceholderDefinition,
  NormalizedTextParagraph,
  PlaceholderDefinitionInput,
  PresentationElement,
  PresentationElementInput,
  TextContentInput,
} from "./element.js";
import type { PresentationSize } from "./geometry.js";
import type { NormalizedPaint, PaintInput, TextStylePresetMap } from "./style.js";
import type { NormalizedPresentationTheme, PresentationThemeInput } from "./theme.js";

export type JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue };

export interface PresentationMetadataInput {
  title?: string;
  author?: string;
  company?: string;
  subject?: string;
  description?: string;
  language?: string;
  keywords?: string[];
  revision?: number;
}

export interface NormalizedPresentationMetadata {
  title: string;
  author: string;
  company?: string;
  subject?: string;
  description?: string;
  language: string;
  keywords: string[];
  revision: number;
}

export interface PresentationSlideInput {
  id?: string;
  layoutId?: string;
  background?: PaintInput;
  elements?: PresentationElementInput[];
  notes?: TextContentInput;
  hidden?: boolean;
  section?: string;
  tags?: string[];
  customData?: Record<string, JsonValue>;
}

export interface PresentationSlide {
  readonly id: string;
  readonly layoutId: string | undefined;
  readonly background: PaintInput | undefined;
  readonly elements: readonly PresentationElement[];
  readonly notes: TextContentInput | undefined;
  readonly hidden: boolean;
  readonly section: string | undefined;
  readonly tags: readonly string[];
  readonly customData: Readonly<Record<string, JsonValue>>;
  addElement(input: PresentationElementInput): PresentationElement;
  insertElement(index: number, input: PresentationElementInput): PresentationElement;
  moveElement(elementId: string, toIndex: number): void;
  removeElement(elementId: string): PresentationElement;
  duplicateElement(elementId: string, toIndex?: number): PresentationElement;
}

export interface SlideLayoutInput {
  id: string;
  name: string;
  background?: PaintInput;
  elements?: PresentationElementInput[];
  placeholders?: PlaceholderDefinitionInput[];
}

export interface SlideLayoutDefinition {
  readonly id: string;
  readonly name: string;
  readonly background?: PaintInput;
  readonly elements: readonly PresentationElement[];
  readonly placeholders: readonly PlaceholderDefinitionInput[];
}

export interface PresentationInit {
  id?: string;
  metadata?: PresentationMetadataInput;
  size?: Partial<PresentationSize>;
  theme?: PresentationThemeInput;
  textStylePresets?: TextStylePresetMap;
}

export interface PresentationDocument {
  readonly id: string;
  readonly metadata: Readonly<PresentationMetadataInput>;
  readonly size: Readonly<PresentationSize>;
  readonly theme: Readonly<PresentationThemeInput>;
  readonly textStylePresets: TextStylePresetMap;
  readonly slides: readonly PresentationSlide[];
  readonly assets: readonly PresentationAsset[];
  readonly layouts: readonly SlideLayoutDefinition[];
  addSlide(input?: PresentationSlideInput): PresentationSlide;
  insertSlide(index: number, input?: PresentationSlideInput): PresentationSlide;
  moveSlide(slideId: string, toIndex: number): void;
  removeSlide(slideId: string): PresentationSlide;
  duplicateSlide(slideId: string, toIndex?: number): PresentationSlide;
  registerAsset(input: PresentationAssetInput): PresentationAsset;
  getAsset(assetId: string): PresentationAsset | undefined;
  defineSlideLayout(input: SlideLayoutInput): SlideLayoutDefinition;
}

export interface NormalizedSlideLayout {
  id: string;
  name: string;
  background: NormalizedPaint;
  elements: NormalizedElement[];
  placeholders: NormalizedPlaceholderDefinition[];
}

export interface NormalizedSlide {
  id: string;
  layoutId: string;
  background: NormalizedPaint;
  backgroundSource: "slide" | "layout" | "theme";
  elements: NormalizedElement[];
  notes: NormalizedTextParagraph[];
  hidden: boolean;
  section?: string;
  tags: string[];
  customData: Record<string, JsonValue>;
}

export interface NormalizedPresentation {
  irVersion: 1;
  id: string;
  metadata: NormalizedPresentationMetadata;
  size: PresentationSize;
  theme: NormalizedPresentationTheme;
  slides: NormalizedSlide[];
  assets: NormalizedAsset[];
  layouts: NormalizedSlideLayout[];
}
