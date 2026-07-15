import type { Box, PresentationAssetSource, PresentationSize, TextFrameStyleInput, TextStylePresetMap } from "@pptkit/core";

export type FeatureId = string;

export type InputKind = "html" | "markdown" | "presentation-config";

export type CapabilityStatus = "implemented" | "placeholder" | "not-implemented";

export interface CapabilityMatrix {
  normalize: CapabilityStatus;
  render: CapabilityStatus;
  exportPptx: CapabilityStatus;
}

export interface ExampleSource {
  label: string;
  content: string;
}

export type ExampleTextBoxSpec = Omit<Box, "height"> & { height?: number };

export interface ExampleTextElementSpec {
  type: "text";
  text: string;
  textStylePreset?: string;
  box?: ExampleTextBoxSpec;
  style?: {
    fontSize?: number;
    fontFamily?: string;
    fontWeight?: "normal" | "bold";
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    color?: string;
    language?: string;
    align?: "left" | "center" | "right" | "justify";
    indent?: number;
    hanging?: number;
    lineSpacing?: number;
    spaceBefore?: number;
    spaceAfter?: number;
    bullet?:
      | { type: "none" }
      | { type: "bullet"; character?: string }
      | {
          type: "number";
          style?: "arabicPeriod" | "arabicParen" | "alphaLowerPeriod" | "alphaUpperPeriod";
          startAt?: number;
        };
    autoFit?: { mode: "none" } | { mode: "shrink"; fontScale?: number };
  };
}

export interface ExampleImageAssetSpec {
  id?: string;
  source: PresentationAssetSource;
  mimeType?: string;
  width?: number;
  height?: number;
  altText?: string;
  dedupeKey?: string;
}

export interface ExampleImageElementSpec {
  type: "image";
  asset: ExampleImageAssetSpec;
  box?: Box;
  altText?: string;
}

export interface ExampleShapeElementSpec {
  type: "shape";
  shape: "rect" | "roundRect" | "ellipse" | "line";
  box?: Box;
  text?: {
    content: string;
    textStylePreset?: string;
    frame?: TextFrameStyleInput;
  };
  style?: {
    fill?: string;
    stroke?: string;
    strokeWidth?: number;
  };
}

export interface ExampleGroupElementSpec {
  type: "group";
  coordinateSize: { width: number; height: number };
  children: ExampleElementSpec[];
  box?: Box;
}

export interface ExampleTableElementSpec {
  type: "table";
  columns: number[];
  rows: Array<{
    height?: number;
    cells: Array<{ content: string; rowSpan?: number; colSpan?: number }>;
  }>;
  box?: Box;
}

export type ExampleElementSpec =
  | string
  | ExampleTextElementSpec
  | ExampleImageElementSpec
  | ExampleShapeElementSpec
  | ExampleGroupElementSpec
  | ExampleTableElementSpec;

export interface ExampleSlideSpec {
  id?: string;
  title: string;
  background?: string;
  elements: ExampleElementSpec[];
}

export interface ExampleInputData {
  title: string;
  summary: string;
  size?: Partial<PresentationSize>;
  textStylePresets?: TextStylePresetMap;
  slides: ExampleSlideSpec[];
}

export interface ExampleDefinition {
  id: string;
  title: string;
  feature: FeatureId;
  description: string;
  inputKind: InputKind;
  source: ExampleSource;
  scenarioTags: string[];
  expectedCapabilities: CapabilityMatrix;
  status: "ready" | "placeholder";
  createInput(): ExampleInputData;
}

export interface ExampleSummary {
  id: string;
  title: string;
  feature: FeatureId;
  description: string;
  inputKind: InputKind;
  scenarioTags: string[];
  status: "ready" | "placeholder";
}

export interface WorkbenchPayload {
  features: FeatureId[];
  examples: ExampleSummary[];
}

export interface NormalizedDocumentReport {
  title: string;
  summary: string;
  slideCount: number;
  slides: Array<{
    id: string;
    title: string;
    elements: string[];
  }>;
}

export interface ExampleReport {
  example: ExampleDefinition;
  presentationInput: ExampleInputData;
  normalizedDocument: NormalizedDocumentReport;
  renderResult: {
    slideCount: number;
    status: string;
  };
  exportResult: {
    output: string;
    slideCount: number;
    status: string;
  };
  diagnostics: string[];
}

export interface ExampleSourceRequest {
  source: string;
}
