import type { Box, PresentationAssetSource, PresentationSize, ShapeStyle, TextStyle } from "@pptkit/core";

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

export interface ExampleTextElementSpec {
  type: "text";
  text: string;
  box?: Box;
  style?: TextStyle;
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
  shape: "rect" | "ellipse" | "line";
  box?: Box;
  style?: ShapeStyle;
}

export type ExampleElementSpec =
  | string
  | ExampleTextElementSpec
  | ExampleImageElementSpec
  | ExampleShapeElementSpec;

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

export interface VisualPreviewSlide {
  id: string;
  title: string;
  body: string[];
}

export interface ExampleReport {
  example: ExampleDefinition;
  normalizedDocument: NormalizedDocumentReport;
  visualPreview: {
    status: "structural-preview";
    slides: VisualPreviewSlide[];
  };
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
