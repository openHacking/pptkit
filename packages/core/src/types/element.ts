import type { Box } from "./geometry.js";
import type { ShapeStyle, TextStyle } from "./style.js";

export interface TextElementInput {
  type: "text";
  text: string;
  box: Box;
  style?: TextStyle;
}

export interface ImageElementInput {
  type: "image";
  assetId: string;
  box: Box;
  altText?: string;
}

export interface ShapeElementInput {
  type: "shape";
  shape: "rect" | "ellipse" | "line";
  box: Box;
  style?: ShapeStyle;
}

export type PresentationElementInput = TextElementInput | ImageElementInput | ShapeElementInput;

export interface NormalizedTextElement {
  type: "text";
  text: string;
  box: Box;
  style: TextStyle;
}

export interface NormalizedImageElement extends ImageElementInput {}

export interface NormalizedShapeElement {
  type: "shape";
  shape: "rect" | "ellipse" | "line";
  box: Box;
  style: ShapeStyle;
}

export type NormalizedElement =
  | NormalizedTextElement
  | NormalizedImageElement
  | NormalizedShapeElement;
