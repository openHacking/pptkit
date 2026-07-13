export type LengthUnit = "pt";

export interface Point {
  x: number;
  y: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface Box extends Point, Size {}

export interface Insets {
  top: number;
  right: number;
  bottom: number;
  left: number;
}

export interface PresentationSize extends Size {
  unit: LengthUnit;
}
