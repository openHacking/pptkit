export type LengthUnit = "pt";

export interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PresentationSize {
  width: number;
  height: number;
  unit: LengthUnit;
}
