export type TextAutoFit =
  | { mode: "none" }
  | { mode: "shrink"; fontScale?: number };

export interface TextStyle {
  fontSize?: number;
  fontFamily?: string;
  fontWeight?: "normal" | "bold";
  color?: string;
  align?: "left" | "center" | "right";
  /** Line-height multiplier, where 1 is 100%. */
  lineSpacing?: number;
  autoFit?: TextAutoFit;
}

export interface ShapeStyle {
  fill?: string;
  stroke?: string;
  strokeWidth?: number;
}
