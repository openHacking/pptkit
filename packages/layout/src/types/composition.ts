import type { Box, Insets } from "@pptkit/core";

export type CompositionAlignment = "start" | "center" | "end";
export type CompositionDirection = "horizontal" | "vertical";

interface CompositionBase {
  container: { kind: "slide" | "layout" | "master" | "group"; id: string };
  /** Direct children, in placement order. Drawing order is unchanged. */
  elementIds: readonly string[];
  /** Parent-local area. Defaults to slide size or group coordinateSize. */
  area?: Readonly<Box>;
  padding?: number | Readonly<Partial<Insets>>;
  /** Unrotated authored frame (default), or its transformed axis-aligned bounds. */
  bounds?: "frame" | "visual";
  includeHidden?: boolean;
}

export interface AlignComposition extends CompositionBase {
  type: "align";
  horizontal?: CompositionAlignment;
  vertical?: CompositionAlignment;
}

export interface StackComposition extends CompositionBase {
  type: "stack";
  direction: CompositionDirection;
  gap?: number;
  justify?: CompositionAlignment | "space-between";
  align?: CompositionAlignment;
}

export interface DistributeComposition extends CompositionBase {
  type: "distribute";
  direction: CompositionDirection;
}

export type GridTrack = { type: "fixed"; size: number } | {
  type: "content";
  min?: number;
  max?: number;
} | {
  type: "fraction";
  /** Positive relative weight for fraction tracks; defaults to 1. */
  weight?: number;
  min?: number;
  max?: number;
};

export interface GridPlacement {
  elementId: string;
  /** Zero-based explicit placement. No implicit tracks are created. */
  row: number;
  column: number;
  rowSpan?: number;
  columnSpan?: number;
  horizontal?: CompositionAlignment;
  vertical?: CompositionAlignment;
}

export interface GridComposition extends Omit<CompositionBase, "elementIds"> {
  type: "grid";
  columns: readonly GridTrack[];
  rows: readonly GridTrack[];
  placements: readonly GridPlacement[];
  columnGap?: number;
  rowGap?: number;
}

export type ConstraintProperty = "x" | "y" | "width" | "height";
export type ConstraintMetric = ConstraintProperty | "right" | "bottom" | "centerX" | "centerY";
export interface ConstraintTerm {
  source: "container" | { elementId: string };
  metric: ConstraintMetric;
  factor?: number;
}
/** Affine values in container-local points. References use this rule's resolved frames. */
export type ConstraintExpression = number | { terms: readonly ConstraintTerm[]; constant?: number };
export interface FrameConstraint {
  /** Exact equality. Without value, clamp the incoming frame property to min/max. */
  value?: ConstraintExpression;
  min?: ConstraintExpression;
  max?: ConstraintExpression;
}
export interface ConstrainedElement {
  elementId: string;
  x?: FrameConstraint;
  y?: FrameConstraint;
  width?: FrameConstraint;
  height?: FrameConstraint;
}
export interface ConstraintsComposition extends Omit<CompositionBase, "elementIds" | "bounds"> {
  type: "constraints";
  /** Constraints operate on unrotated frames, never painted/visual bounds. */
  bounds?: "frame";
  elements: readonly ConstrainedElement[];
}

/** Rules apply in order on detached geometry before connectors, images and charts resolve. */
export type LayoutComposition = AlignComposition | StackComposition | DistributeComposition | GridComposition | ConstraintsComposition;
