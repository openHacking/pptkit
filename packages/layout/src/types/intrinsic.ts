export type IntrinsicDimension = {
  type: "content";
  min?: number;
  max?: number;
} | {
  type: "parent";
  /** Fraction of the parent's resolved local coordinate size, default 1. */
  fraction?: number;
  min?: number;
  max?: number;
};

export interface IntrinsicSizingRule {
  elementId: string;
  /** An omitted axis keeps its normalized authored size. */
  width?: IntrinsicDimension;
  height?: IntrinsicDimension;
  includeHidden?: boolean;
}
