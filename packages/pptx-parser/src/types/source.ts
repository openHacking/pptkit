/** A format-layer location, never a Core/Canonical IR identity. */
export interface PptxSourceLocation {
  readonly part: string;
  readonly path?: string;
  readonly sourceObjectId?: string;
  readonly relationshipId?: string;
  /** One-based raw element preorder when provided, including inactive markup. */
  readonly elementIndex?: number;
}
