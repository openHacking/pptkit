export interface AutoTableRow {
  /** Zero-based row index. All other rows retain their authored height. */
  row: number;
  /** Positive minimum, default 1 point. */
  minHeight?: number;
  maxHeight?: number;
}

export interface TableSizingRule {
  elementId: string;
  autoRows: readonly AutoTableRow[];
  /** preserve (default): reject an overflowing sizing result; clip: apply bounded heights with warnings. */
  overflow?: "preserve" | "clip";
  /** Opt in to a bounded wrapped-height search for quarter-turn cell text. */
  rotatedTextSizing?: "conservative" | "compact";
  includeHidden?: boolean;
  /** Measurement tolerance in points, default 0.02 (including browser subpixel rounding). */
  tolerance?: number;
}
