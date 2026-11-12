import type { NormalizedStrokeStyle, Box } from "@pptkit/core";

export interface TableBorderLimits {
  readonly maxGridSlots?: number;
  readonly maxEdges?: number;
  readonly maxWorkUnits?: number;
}
export interface LayoutTableBorderGrid {
  readonly schemaVersion: "pptkit-table-border-grid-v1";
  readonly rowCount: number;
  readonly columnCount: number;
  /** (rowCount + 1) * columnCount; line-major, including invisible merge interiors. */
  readonly horizontal: readonly NormalizedStrokeStyle[];
  /** rowCount * (columnCount + 1); row-major, including invisible merge interiors. */
  readonly vertical: readonly NormalizedStrokeStyle[];
}
export interface TableBorderLayout {
  readonly grid: LayoutTableBorderGrid;
  readonly cells: { rowIndex: number; cellIndex: number; columnIndex: number; box: Box }[];
}
export class TableBorderError extends Error {
  constructor(readonly code: "table-border-limit-exceeded" | "table-border-grid-invalid" | "table-border-style-invalid", message: string, readonly path: string, readonly elementId?: string) {
    super(message); this.name = "TableBorderError";
  }
}
