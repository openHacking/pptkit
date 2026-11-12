import type { TableCellInput, TableElementInput } from "./element.js";

export interface TableEditRegion {
  readonly rowIndex: number;
  readonly columnIndex: number;
  readonly rowCount: number;
  readonly columnCount: number;
}
export type TableEditOperation =
  | { readonly type: "insert-row"; readonly index: number; readonly height: number; readonly cell?: TableCellInput }
  | { readonly type: "insert-column"; readonly index: number; readonly width: number; readonly cell?: TableCellInput }
  | { readonly type: "delete-row"; readonly index: number; readonly content?: "discard" | "reject" | "reject-data" }
  | { readonly type: "delete-column"; readonly index: number; readonly content?: "discard" | "reject" | "reject-data" }
  | { readonly type: "merge"; readonly region: TableEditRegion; readonly content?: "reject" | "append-paragraphs" }
  | { readonly type: "split"; readonly rowIndex: number; readonly columnIndex: number };
export interface TableEditLimits {
  readonly maxRows?: number;
  readonly maxColumns?: number;
  readonly maxGridSlots?: number;
  readonly maxCells?: number;
  readonly maxOperations?: number;
  readonly maxDataValues?: number;
  readonly maxTextCharacters?: number;
  readonly maxWorkUnits?: number;
  readonly maxDepth?: number;
}
export interface TableEditOptions {
  /** Explicit geometry policy: retain the frame or set its dimensions to the grid sums. */
  readonly frame: "preserve" | "fit-grid";
  readonly limits?: TableEditLimits;
}
export interface TableCellAddress {
  readonly rowIndex: number;
  readonly cellIndex: number;
  readonly columnIndex: number;
}
export interface TableCellEditMapping {
  readonly source: TableCellAddress;
  readonly status: "retained" | "merged" | "deleted";
  readonly target?: TableCellAddress;
}
export interface TableEditResult {
  /** Detached, deeply frozen at runtime; pass to replaceElement to commit. */
  readonly table: Readonly<TableElementInput>;
  /** Complete mapping for original explicit anchors; implicit/new blanks have no source. */
  readonly cellMapping: readonly TableCellEditMapping[];
}
export type TableEditErrorCode = "invalid-table-edit-input" | "invalid-table-grid" | "invalid-table-operation" | "table-edit-content-conflict" | "table-edit-header-conflict" | "table-edit-limit-exceeded";
export class TableEditError extends Error {
  constructor(readonly code: TableEditErrorCode, message: string, readonly path: string) {
    super(message); this.name = "TableEditError";
  }
}
