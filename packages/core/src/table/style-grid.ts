import type { TableCellInput, TableElementInput } from "../types/element.js";
export const TABLE_STYLE_MAX_SLOTS = 100000;
export class TableStyleGridError extends Error {
  constructor(readonly path: string, message: string) { super(message); }
}
export interface TableStyleAnchor { row: number; column: number; rowSpan: number; colSpan: number; cell: TableCellInput; inputIndex?: number }
/** Bounded compact-grid proof shared by validation and normalization. */
export function tableStyleAnchors(table: Pick<TableElementInput, "rows" | "columns">): TableStyleAnchor[][] {
  const rows = table.rows.length, columns = table.columns.length;
  if (!rows || !columns || rows * columns > TABLE_STYLE_MAX_SLOTS) throw new TableStyleGridError("style", `Styled tables require 1..${TABLE_STYLE_MAX_SLOTS} physical grid slots.`);
  const grid = new Uint8Array(rows * columns), result: TableStyleAnchor[][] = [];
  for (let r = 0; r < rows; r++) {
    const anchors: TableStyleAnchor[] = []; let c = 0;
    for (let i = 0; i < table.rows[r]!.cells.length; i++) {
      const cell = table.rows[r]!.cells[i]!, rs = cell.rowSpan ?? 1, cs = cell.colSpan ?? 1;
      while (c < columns && grid[r * columns + c]) c++;
      const path = `rows.${r}.cells.${i}`;
      if (!Number.isSafeInteger(rs) || !Number.isSafeInteger(cs) || rs < 1 || cs < 1 || r + rs > rows || c + cs > columns) throw new TableStyleGridError(path, "Styled cell span exceeds the physical grid.");
      for (let y = r; y < r + rs; y++) for (let x = c; x < c + cs; x++) {
        if (grid[y * columns + x]) throw new TableStyleGridError(path, "Styled cell rectangles overlap.");
        grid[y * columns + x] = 1;
      }
      anchors.push({ row: r, column: c, rowSpan: rs, colSpan: cs, cell, inputIndex: i }); c += cs;
    }
    result.push(anchors);
  }
  // Materialize implicit blanks so bands/outer edges apply to the whole table.
  for (let r = 0; r < rows; r++) {
    const anchors = result[r]!, byColumn = new Map(anchors.map(a => [a.column, a]));
    const ordered: TableStyleAnchor[] = [];
    for (let c = 0; c < columns; c++) {
      const anchor = byColumn.get(c);
      if (anchor) ordered.push(anchor);
      else if (!grid[r * columns + c]) ordered.push({ row: r, column: c, rowSpan: 1, colSpan: 1, cell: { content: "" } });
    }
    result[r] = ordered;
  }
  return result;
}
