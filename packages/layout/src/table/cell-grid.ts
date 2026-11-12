import type { Box, NormalizedTableElement } from "@pptkit/core";

export interface TableCellGeometry { rowIndex: number; cellIndex: number; columnIndex: number; box: Box }
export interface InvalidTableCell { rowIndex: number; cellIndex: number }

/** Resolve compact row cells into a grid, accounting for occupied merged slots. */
export function cellGrid(table: NormalizedTableElement): { cells: TableCellGeometry[]; invalid: InvalidTableCell[] } {
  const cells: TableCellGeometry[] = [];
  const invalid: InvalidTableCell[] = [];
  const occupied = table.columns.map(() => 0);
  let y = 0;
  table.rows.forEach((row, rowIndex) => {
    let column = 0, x = 0;
    row.cells.forEach((cell, cellIndex) => {
      while ((occupied[column] ?? 0) > rowIndex) { x += table.columns[column]!; column++; }
      const end = column + cell.colSpan;
      if (!Number.isSafeInteger(cell.colSpan) || cell.colSpan <= 0 || !Number.isSafeInteger(cell.rowSpan) || cell.rowSpan <= 0 || end > table.columns.length || rowIndex + cell.rowSpan > table.rows.length) {
        invalid.push({ rowIndex, cellIndex }); return;
      }
      let width = 0, height = 0, overlap = false;
      for (let i = column; i < end; i++) { if ((occupied[i] ?? 0) > rowIndex) { overlap = true; break; } width += table.columns[i]!; }
      if (overlap) { invalid.push({ rowIndex, cellIndex }); return; }
      for (let i = rowIndex; i < rowIndex + cell.rowSpan; i++) height += table.rows[i]!.height;
      const box = { x, y, width, height };
      if (![box.x, box.y, box.width, box.height].every(Number.isFinite)) { invalid.push({ rowIndex, cellIndex }); return; }
      cells.push({ rowIndex, cellIndex, columnIndex: column, box });
      for (let i = column; i < end; i++) { occupied[i] = rowIndex + cell.rowSpan; x += table.columns[i]!; }
      column = end;
    });
    y += row.height;
  });
  return { cells, invalid };
}
