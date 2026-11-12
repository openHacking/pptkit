import type { TableCellInput, TableElementInput, TextContentInput, TextParagraphInput } from "../types/element.js";
import type { TableCellAddress } from "../types/table-edit.js";
import { EditBudget, fail, fields, integer, object, positive } from "./edit-input.js";

export interface Anchor {
  row: number;
  column: number;
  rowSpan: number;
  colSpan: number;
  cell: TableCellInput;
  origins: { index: number; merged: boolean }[];
}
export interface GridState {
  table: TableElementInput;
  columns: number[];
  heights: number[];
  anchors: Anchor[];
  sources: TableCellAddress[];
}
export function textContent(value: unknown, path: string, budget: EditBudget): asserts value is TextContentInput {
  if (typeof value === "string") return;
  if (!Array.isArray(value)) fail("invalid-table-edit-input", path, "Expected text or paragraph content.");
  for (let p = 0; p < value.length; p++) {
    const paragraph = fields(value[p], ["runs", "style"], `${path}/${p}`);
    if (!Array.isArray(paragraph.runs)) fail("invalid-table-edit-input", `${path}/${p}/runs`, "Expected paragraph runs.");
    for (let r = 0; r < paragraph.runs.length; r++) {
      budget.step(); const run = fields(paragraph.runs[r], ["text", "style", "action"], `${path}/${p}/runs/${r}`);
      if (typeof run.text !== "string") fail("invalid-table-edit-input", `${path}/${p}/runs/${r}/text`, "Expected scalar text.");
    }
  }
}
export function cellInput(value: unknown, path: string, budget: EditBudget): TableCellInput {
  const cell = fields(value, ["content", "textStylePreset", "rowSpan", "colSpan", "style", "headerScope"], path);
  textContent(cell.content, `${path}/content`, budget);
  if (cell.headerScope !== undefined && cell.headerScope !== "row" && cell.headerScope !== "column") fail("invalid-table-edit-input", `${path}/headerScope`, "Expected row or column header scope.");
  if (cell.textStylePreset !== undefined && typeof cell.textStylePreset !== "string") fail("invalid-table-edit-input", `${path}/textStylePreset`, "Expected a text preset name.");
  if (cell.rowSpan !== undefined) integer(cell.rowSpan, `${path}/rowSpan`, 1);
  if (cell.colSpan !== undefined) integer(cell.colSpan, `${path}/colSpan`, 1);
  return value as TableCellInput;
}
export function anchorGrid(anchors: readonly Anchor[], rows: number, columns: number, budget: EditBudget): (Anchor | undefined)[] {
  const size = budget.size(rows, columns, anchors.length); budget.step(size);
  const grid: (Anchor | undefined)[] = new Array(size).fill(undefined);
  for (const anchor of anchors) {
    budget.step();
    if (anchor.row < 0 || anchor.column < 0 || anchor.row + anchor.rowSpan > rows || anchor.column + anchor.colSpan > columns) fail("invalid-table-grid", "/table/rows", "Cell span extends outside the table.");
    for (let row = anchor.row; row < anchor.row + anchor.rowSpan; row++) for (let column = anchor.column; column < anchor.column + anchor.colSpan; column++) {
      budget.step(); const index = row * columns + column;
      if (grid[index]) fail("invalid-table-grid", "/table/rows", "Merged table cells overlap.");
      grid[index] = anchor;
    }
  }
  return grid;
}
export function blankCell(template: TableCellInput, budget: EditBudget): TableCellInput {
  return budget.capture(template, "/materialized-cell");
}
export function fillGrid(state: GridState, template: TableCellInput, budget: EditBudget): void {
  const rows = state.heights.length, columns = state.columns.length, grid = anchorGrid(state.anchors, rows, columns, budget);
  for (let i = 0; i < grid.length; i++) {
    budget.step(); if (grid[i]) continue;
    budget.size(rows, columns, state.anchors.length + 1);
    state.anchors.push({ row: Math.floor(i / columns), column: i % columns, rowSpan: 1, colSpan: 1, cell: blankCell(template, budget), origins: [] });
  }
}
export function readGrid(input: TableElementInput, budget: EditBudget): GridState {
  const table = budget.capture(input, "/table"); object(table, "/table");
  if (table.type !== "table" || !Array.isArray(table.columns) || !Array.isArray(table.rows) || !table.columns.length || !table.rows.length) fail("invalid-table-grid", "/table", "Expected a nonempty table grid.");
  if (table.nativeFirstRowStyle !== undefined && typeof table.nativeFirstRowStyle !== "boolean") fail("invalid-table-edit-input", "/table/nativeFirstRowStyle", "Expected a boolean.");
  if (table.nativeTableStylePreset !== undefined && table.nativeTableStylePreset !== "medium-2-accent-1" && table.nativeTableStylePreset !== "medium-2-accent-2") fail("invalid-table-edit-input", "/table/nativeTableStylePreset", "Unsupported native table style preset.");
  if (table.nativeTableStyleFlags !== undefined) {
    const flags = fields(table.nativeTableStyleFlags, ["firstColumn", "lastColumn", "lastRow", "bandedRows", "bandedColumns"], "/table/nativeTableStyleFlags");
    for (const [key, value] of Object.entries(flags)) if (typeof value !== "boolean") fail("invalid-table-edit-input", `/table/nativeTableStyleFlags/${key}`, "Expected a boolean.");
  }
  const count = budget.size(table.rows.length, table.columns.length); budget.step(count);
  const columns = table.columns.map((width, index) => positive(width, `/table/columns/${index}`));
  if (table.box !== undefined) {
    const box = fields(table.box, ["x", "y", "width", "height"], "/table/box");
    if (![box.x, box.y, box.width, box.height].every(v => typeof v === "number" && Number.isFinite(v)) || (box.width as number) < 0 || (box.height as number) < 0) fail("invalid-table-edit-input", "/table/box", "Expected a finite frame with nonnegative dimensions.");
  }
  const state: GridState = { table, columns, heights: [], anchors: [], sources: [] }, occupied: (Anchor | undefined)[] = new Array(count).fill(undefined);
  for (let row = 0; row < table.rows.length; row++) {
    budget.step(); const sourceRow = fields(table.rows[row], ["cells", "height"], `/table/rows/${row}`);
    if (!Array.isArray(sourceRow.cells)) fail("invalid-table-grid", `/table/rows/${row}/cells`, "Expected compact anchor cells.");
    const height = sourceRow.height ?? (table.box ? table.box.height / table.rows.length : undefined);
    state.heights.push(positive(height, `/table/rows/${row}/height`));
    let column = 0;
    for (let cellIndex = 0; cellIndex < sourceRow.cells.length; cellIndex++) {
      const path = `/table/rows/${row}/cells/${cellIndex}`, cell = cellInput(sourceRow.cells[cellIndex], path, budget);
      while (column < columns.length && occupied[row * columns.length + column]) { budget.step(); column++; }
      const rowSpan = cell.rowSpan ?? 1, colSpan = cell.colSpan ?? 1;
      if (column + colSpan > columns.length || row + rowSpan > table.rows.length) fail("invalid-table-grid", path, "Cell span extends outside the table.");
      budget.size(table.rows.length, columns.length, state.anchors.length + 1);
      const source = state.sources.length, anchor: Anchor = { row, column, rowSpan, colSpan, cell, origins: [{ index: source, merged: false }] };
      state.sources.push({ rowIndex: row, cellIndex, columnIndex: column }); state.anchors.push(anchor);
      for (let y = row; y < row + rowSpan; y++) for (let x = column; x < column + colSpan; x++) {
        budget.step(); const at = y * columns.length + x;
        if (occupied[at]) fail("invalid-table-grid", path, "A compact cell crosses an occupied merged slot.");
        occupied[at] = anchor;
      }
      column += colSpan;
    }
  }
  // Existing exporters already represent omitted grid slots as blank physical cells.
  fillGrid(state, { content: "" }, budget);
  return state;
}
export function hasContent(cell: TableCellInput, budget: EditBudget): boolean {
  if (typeof cell.content === "string") return cell.content.length > 0;
  for (const paragraph of cell.content) for (const run of paragraph.runs) { budget.step(); if (run.text.length || run.action !== undefined) return true; }
  return false;
}
export function appendParagraphs(anchors: readonly Anchor[], budget: EditBudget, path: string): TextParagraphInput[] {
  const preset = anchors[0]!.cell.textStylePreset;
  const output: TextParagraphInput[] = [];
  for (const anchor of anchors) {
    budget.step();
    if (anchor.cell.textStylePreset !== preset) fail("table-edit-content-conflict", path, "Append requires identical cell text presets; materialize effective text styles before combining different presets.");
    if (typeof anchor.cell.content !== "string") {
      for (const paragraph of anchor.cell.content) output.push(budget.capture(paragraph, path));
      continue;
    }
    const text = anchor.cell.content, newline = /\r\n|\r|\n/g;
    let start = 0, match: RegExpExecArray | null;
    while ((match = newline.exec(text))) {
      output.push(budget.capture({ runs: [{ text: text.slice(start, match.index) }] }, path));
      start = match.index + match[0].length;
    }
    output.push(budget.capture({ runs: [{ text: text.slice(start) }] }, path));
  }
  return output;
}
