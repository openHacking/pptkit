import type { TableCellInput, TableElementInput } from "../types/element.js";
import type { TableCellEditMapping, TableEditOperation, TableEditOptions, TableEditResult } from "../types/table-edit.js";
import { EditBudget, fail, fields, freeze, integer, positive } from "./edit-input.js";
import { anchorGrid, appendParagraphs, blankCell, cellInput, fillGrid, hasContent, readGrid, type Anchor, type GridState } from "./edit-grid.js";

function at(value: unknown, length: number, path: string, insertion = false): number {
  const index = integer(value, path);
  if (index > length || (!insertion && index === length)) fail("invalid-table-operation", path, "Table index is out of range.");
  return index;
}
function sortAnchors(anchors: Anchor[], budget: EditBudget): void {
  // Charge a fixed n-log-n cost before native sort, keeping the resource
  // contract independent of the JavaScript engine's sorting implementation.
  budget.step(anchors.length * Math.ceil(Math.log2(anchors.length + 1)));
  anchors.sort((a, b) => a.row - b.row || a.column - b.column);
}
function insert(state: GridState, operation: Record<string, unknown>, row: boolean, path: string, budget: EditBudget): void {
  const dimensions = row ? state.heights : state.columns, index = at(operation.index, dimensions.length, `${path}/index`, true);
  const value = positive(row ? operation.height : operation.width, `${path}/${row ? "height" : "width"}`);
  budget.size(state.heights.length + (row ? 1 : 0), state.columns.length + (row ? 0 : 1));
  const template = operation.cell === undefined ? { content: "" } : cellInput(operation.cell, `${path}/cell`, budget);
  if ((template.rowSpan ?? 1) !== 1 || (template.colSpan ?? 1) !== 1) fail("invalid-table-operation", `${path}/cell`, "Insertion templates must be unit cells.");
  for (const anchor of state.anchors) {
    budget.step(); const start = row ? anchor.row : anchor.column, span = row ? anchor.rowSpan : anchor.colSpan;
    if (start < index && index < start + span) { if (row) anchor.rowSpan++; else anchor.colSpan++; }
    else if (start >= index) { if (row) anchor.row++; else anchor.column++; }
  }
  dimensions.splice(index, 0, value); fillGrid(state, template, budget);
}
function remove(state: GridState, operation: Record<string, unknown>, row: boolean, path: string, budget: EditBudget): void {
  const dimensions = row ? state.heights : state.columns, index = at(operation.index, dimensions.length, `${path}/index`);
  if (dimensions.length === 1) fail("invalid-table-operation", path, "Delete the table element instead of its final row/column.");
  if (operation.content !== undefined && operation.content !== "discard" && operation.content !== "reject" && operation.content !== "reject-data") fail("invalid-table-operation", `${path}/content`, "Deletion content policy must be discard, reject or reject-data.");
  const retained: Anchor[] = [];
  for (const anchor of state.anchors) {
    budget.step(); const start = row ? anchor.row : anchor.column, span = row ? anchor.rowSpan : anchor.colSpan;
    if (index >= start && index < start + span) {
      if (span === 1) {
        if (operation.content === "reject" && hasContent(anchor.cell, budget)) fail("table-edit-content-conflict", path, "Deletion would discard nonempty or actionable cell content.");
        if (operation.content === "reject-data" && (anchor.cell.style !== undefined || anchor.cell.textStylePreset !== undefined || anchor.cell.headerScope !== undefined || typeof anchor.cell.content !== "string" || anchor.cell.content.length > 0)) fail("table-edit-content-conflict", path, "Deletion would discard cell content, structured text, header scope or explicit cell styling.");
        continue;
      }
      if (row) anchor.rowSpan--; else anchor.colSpan--;
    } else if (start > index) { if (row) anchor.row--; else anchor.column--; }
    retained.push(anchor);
  }
  state.anchors = retained; dimensions.splice(index, 1);
}
function merge(state: GridState, operation: Record<string, unknown>, path: string, budget: EditBudget): void {
  const region = fields(operation.region, ["rowIndex", "columnIndex", "rowCount", "columnCount"], `${path}/region`);
  const row = at(region.rowIndex, state.heights.length, `${path}/region/rowIndex`), column = at(region.columnIndex, state.columns.length, `${path}/region/columnIndex`);
  const rows = integer(region.rowCount, `${path}/region/rowCount`, 1), columns = integer(region.columnCount, `${path}/region/columnCount`, 1);
  if (rows > state.heights.length - row || columns > state.columns.length - column) fail("invalid-table-operation", `${path}/region`, "Merge region extends outside the grid.");
  if (operation.content !== undefined && operation.content !== "reject" && operation.content !== "append-paragraphs") fail("invalid-table-operation", `${path}/content`, "Unknown merge content policy.");
  const selected: Anchor[] = [];
  for (const anchor of state.anchors) {
    budget.step();
    if (anchor.row >= row + rows || anchor.column >= column + columns || anchor.row + anchor.rowSpan <= row || anchor.column + anchor.colSpan <= column) continue;
    if (anchor.row < row || anchor.column < column || anchor.row + anchor.rowSpan > row + rows || anchor.column + anchor.colSpan > column + columns) fail("invalid-table-operation", `${path}/region`, "Merge selection cuts through an existing merged cell.");
    selected.push(anchor);
  }
  sortAnchors(selected, budget);
  const anchor = selected[0]!;
  if (selected.some(cell => cell.cell.headerScope !== anchor.cell.headerScope)) {
    fail("table-edit-header-conflict", path, "Merged cells must have the same explicit header scope.");
  }
  if (operation.content !== "append-paragraphs" && selected.slice(1).some(cell => hasContent(cell.cell, budget))) fail("table-edit-content-conflict", path, "Non-anchor content requires explicit append-paragraphs merging.");
  if (operation.content === "append-paragraphs" && selected.length > 1) anchor.cell = { ...anchor.cell, content: appendParagraphs(selected, budget, `${path}/content`) };
  for (const cell of selected.slice(1)) for (const origin of cell.origins) { budget.step(); anchor.origins.push({ index: origin.index, merged: true }); }
  anchor.rowSpan = rows; anchor.colSpan = columns;
  const removed = new Set(selected.slice(1)); state.anchors = state.anchors.filter(cell => !removed.has(cell));
}
function split(state: GridState, operation: Record<string, unknown>, path: string, budget: EditBudget): void {
  const row = at(operation.rowIndex, state.heights.length, `${path}/rowIndex`), column = at(operation.columnIndex, state.columns.length, `${path}/columnIndex`);
  const anchor = anchorGrid(state.anchors, state.heights.length, state.columns.length, budget)[row * state.columns.length + column]!;
  if (anchor.rowSpan === 1 && anchor.colSpan === 1) return;
  const bottom = anchor.row + anchor.rowSpan, right = anchor.column + anchor.colSpan;
  budget.size(state.heights.length, state.columns.length, state.anchors.length + anchor.rowSpan * anchor.colSpan - 1);
  const template: TableCellInput = { content: "", ...(anchor.cell.headerScope === undefined ? {} : { headerScope: anchor.cell.headerScope }), ...(anchor.cell.style !== undefined ? { style: anchor.cell.style } : {}), ...(anchor.cell.textStylePreset !== undefined ? { textStylePreset: anchor.cell.textStylePreset } : {}) };
  for (let y = anchor.row; y < bottom; y++) for (let x = anchor.column; x < right; x++) {
    budget.step(); if (y === anchor.row && x === anchor.column) continue;
    state.anchors.push({ row: y, column: x, rowSpan: 1, colSpan: 1, cell: blankCell(template, budget), origins: [] });
  }
  anchor.rowSpan = 1; anchor.colSpan = 1;
}

/** Pure, atomic fixed-grid structural transaction. The caller explicitly commits the returned snapshot. */
export function editTable(input: TableElementInput, operations: readonly TableEditOperation[], options: TableEditOptions): TableEditResult {
  const captured = fields(options, ["frame", "limits"], "/options"), budget = new EditBudget(captured.limits);
  if (captured.frame !== "preserve" && captured.frame !== "fit-grid") fail("invalid-table-edit-input", "/options/frame", "Choose preserve or fit-grid frame geometry explicitly.");
  if (Array.isArray(operations) && operations.length > budget.limits.maxOperations) fail("table-edit-limit-exceeded", "/operations", "Operation count exceeds maxOperations.");
  const edits = budget.capture(operations, "/operations");
  if (!Array.isArray(edits)) fail("invalid-table-edit-input", "/operations", "Expected an operation array.");
  if (edits.length > budget.limits.maxOperations) fail("table-edit-limit-exceeded", "/operations", "Operation count exceeds maxOperations.");
  const state = readGrid(input, budget);
  for (let i = 0; i < edits.length; i++) {
    const path = `/operations/${i}`, operation = fields(edits[i], ["type", "index", "height", "width", "cell", "region", "content", "rowIndex", "columnIndex"], path);
    const allowed: Record<string, readonly string[]> = {
      "insert-row": ["type", "index", "height", "cell"], "insert-column": ["type", "index", "width", "cell"],
      "delete-row": ["type", "index", "content"], "delete-column": ["type", "index", "content"], "merge": ["type", "region", "content"], "split": ["type", "rowIndex", "columnIndex"],
    };
    if (typeof operation.type !== "string" || !Object.hasOwn(allowed, operation.type)) fail("invalid-table-operation", `${path}/type`, "Unknown table operation.");
    fields(operation, allowed[operation.type]!, path);
    if (operation.type === "insert-row" || operation.type === "insert-column") insert(state, operation, operation.type === "insert-row", path, budget);
    else if (operation.type === "delete-row" || operation.type === "delete-column") remove(state, operation, operation.type === "delete-row", path, budget);
    else if (operation.type === "merge") merge(state, operation, path, budget);
    else split(state, operation, path, budget);
    anchorGrid(state.anchors, state.heights.length, state.columns.length, budget);
  }
  const rows = state.heights.map(height => ({ height, cells: [] as TableCellInput[] }));
  const mapping: TableCellEditMapping[] = state.sources.map(source => ({ source, status: "deleted" }));
  sortAnchors(state.anchors, budget);
  for (const anchor of state.anchors) {
    budget.step(); const row = rows[anchor.row]!, target = { rowIndex: anchor.row, cellIndex: row.cells.length, columnIndex: anchor.column };
    row.cells.push({ ...anchor.cell, rowSpan: anchor.rowSpan, colSpan: anchor.colSpan });
    for (const origin of anchor.origins) { budget.step(); mapping[origin.index] = { source: state.sources[origin.index]!, status: origin.merged ? "merged" : "retained", target }; }
  }
  const table: TableElementInput = { ...state.table, columns: state.columns, rows };
  const width = state.columns.reduce((a, b) => a + b, 0), height = state.heights.reduce((a, b) => a + b, 0);
  if (!Number.isFinite(width) || !Number.isFinite(height)) fail("invalid-table-operation", "/table/box", "Grid dimensions overflow finite geometry.");
  if (captured.frame === "fit-grid") {
    table.box = { x: table.box?.x ?? 0, y: table.box?.y ?? 0, width, height };
  }
  // Bound final materialized data too, including repetitions of template/style references.
  return freeze(budget.capture({ table, cellMapping: mapping }, "/result"));
}
