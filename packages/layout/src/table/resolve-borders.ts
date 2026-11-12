import type { NormalizedStrokeStyle, NormalizedTableElement, TableCellEdge } from "@pptkit/core";
import { TableBorderError, type TableBorderLayout, type TableBorderLimits } from "../types/table-borders.js";

const sides = ["left", "right", "top", "bottom"] as const;
const themeRoles = new Set(["background1", "background2", "text1", "text2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink", "followedHyperlink"]);
const dashRank = { solid: 4, dash: 3, dashDot: 2, dot: 1 } as const;
const invisible: NormalizedStrokeStyle = { paint: { type: "none" }, width: 0, dash: "solid", beginArrow: "none", endArrow: "none" };

/** Internal shared accounting for all opt-in tables in one Layout request. */
export class TableBorderBudget {
  private slots = 0;
  private edges = 0;
  private work = 0;
  readonly limits: Required<TableBorderLimits>;
  constructor(input?: TableBorderLimits) {
    const defaults = { maxGridSlots: 100000, maxEdges: 250000, maxWorkUnits: 2000000 }, caps = { maxGridSlots: 2000000, maxEdges: 6000002, maxWorkUnits: 100000000 };
    if (input !== undefined && (!input || typeof input !== "object" || Array.isArray(input) || ![Object.prototype, null].includes(Object.getPrototypeOf(input)))) throw new TypeError("Expected plain table border limits.");
    for (const key of Reflect.ownKeys(input ?? {})) {
      if (typeof key !== "string" || !Object.hasOwn(defaults, key)) throw new TypeError("Unknown table border limit.");
      const descriptor = Object.getOwnPropertyDescriptor(input ?? {}, key)!;
      if (!("value" in descriptor) || !descriptor.enumerable) throw new TypeError("Table border limits must be data properties.");
      const name = key as keyof typeof defaults, value = descriptor.value === undefined ? defaults[name] : descriptor.value;
      if (!Number.isSafeInteger(value) || value <= 0 || value > caps[name]) throw new RangeError(`Invalid table border limit ${key}.`);
      defaults[name] = value;
    }
    this.limits = defaults;
  }
  step(count: number, table: NormalizedTableElement, path = "table"): void {
    this.work += count;
    if (!Number.isSafeInteger(this.work) || this.work > this.limits.maxWorkUnits) throw new TableBorderError("table-border-limit-exceeded", "Table border traversal exceeds maxWorkUnits.", path, table.id);
  }
  grid(rows: number, columns: number, table: NormalizedTableElement): void {
    const slots = rows * columns, edges = (rows + 1) * columns + rows * (columns + 1);
    this.slots += slots; this.edges += edges;
    if (![slots, edges, this.slots, this.edges].every(Number.isSafeInteger) || this.slots > this.limits.maxGridSlots || this.edges > this.limits.maxEdges) throw new TableBorderError("table-border-limit-exceeded", "Table border grids exceed shared slot/edge limits.", "table", table.id);
    this.step(slots + edges + rows + columns, table);
  }
}
function invalid(table: NormalizedTableElement, path: string, message: string): never { throw new TableBorderError("table-border-grid-invalid", message, path, table?.id); }
function stroke(value: NormalizedStrokeStyle, table: NormalizedTableElement, path: string): void {
  const fail = (): never => { throw new TableBorderError("table-border-style-invalid", "Expected a finite normalized border stroke without arrowheads.", path, table.id); };
  if (!value || typeof value !== "object" || Object.keys(value).some(key => !["paint", "width", "dash", "beginArrow", "endArrow"].includes(key)) || !Number.isFinite(value.width) || value.width < 0 || !Object.hasOwn(dashRank, value.dash) || value.beginArrow !== "none" || value.endArrow !== "none") fail();
  const paint = value.paint;
  if (!paint || typeof paint !== "object" || (paint.type !== "none" && paint.type !== "solid")) fail();
  if (paint.type === "solid") {
    const color = paint.color;
    if (!Number.isFinite(paint.opacity) || paint.opacity < 0 || paint.opacity > 1 || !(typeof color === "string" ? /^#?[a-fA-F\d]{6}$/.test(color) : color && typeof color === "object" && themeRoles.has(color.theme))) fail();
  }
}
interface Owner { row: number; column: number; rowSpan: number; colSpan: number; cellIndex: number; style: NormalizedTableElement["rows"][number]["cells"][number]["style"] }
interface Candidate { stroke: NormalizedStrokeStyle; explicit: boolean }
function visible(style: NormalizedStrokeStyle): boolean { return style.width > 0 && style.paint.type === "solid" && style.paint.opacity > 0; }
function winner(first: Candidate | undefined, second: Candidate | undefined): NormalizedStrokeStyle {
  if (!first) return structuredClone(second?.stroke ?? invisible);
  if (!second) return structuredClone(first.stroke);
  const a = first.stroke, b = second.stroke;
  const choice = visible(a) !== visible(b) ? visible(a) : first.explicit !== second.explicit ? first.explicit : a.width !== b.width ? a.width > b.width : dashRank[a.dash] !== dashRank[b.dash] ? dashRank[a.dash] > dashRank[b.dash] : true;
  return structuredClone(choice ? a : b);
}

export function resolveBorders(table: NormalizedTableElement, budget: TableBorderBudget): TableBorderLayout {
  if (!table || !Array.isArray(table.rows) || !Array.isArray(table.columns) || !table.rows.length || !table.columns.length) invalid(table, "table", "A border grid requires nonempty rows and columns.");
  const rows = table.rows.length, columns = table.columns.length;
  budget.grid(rows, columns, table);
  const xs = [0], ys = [0];
  for (const width of table.columns) {
    if (!Number.isFinite(width) || width <= 0) invalid(table, "table.columns", "Border columns must be positive and finite.");
    xs.push(xs[xs.length - 1]! + width);
  }
  for (const row of table.rows) {
    if (!row || !Number.isFinite(row.height) || row.height <= 0 || !Array.isArray(row.cells)) invalid(table, "table.rows", "Border rows require positive finite heights and compact cell arrays.");
    ys.push(ys[ys.length - 1]! + row.height);
  }
  if (!Number.isFinite(xs[xs.length - 1]) || !Number.isFinite(ys[ys.length - 1])) invalid(table, "table", "Border geometry sums overflow finite coordinates.");
  const grid = new Int32Array(rows * columns).fill(-1), owners: Owner[] = [], cells: TableBorderLayout["cells"] = [];
  for (let row = 0; row < rows; row++) {
    let column = 0;
    for (let cellIndex = 0; cellIndex < table.rows[row]!.cells.length; cellIndex++) {
      budget.step(1, table); const cell = table.rows[row]!.cells[cellIndex]!, path = `table.rows.${row}.cells.${cellIndex}`;
      if (!cell || !cell.style || !Number.isSafeInteger(cell.rowSpan) || cell.rowSpan < 1 || !Number.isSafeInteger(cell.colSpan) || cell.colSpan < 1) invalid(table, path, "A border anchor requires positive integral spans and normalized style.");
      while (column < columns && grid[row * columns + column] !== -1) { budget.step(1, table); column++; }
      if (column + cell.colSpan > columns || row + cell.rowSpan > rows) invalid(table, path, "Border anchor extends outside the grid.");
      stroke(cell.style.stroke, table, `${path}.style.stroke`);
      if (cell.style.borders !== undefined) {
        if (!cell.style.borders || typeof cell.style.borders !== "object" || Array.isArray(cell.style.borders) || Object.keys(cell.style.borders).some(side => !sides.includes(side as TableCellEdge))) throw new TableBorderError("table-border-style-invalid", "Unknown border edge.", `${path}.style.borders`, table.id);
        for (const side of sides) if (cell.style.borders[side] !== undefined) stroke(cell.style.borders[side]!, table, `${path}.style.borders.${side}`);
      }
      const index = owners.length;
      owners.push({ row, column, rowSpan: cell.rowSpan, colSpan: cell.colSpan, cellIndex, style: cell.style });
      for (let y = row; y < row + cell.rowSpan; y++) for (let x = column; x < column + cell.colSpan; x++) {
        budget.step(1, table); const at = y * columns + x;
        if (grid[at] !== -1) invalid(table, path, "Border anchor overlaps another merge.");
        grid[at] = index;
      }
      cells.push({ rowIndex: row, cellIndex, columnIndex: column, box: { x: xs[column]!, y: ys[row]!, width: xs[column + cell.colSpan]! - xs[column]!, height: ys[row + cell.rowSpan]! - ys[row]! } });
      column += cell.colSpan;
    }
  }
  const candidate = (index: number, side: TableCellEdge): Candidate | undefined => {
    const owner = owners[index]; if (!owner) return undefined;
    const explicit = owner.style.borders?.[side]; return { stroke: explicit ?? owner.style.stroke, explicit: explicit !== undefined };
  };
  const horizontal: NormalizedStrokeStyle[] = [], vertical: NormalizedStrokeStyle[] = [];
  for (let line = 0; line <= rows; line++) for (let column = 0; column < columns; column++) {
    budget.step(1, table); const above = line ? grid[(line - 1) * columns + column]! : -1, below = line < rows ? grid[line * columns + column]! : -1;
    horizontal.push(above >= 0 && above === below ? structuredClone(invisible) : winner(candidate(above, "bottom"), candidate(below, "top")));
  }
  for (let row = 0; row < rows; row++) for (let line = 0; line <= columns; line++) {
    budget.step(1, table); const left = line ? grid[row * columns + line - 1]! : -1, right = line < columns ? grid[row * columns + line]! : -1;
    vertical.push(left >= 0 && left === right ? structuredClone(invisible) : winner(candidate(left, "right"), candidate(right, "left")));
  }
  return { grid: { schemaVersion: "pptkit-table-border-grid-v1", rowCount: rows, columnCount: columns, horizontal, vertical }, cells };
}

/** Opt-in standalone resolution over normalized cells; no document mutation. */
export function resolveTableBorders(table: NormalizedTableElement, limits?: TableBorderLimits): TableBorderLayout { return resolveBorders(table, new TableBorderBudget(limits)); }
