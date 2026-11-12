import { ownerTheme } from "../analysis/owner-theme.js";
import { cellGrid } from "./cell-grid.js";
import type { NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import type { LayoutDiagnostic, TextMeasurementOptions } from "../types/analysis.js";
import type { TableSizingRule } from "../types/table.js";
import { measureTextBody, type MeasurementContext } from "../analysis/measure-text.js";

type Table = Extract<NormalizedElement, { type: "table" }>;

export function validateTableSizing(rules: readonly TableSizingRule[] | undefined): void {
  if (rules === undefined) return;
  if (!Array.isArray(rules)) throw new TypeError("tableSizing must be an array");
  const ids = new Set<string>();
  for (const rule of rules as readonly TableSizingRule[]) {
    if (!rule || typeof rule.elementId !== "string" || ids.has(rule.elementId)) throw new TypeError("Table sizing requires unique element IDs");
    ids.add(rule.elementId);
    if (!Array.isArray(rule.autoRows)) throw new TypeError("autoRows must be an array");
    const rows = new Set<number>();
    for (const row of rule.autoRows) {
      if (!row || !Number.isSafeInteger(row.row) || row.row < 0 || rows.has(row.row)) throw new RangeError("Automatic row indices must be unique nonnegative integers");
      rows.add(row.row);
      if (row.minHeight !== undefined && typeof row.minHeight !== "number") throw new TypeError("Automatic minimum height must be a number");
      const minimum = row.minHeight ?? 1;
      if (!Number.isFinite(minimum) || minimum <= 0 || (row.maxHeight !== undefined && (!Number.isFinite(row.maxHeight) || row.maxHeight < minimum))) throw new RangeError("Automatic row bounds must be finite, positive and ordered");
    }
    if (rule.tolerance !== undefined && (!Number.isFinite(rule.tolerance) || rule.tolerance < 0)) throw new RangeError("Table tolerance must be finite and nonnegative");
    if (rule.overflow !== undefined && rule.overflow !== "preserve" && rule.overflow !== "clip") throw new TypeError("Table overflow must be preserve or clip");
    if (rule.rotatedTextSizing !== undefined && rule.rotatedTextSizing !== "conservative" && rule.rotatedTextSizing !== "compact") throw new TypeError("Rotated text sizing must be conservative or compact");
    if (rule.includeHidden !== undefined && typeof rule.includeHidden !== "boolean") throw new TypeError("Table includeHidden must be boolean");
  }
}

interface Constraint { row: number; span: number; required: number; context: MeasurementContext }

/** Operates only on the private normalized clone owned by the Layout orchestrator. */
export function sizeTables(document: NormalizedPresentation, rules: readonly TableSizingRule[], provider: TextMeasurementOptions | undefined): { diagnostics: LayoutDiagnostic[]; resolved: Set<string> } {
  const themeFor = ownerTheme(document);
  const diagnostics: LayoutDiagnostic[] = [];
  const resolved = new Set<string>();
  const tables = new Map<string, { table: Table; hidden: boolean; context: MeasurementContext }>();
  function walk(elements: readonly NormalizedElement[], path: string, owner: { slideId?: string; layoutId?: string; masterId?: string }, hidden: boolean): void {
    elements.forEach((element, index) => {
      const context = { ...owner, elementId: element.id, path: `${path}.${index}` };
      if (element.type === "table") tables.set(element.id, { table: element, hidden: hidden || element.hidden, context });
      if (element.type === "group") walk(element.children, `${context.path}.children`, owner, hidden || element.hidden);
    });
  }
  (document.masters ?? []).forEach((master, i) => walk(master.elements ?? [], `masters.${i}.elements`, { masterId: master.id }, false));
  document.layouts.forEach((layout, i) => walk(layout.elements, `layouts.${i}.elements`, { layoutId: layout.id }, false));
  document.slides.forEach((slide, i) => walk(slide.elements, `slides.${i}.elements`, { slideId: slide.id }, slide.hidden));
  rules.forEach((rule, index) => {
    const found = tables.get(rule.elementId);
    const context = found?.context ?? { elementId: rule.elementId, path: `options.tableSizing.${index}` };
    const warn = (code: string, message: string, location = context): void => { diagnostics.push({ ...location, severity: "warning", code, message }); };
    if (!found) { warn("table-sizing-target-missing", "Table sizing target was not found."); return; }
    if (found.hidden && !rule.includeHidden) return;
    const table = found.table;
    const tolerance = rule.tolerance ?? 0.02;
    if (rule.autoRows.some(({ row }) => row >= table.rows.length)) { warn("table-sizing-row-missing", "An automatic row index exceeds the table. Authored geometry was retained."); return; }
    if (!provider) { warn("text-measurement-unavailable", "Table sizing requires a text measurement provider. Authored geometry was retained."); return; }
    const textProvider = provider;
    const width = table.columns.reduce((sum, value) => sum + value, 0);
    if (!Number.isFinite(width) || Math.abs(width - table.box.width) > tolerance) { warn("table-geometry-mismatch", "Table width must match explicit column widths before sizing rows. Authored geometry was retained."); return; }
    const heights = table.rows.map((row) => row.height);
    const maximum = table.rows.map((row) => row.height);
    const automatic = new Set<number>();
    for (const row of rule.autoRows) { heights[row.row] = row.minHeight ?? 1; maximum[row.row] = row.maxHeight ?? Infinity; automatic.add(row.row); }
    function measureVertical(cell: Table["rows"][number]["cells"][number], inline: number, cross: number, wrap: boolean, location: MeasurementContext) {
      const direction = cell.style.textDirection;
      const vertical = direction !== undefined && direction !== "horizontal" && direction !== "rotate90" && direction !== "rotate270";
      const rightToLeft = direction === "eastAsianVertical" || direction === "stackedVerticalRtl";
      const content = wrap ? cell.content : cell.content.map(paragraph => ({ ...paragraph, style: { ...paragraph.style, align: "left" as const } }));
      const bounds = measureTextBody(content, { ...(vertical ? { textDirection: direction } : {}),
        margin: cell.style.margin, verticalAlign: wrap ? cell.style.verticalAlign : "top", wrap, autoFit: { mode: "none" } },
        { x: 0, y: 0, width: vertical ? cross : inline, height: vertical ? inline : cross }, themeFor(location), textProvider, location, diagnostics);
      return vertical && bounds ? { x: bounds.y, y: rightToLeft ? cross - bounds.x - bounds.width : bounds.x, width: bounds.height, height: bounds.width } : bounds;
    }
    const constraints: Constraint[] = [];
    const rotatedChecks: { row: number; span: number; width: number; cell: Table["rows"][number]["cells"][number]; context: MeasurementContext }[] = [];
    let failed = false;
    let overflow = false;
    const grid = cellGrid(table);
    if (grid.invalid.length) {
      for (const { rowIndex, cellIndex } of grid.invalid) warn("table-span-invalid", "Invalid merged cell grid prevents table sizing. Authored geometry was retained.", { ...context, path: `${context.path}.rows.${rowIndex}.cells.${cellIndex}.content` });
      return;
    }
    for (const { rowIndex, cellIndex, box } of grid.cells) {
      const cell = table.rows[rowIndex]!.cells[cellIndex]!;
      const location = { ...context, path: `${context.path}.rows.${rowIndex}.cells.${cellIndex}.content` };
      const margin = cell.style.margin;
      const availableWidth = Math.max(0, box.width - margin.left - margin.right);
      if (cell.style.textDirection !== undefined && cell.style.textDirection !== "horizontal") {
        // Use the unwrapped inline extent as a conservative physical-height
        // constraint. Final dimensions are measured again with wrapping below.
        const bounds = measureVertical(cell, 0, 0, false, location);
        if (!bounds) { failed = true; continue; }
        warn("table-text-direction-sizing-approximate", "Vertical cell sizing uses an unwrapped probe and final remeasurement; native text layout can differ.", location);
        if (margin.left + margin.right > box.width + tolerance || bounds.y < -tolerance || bounds.y + bounds.height > availableWidth + tolerance) {
          overflow = true;
          warn("table-cell-overflow", "Vertical text exceeds the fixed cell width; automatic row height cannot resolve it.", location);
        }
        let required = Math.max(0, bounds.x + bounds.width) + margin.top + margin.bottom;
        if (!Number.isFinite(required)) { warn("table-sizing-invalid", "Measured row height exceeded finite range. Authored geometry was retained.", location); failed = true; continue; }
        const spanMaximum = maximum.slice(rowIndex, rowIndex + cell.rowSpan).reduce((sum, value) => sum + value, 0);
        const spanMinimum = heights.slice(rowIndex, rowIndex + cell.rowSpan).reduce((sum, value) => sum + value, 0);
        const upperCandidate = Math.min(spanMaximum, Math.max(spanMinimum, required));
        if (Number.isFinite(upperCandidate) && (rule.rotatedTextSizing === "compact" || required > spanMaximum + tolerance) && margin.top + margin.bottom <= upperCandidate + tolerance) {
          const inline = Math.max(0, upperCandidate - margin.top - margin.bottom);
          const wrapped = measureVertical(cell, inline, availableWidth, true, location);
          if (!wrapped) { failed = true; continue; }
          if (wrapped.x >= -tolerance && wrapped.y >= -tolerance && wrapped.x + wrapped.width <= inline + tolerance && wrapped.y + wrapped.height <= availableWidth + tolerance) {
            // Keep a known-fitting upper bound. Measurements need not be
            // monotonic, so this is a bounded search, not a global optimum.
            let lower = spanMinimum, upper = upperCandidate;
            for (let attempt = 0; attempt < 16 && upper - lower > tolerance; attempt++) {
              const trial = lower + (upper - lower) / 2, trialInline = Math.max(0, trial - margin.top - margin.bottom);
              if (margin.top + margin.bottom > trial + tolerance) { lower = trial; continue; }
              const measured = measureVertical(cell, trialInline, availableWidth, true, location);
              if (!measured) { failed = true; break; }
              if (measured.x >= -tolerance && measured.y >= -tolerance && measured.x + measured.width <= trialInline + tolerance && measured.y + measured.height <= availableWidth + tolerance) upper = trial;
              else lower = trial;
            }
            if (failed) continue;
            // The constraint solver may stop within tolerance of its target.
            required = Math.min(upperCandidate, upper + tolerance);
            warn("table-text-direction-wrapped-fallback", "Vertical text fit after a bounded wrapped search; the selected height is remeasured and is not guaranteed minimal.", location);
          }
        }
        constraints.push({ row: rowIndex, span: cell.rowSpan, required, context: location });
        rotatedChecks.push({ row: rowIndex, span: cell.rowSpan, width: box.width, cell, context: location });
        continue;
      }
      const bounds = measureTextBody(cell.content, { margin, verticalAlign: "top", wrap: true, autoFit: { mode: "none" } },
        { x: 0, y: 0, width: availableWidth, height: 0 }, themeFor(location), provider, location, diagnostics);
      if (!bounds) { failed = true; continue; }
      if (margin.left + margin.right > box.width + tolerance || bounds.x < -tolerance || bounds.y < -tolerance || bounds.x + bounds.width > availableWidth + tolerance) {
        overflow = true;
        warn("table-cell-overflow", "Cell text exceeds its fixed width or extends before its content origin; automatic row height cannot resolve it.", location);
      }
      const required = Math.max(0, bounds.y + bounds.height) + margin.top + margin.bottom;
      if (!Number.isFinite(required)) { warn("table-sizing-invalid", "Measured row height exceeded finite range. Authored geometry was retained.", location); failed = true; continue; }
      constraints.push({ row: rowIndex, span: cell.rowSpan, required, context: location });
    }
    if (failed) return;
    // Solve shorter spans first. Increasing heights never invalidates an earlier
    // lower-bound constraint; excess is shared equally among uncapped auto rows.
    constraints.sort((a, b) => a.span - b.span || a.row - b.row);
    for (const constraint of constraints) {
      const indices = Array.from({ length: constraint.span }, (_, i) => constraint.row + i);
      let deficit = constraint.required - indices.reduce((sum, i) => sum + heights[i]!, 0);
      let candidates = indices.filter((i) => automatic.has(i) && heights[i]! < maximum[i]!);
      while (deficit > tolerance && candidates.length) {
        const share = deficit / candidates.length;
        let growth = 0;
        for (const i of candidates) {
          const added = Math.min(share, maximum[i]! - heights[i]!);
          heights[i]! += added;
          growth += added;
        }
        if (growth <= 0) break;
        deficit -= growth;
        candidates = candidates.filter((i) => heights[i]! < maximum[i]!);
      }
      if (deficit > tolerance) {
        overflow = true;
        warn("table-row-overflow", "Fixed heights or automatic row maxima cannot contain this cell. Increase row limits or shorten its content.", constraint.context);
      }
    }
    if (overflow && rule.overflow !== "clip") return;
    for (const check of rotatedChecks) {
      const margin = check.cell.style.margin;
      const physicalHeight = heights.slice(check.row, check.row + check.span).reduce((sum, value) => sum + value, 0);
      const inline = Math.max(0, physicalHeight - margin.top - margin.bottom), cross = Math.max(0, check.width - margin.left - margin.right);
      const bounds = measureVertical(check.cell, inline, cross, true, check.context);
      if (!bounds) { failed = true; continue; }
      if (margin.top + margin.bottom > physicalHeight + tolerance || bounds.x < -tolerance || bounds.y < -tolerance || bounds.x + bounds.width > inline + tolerance || bounds.y + bounds.height > cross + tolerance) {
        overflow = true;
        warn("table-cell-overflow", "Vertical text exceeds the resolved cell content area after remeasurement.", check.context);
      }
    }
    if (failed || (overflow && rule.overflow !== "clip")) return;
    const height = heights.reduce((sum, value) => sum + value, 0);
    if (!Number.isFinite(height) || !Number.isFinite(table.box.y + height)) { warn("table-sizing-invalid", "Resolved table height exceeded finite range. Authored geometry was retained."); return; }
    table.rows.forEach((row, i) => { row.height = heights[i]!; });
    table.box.height = height;
    resolved.add(table.id);
  });
  return { diagnostics, resolved };
}
