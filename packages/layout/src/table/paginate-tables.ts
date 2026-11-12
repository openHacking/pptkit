import type { Box, NormalizedElement, NormalizedPresentation, NormalizedTableElement } from "@pptkit/core";
import type { LayoutDiagnostic } from "../types/analysis.js";
import type { TablePaginationFragment, TablePaginationRule } from "../types/pagination.js";
import { cellGrid } from "./cell-grid.js";

const finiteArea = (area: Box): boolean => !!area && [area.x, area.y, area.width, area.height, area.x + area.width, area.y + area.height].every(Number.isFinite) && area.width > 0 && area.height > 0;
const fits = (size: number, available: number): boolean => size <= available + Number.EPSILON * 16 * Math.max(1, size, available);

export function validatePagination(rules: readonly TablePaginationRule[] | undefined): void {
  if (rules === undefined) return;
  if (!Array.isArray(rules)) throw new TypeError("tablePagination must be an array");
  const ids = new Set<string>();
  for (const rule of rules) {
    if (!rule || typeof rule.elementId !== "string" || ids.has(rule.elementId)) throw new TypeError("Pagination requires unique table IDs");
    ids.add(rule.elementId);
    if (!finiteArea(rule.area)) throw new RangeError("Pagination areas require finite coordinates and positive dimensions");
    if (rule.continuation !== undefined) {
      if (!rule.continuation || typeof rule.continuation !== "object" || Array.isArray(rule.continuation)) throw new TypeError("Invalid continuation settings");
      if (rule.continuation.layoutId !== undefined && typeof rule.continuation.layoutId !== "string") throw new TypeError("Continuation layout ID must be a string");
      if (rule.continuation.area !== undefined && !finiteArea(rule.continuation.area)) throw new RangeError("Invalid continuation area");
    }
    if (rule.headerRows !== undefined && (!Number.isSafeInteger(rule.headerRows) || rule.headerRows < 0)) throw new RangeError("Header row count must be a nonnegative integer");
    if (rule.maxPages !== undefined && (!Number.isSafeInteger(rule.maxPages) || rule.maxPages < 1)) throw new RangeError("Pagination maxPages must be a positive integer");
    if (rule.overflow !== undefined && !["preserve", "overflow"].includes(rule.overflow)) throw new TypeError("Invalid pagination overflow policy");
    if (rule.nativeRowRegions !== undefined && !["fragment", "source-boundaries"].includes(rule.nativeRowRegions)) throw new TypeError("Invalid native row-region policy");
    if (rule.includeHidden !== undefined && typeof rule.includeHidden !== "boolean") throw new TypeError("Pagination includeHidden must be boolean");
  }
}

interface Page { rows: number[]; height: number; oversized: boolean }
type Plan = { pages: Page[] } | { code: string; message: string };

/** Split only between closed rowspan intervals. No document identity or mutation here. */
export function planTablePages(table: NormalizedTableElement, rule: TablePaginationRule): Plan {
  const grid = cellGrid(table);
  if (grid.invalid.length || !table.rows.length || !table.columns.length || table.rows.some((row) => !Number.isFinite(row.height) || row.height < 0) || table.columns.some((width) => !Number.isFinite(width) || width <= 0)) return { code: "pagination-table-invalid", message: "Pagination requires a valid nonempty table with finite row heights, columns and merged cells." };
  const headers = rule.headerRows ?? 0;
  if (headers > table.rows.length) return { code: "pagination-header-invalid", message: "Header row count exceeds table rows." };
  const forbidden = new Set<number>();
  for (const cell of grid.cells) {
    const span = table.rows[cell.rowIndex]!.cells[cell.cellIndex]!.rowSpan;
    for (let cut = cell.rowIndex + 1; cut < cell.rowIndex + span; cut++) forbidden.add(cut);
  }
  if (forbidden.has(headers)) return { code: "pagination-header-merged", message: "A merged cell crosses the header/body boundary; choose a closed header range." };
  const continuation = rule.continuation?.area ?? rule.area;
  const width = table.columns.reduce((sum, column) => sum + column, 0);
  if (!Number.isFinite(width) || !fits(width, rule.area.width) || !fits(width, continuation.width) || Math.abs(width - table.box.width) > Number.EPSILON * 16 * Math.max(1, width, table.box.width)) return { code: "pagination-width-unavailable", message: "Table columns must match the frame width and fit both areas. Pagination does not resize or reflow columns." };
  const headerIndices = Array.from({ length: headers }, (_, i) => i);
  const headerHeight = table.rows.slice(0, headers).reduce((sum, row) => sum + row.height, 0);
  if (!Number.isFinite(headerHeight)) return { code: "pagination-height-invalid", message: "Header height exceeds finite coordinates." };
  const blocks: { start: number; end: number; height: number }[] = [];
  for (let start = headers; start < table.rows.length;) {
    let end = start + 1;
    while (forbidden.has(end)) end++;
    const height = table.rows.slice(start, end).reduce((sum, row) => sum + row.height, 0);
    if (!Number.isFinite(height)) return { code: "pagination-height-invalid", message: "Merged row block height exceeds finite coordinates." };
    blocks.push({ start, end, height }); start = end;
  }
  const pages: Page[] = [];
  let current: Page = { rows: [...headerIndices], height: headerHeight, oversized: false };
  const capacity = (): number => pages.length ? continuation.height : rule.area.height;
  for (const block of blocks) {
    if (current.rows.length > headers && !fits(current.height + block.height, capacity())) {
      pages.push(current); current = { rows: [...headerIndices], height: headerHeight, oversized: false };
    }
    const height = current.height + block.height;
    if (!Number.isFinite(height)) return { code: "pagination-height-invalid", message: "Table fragment height exceeds finite coordinates." };
    if (!fits(height, capacity())) {
      if (rule.overflow !== "overflow") return { code: "pagination-row-too-tall", message: `Source rows ${block.start}–${block.end - 1} and repeated headers cannot fit an empty page. No pagination changes were applied for this table.` };
      current.oversized = true;
    }
    for (let row = block.start; row < block.end; row++) current.rows.push(row);
    current.height = height;
    if (pages.length + 1 > (rule.maxPages ?? 1000)) return { code: "pagination-page-limit", message: "Table pagination exceeds maxPages. No pagination changes were applied for this table." };
  }
  if (!blocks.length && !fits(current.height, capacity())) {
    if (rule.overflow !== "overflow") return { code: "pagination-row-too-tall", message: "The header-only table cannot fit the first area." };
    current.oversized = true;
  }
  pages.push(current);
  return { pages };
}

/** Mutates only the caller's private normalized clone, committing each table atomically. */
export function paginateTables(document: NormalizedPresentation, rules: readonly TablePaginationRule[]): { diagnostics: LayoutDiagnostic[]; fragments: TablePaginationFragment[]; tables: Set<string> } {
  const diagnostics: LayoutDiagnostic[] = [];
  const fragments: TablePaginationFragment[] = [];
  const tables = new Set<string>();
  const layouts = new Map(document.layouts.map((layout) => [layout.id, layout]));
  const reserved = new Set([document.id, ...document.slides.map((slide) => slide.id), ...document.layouts.map((layout) => layout.id), ...document.assets.map((asset) => asset.id), ...(document.masters ?? []).map(master => master.id)]);
  function collect(elements: readonly NormalizedElement[]): void { for (const element of elements) { reserved.add(element.id); if (element.type === "group") collect(element.children); } }
  (document.masters ?? []).forEach(master => collect(master.elements ?? [])); document.slides.forEach((slide) => collect(slide.elements)); document.layouts.forEach((layout) => collect(layout.elements));
  const targets = new Map(document.slides.flatMap((slide) => slide.elements.map((element) => [element.id, { slide, element }] as const)));
  const counts = new Map<string, number>();
  for (const rule of rules) { const target = targets.get(rule.elementId); if (target?.element.type === "table" && (rule.includeHidden || (!target.slide.hidden && !target.element.hidden))) counts.set(target.slide.id, (counts.get(target.slide.id) ?? 0) + 1); }
  rules.forEach((rule, ruleIndex) => {
    const target = targets.get(rule.elementId);
    const context = { path: `options.tablePagination.${ruleIndex}`, elementId: rule.elementId, ...(target ? { slideId: target.slide.id } : {}) };
    const warn = (code: string, message: string): void => { diagnostics.push({ ...context, severity: "warning", code, message }); };
    if (!target || target.element.type !== "table") { warn("pagination-target-unavailable", "Pagination targets must be top-level tables on source slides, not layout/group children."); return; }
    const { slide } = target;
    const table = target.element;
    if ((slide.hidden || table.hidden) && !rule.includeHidden) return;
    if (counts.get(slide.id)! > 1) { warn("pagination-source-conflict", "Only one paginated table per source slide is supported; none of this slide's pagination rules were applied."); return; }
    if (table.transform.rotation !== 0 || table.transform.flipH || table.transform.flipV) { warn("pagination-transform-unsupported", "Paginated tables must be unrotated and unflipped; pagination areas use slide coordinates."); return; }
    const layoutId = rule.continuation?.layoutId ?? slide.layoutId;
    const layout = layouts.get(layoutId);
    if (!layout) { warn("pagination-layout-missing", `Continuation layout "${layoutId}" was not found.`); return; }
    const plan = planTablePages(table, rule);
    if (!("pages" in plan)) { warn(plan.code, plan.message); return; }
    if (slide.eventTimeline !== undefined && plan.pages.length > 1) { warn("pagination-event-timeline-unsupported", "Splitting a timed slide requires explicit timeline/target reconciliation; no pagination changes were applied."); return; }
    const ids = plan.pages.map((_, i) => i === 0 ? { slideId: slide.id, elementId: table.id } : { slideId: `pptkit:page:${JSON.stringify([slide.id, table.id, i])}`, elementId: `pptkit:table:${JSON.stringify([slide.id, table.id, i])}` });
    if (ids.slice(1).some((id) => reserved.has(id.slideId) || reserved.has(id.elementId))) { warn("pagination-id-conflict", "A deterministic continuation ID already exists. Rename the conflicting object; no pagination changes were applied."); return; }
    // Clone only the rows that belong to each fragment, not the entire source table per page.
    const { rows: _sourceRows, ...tableMetadata } = table;
    const parts = plan.pages.map((page, i) => {
      const area = i === 0 ? rule.area : rule.continuation?.area ?? rule.area;
      const part: NormalizedTableElement = { ...structuredClone(tableMetadata), id: ids[i]!.elementId, rows: page.rows.map((index) => structuredClone(table.rows[index]!)), box: { x: area.x, y: area.y, width: table.box.width, height: page.height } };
      if (i > 0) delete part.placeholderKey;
      if (plan.pages.length > 1 && rule.nativeRowRegions === "source-boundaries") {
        if (i > 0 && (rule.headerRows ?? 0) === 0 && table.nativeFirstRowStyle !== false) part.nativeFirstRowStyle = false;
        if (i < plan.pages.length - 1 && table.nativeTableStyleFlags?.lastRow === true) part.nativeTableStyleFlags = { ...part.nativeTableStyleFlags, lastRow: false };
      }
      return part;
    });
    if (parts.some((part) => !Number.isFinite(part.box.x + part.box.width) || !Number.isFinite(part.box.y + part.box.height))) { warn("pagination-height-invalid", "Paginated frame extents exceed finite coordinates."); return; }
    const { elements: _sourceElements, notes: _sourceNotes, readingOrder: _sourceReadingOrder, ...slideMetadata } = slide;
    const continuations = parts.slice(1).map((part, i) => ({
      ...structuredClone(slideMetadata), id: ids[i + 1]!.slideId, layoutId,
      background: structuredClone(rule.continuation?.layoutId === undefined ? slide.background : layout.background),
      backgroundSource: rule.continuation?.layoutId === undefined ? slide.backgroundSource : "layout" as const,
      elements: [part], notes: [],
    }));
    slide.elements[slide.elements.indexOf(table)] = parts[0]!;
    const insertion = document.slides.indexOf(slide) + 1;
    document.slides = document.slides.slice(0, insertion).concat(continuations, document.slides.slice(insertion));
    const explicitNativeStyle = table.nativeFirstRowStyle !== undefined || table.nativeTableStyleFlags !== undefined || table.nativeTableStylePreset !== undefined;
    const nativeRowRegions = table.nativeFirstRowStyle !== false || table.nativeTableStyleFlags?.lastRow === true || table.nativeTableStyleFlags?.bandedRows !== false;
    // Even without authored metadata, PPTX export enables firstRow and bandRow.
    // Repeated headers retain their original positions; body rows may change
    // parity when a fragment begins after an odd number of source rows. For
    // merged rows, do not infer native band placement from an anchor alone.
    const implicitNativeRisk = !explicitNativeStyle && ((rule.headerRows ?? 0) === 0 || table.rows.some(row => row.cells.some(cell => cell.rowSpan > 1)) || plan.pages.some(page => page.rows.some((sourceRow, localRow) => (sourceRow - localRow) % 2 !== 0)));
    if (parts.length > 1 && ((rule.nativeRowRegions === "source-boundaries" && table.nativeTableStyleFlags?.bandedRows !== false) || (rule.nativeRowRegions !== "source-boundaries" && ((explicitNativeStyle && nativeRowRegions) || implicitNativeRisk)))) warn("pagination-native-row-style-reapplied", rule.nativeRowRegions === "source-boundaries" ? "Native first/last-row switches follow source boundaries; native row bands may still restart on each fragment. Core materialized cell formatting remains attached to its source rows." : "Native table row-style switches, including exporter defaults, are copied to each fragment; an application may recalculate first/last rows and bands per fragment. Core materialized cell formatting remains attached to its source rows.");
    plan.pages.forEach((page, i) => {
      const id = ids[i]!; reserved.add(id.slideId); reserved.add(id.elementId); tables.add(id.elementId);
      fragments.push({ sourceSlideId: slide.id, sourceElementId: table.id, ...id, pageIndex: i, sourceRowIndices: [...page.rows], repeatedHeaderRowCount: i === 0 ? 0 : rule.headerRows ?? 0 });
      if (page.oversized) diagnostics.push({ ...context, slideId: id.slideId, elementId: id.elementId, code: "pagination-row-overflow", severity: "warning", message: `Fragment ${i} contains an oversized indivisible row block and may extend outside its area.` });
    });
  });
  return { diagnostics, fragments, tables };
}
