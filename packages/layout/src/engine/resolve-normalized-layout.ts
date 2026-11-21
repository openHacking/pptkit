import { validateSlideEventTimeline } from "../validation/event-timeline.js";
import { refreshLayoutFields } from "../fields/refresh.js";
import { assertSizingFieldsStable, captureSizingFields } from "../fields/sizing-consistency.js";
import { connectorGeometry } from "../geometry/connector.js";
import { paginateTables, validatePagination } from "../table/paginate-tables.js";
import { resolveIntrinsic, validateIntrinsic } from "../sizing/resolve-intrinsic.js";
import { cellGrid } from "../table/cell-grid.js";
import { resolveBorders, TableBorderBudget } from "../table/resolve-borders.js";
import { sizeTables, validateTableSizing } from "../table/size-tables.js";
import { compose, validateComposition } from "../composition/compose.js";
import type {
  Box,
  NormalizedAsset,
  NormalizedElement,
  NormalizedPresentation,
} from "@pptkit/core";
import type { LayoutElement, LayoutGroupElement, LayoutResult } from "../types/layout.js";
import { resolveChartLayout } from "../chart/resolve-chart-layout.js";
import type { LayoutOptions, LayoutDiagnostic } from "../types/analysis.js";
import { analyzeResolvedLayout, validateLayoutOptions } from "../analysis/analyze-layout.js";

function resolveImage(element: Extract<NormalizedElement, { type: "image" }>, assets: ReadonlyMap<string, NormalizedAsset>): LayoutElement {
  const clone = structuredClone(element);
  const asset = assets.get(element.assetId);
  if (asset?.width === undefined || asset.height === undefined || element.box.width === 0 || element.box.height === 0) return clone;
  const sourceAspect = asset.width / asset.height;
  const targetAspect = element.box.width / element.box.height;
  if (![sourceAspect, targetAspect].every(value => Number.isFinite(value) && value > 0)) return clone;
  if (element.fit === "contain") {
    const box = sourceAspect > targetAspect
      ? { x: element.box.x, y: element.box.y + (element.box.height - element.box.width / sourceAspect) / 2, width: element.box.width, height: element.box.width / sourceAspect }
      : { x: element.box.x + (element.box.width - element.box.height * sourceAspect) / 2, y: element.box.y, width: element.box.height * sourceAspect, height: element.box.height };
    if (![box.x, box.y, box.width, box.height].every(Number.isFinite) || box.width <= 0 || box.height <= 0) return clone;
    clone.box = box;
    clone.fit = "stretch";
  } else if (element.fit === "cover") {
    const horizontal = sourceAspect > targetAspect;
    const visible = horizontal ? targetAspect / sourceAspect : sourceAspect / targetAspect;
    const edge = (1 - visible) / 2;
    // Avoid producing a zero visible area when floating-point rounding removes it.
    if (!Number.isFinite(visible) || visible <= 0 || edge + edge >= 1) return clone;
    clone.crop = horizontal ? { left: edge, right: edge, top: 0, bottom: 0 } : { left: 0, right: 0, top: edge, bottom: edge };
    clone.fit = "crop";
  }
  return clone;
}

function collectBoxes(elements: readonly NormalizedElement[], destination: Map<string, Box>): void {
  for (const element of elements) destination.set(element.id, element.box);
}

function resolveElements(elements: readonly NormalizedElement[], assets: ReadonlyMap<string, NormalizedAsset>, additional: readonly NormalizedElement[] = []): LayoutElement[] {
  const boxes = new Map<string, Box>();
  const transforms = new Map([...additional, ...elements].map(element => [element.id, element.transform]));
  collectBoxes(additional, boxes);
  collectBoxes(elements, boxes);
  return elements.map((element): LayoutElement => {
    if (element.type === "connector") {
      const connections = {
        ...("elementId" in element.start ? { start: { elementId: element.start.elementId, anchor: element.start.anchor ?? "center" } } : {}),
        ...("elementId" in element.end ? { end: { elementId: element.end.elementId, anchor: element.end.anchor ?? "center" } } : {}),
      };
      return { ...structuredClone(element), ...connectorGeometry(element, boxes, transforms), ...(Object.keys(connections).length ? { connections } : {}) };
    }
    if (element.type === "group") {
      const group: LayoutGroupElement = {
        ...structuredClone(element),
        children: resolveElements(element.children, assets),
      };
      return group;
    }
    if (element.type === "image") return resolveImage(element, assets);
    if (element.type === "chart") return { ...structuredClone(element), chartLayout: resolveChartLayout(element) };
    return structuredClone(element);
  });
}

export function resolveNormalizedLayout(normalized: NormalizedPresentation, options: LayoutOptions = {}): LayoutResult {
  for (const item of normalized.layouts) if (item.masterId !== undefined && (!item.theme || !(normalized.masters ?? []).some(master => master.id === item.masterId))) throw new TypeError(`Layout "${item.id}" lacks a resolved master/theme domain.`);
  const borderBudget = new TableBorderBudget(options.tableBorderLimits);
  const borderDiagnostics: LayoutDiagnostic[] = [];
  validateLayoutOptions(options);
  for (const slide of normalized.slides) validateSlideEventTimeline(slide);
  validateComposition(options.composition);
  validateTableSizing(options.tableSizing);
  validateIntrinsic(options.intrinsicSizing);
  validatePagination(options.tablePagination);
  const hasSizing = options.tableSizing !== undefined || options.intrinsicSizing !== undefined;
  if (options.composition !== undefined || options.tableSizing !== undefined || options.tablePagination !== undefined || (options.fieldRefresh !== undefined && hasSizing)) normalized = structuredClone(normalized);
  // Measure provisional field caches on a detached input. Pagination may change
  // their values, so compare the actual output against these snapshots below.
  if (options.fieldRefresh !== undefined && hasSizing) refreshLayoutFields(normalized, options.fieldRefresh);
  const sizingFields = options.fieldRefresh !== undefined && hasSizing && (options.tablePagination?.length ?? 0) > 0 ? captureSizingFields(normalized, options) : undefined;
  const tableSizing = options.tableSizing === undefined ? { diagnostics: [], resolved: new Set<string>() } : sizeTables(normalized, options.tableSizing, options.textMeasurement);
  const intrinsic = options.intrinsicSizing === undefined ? { document: normalized, diagnostics: [] } : resolveIntrinsic(normalized, options.intrinsicSizing, options.textMeasurement);
  normalized = intrinsic.document;
  const compositionDiagnostics = options.composition === undefined ? [] : compose(normalized, options.composition);
  const pagination = options.tablePagination === undefined ? { diagnostics: [], fragments: [], tables: new Set<string>() } : paginateTables(normalized, options.tablePagination);
  for (const id of pagination.tables) tableSizing.resolved.add(id);
  const assets = new Map(normalized.assets.map((asset) => [asset.id, asset]));
  const normalizedLayouts = new Map(normalized.layouts.map((layout) => [layout.id, layout]));
  const normalizedMasters = new Map((normalized.masters ?? []).map(master => [master.id, master]));
  const masters = normalized.masters?.map(({ elements, ...master }) => ({ ...structuredClone(master), ...(elements === undefined ? {} : { elements: resolveElements(elements, assets) }) }));
  const layouts = normalized.layouts.map((layout) => ({
    id: layout.id,
    name: layout.name,
    ...(layout.showMasterShapes === undefined ? {} : { showMasterShapes: layout.showMasterShapes }),
    ...(layout.backgroundSource !== undefined ? { backgroundSource: layout.backgroundSource } : {}),
    ...(layout.masterId !== undefined ? { masterId: layout.masterId } : {}),
    ...(layout.theme !== undefined ? { theme: structuredClone(layout.theme) } : {}),
    background: structuredClone(layout.background),
    elements: resolveElements(layout.elements, assets, layout.masterId === undefined ? [] : normalizedMasters.get(layout.masterId)?.elements ?? []),
    placeholders: structuredClone(layout.placeholders),
  }));
  const result: LayoutResult = {
    ...(normalized.sections === undefined ? {} : { sections: structuredClone(normalized.sections) }),
    ...(options.tablePagination !== undefined ? { pagination: { fragments: pagination.fragments } } : {}),
    size: structuredClone(normalized.size),
    ...(normalized.firstSlideNumber === undefined ? {} : { firstSlideNumber: normalized.firstSlideNumber }),
    metadata: structuredClone(normalized.metadata),
    theme: structuredClone(normalized.theme),
    ...(masters === undefined ? {} : { masters }),
    assets: structuredClone(normalized.assets),
    layouts,
    slides: normalized.slides.map((slide) => ({
      id: slide.id,
      ...(slide.showInheritedShapes === undefined ? {} : { showInheritedShapes: slide.showInheritedShapes }),
      layoutId: slide.layoutId,
      background: structuredClone(slide.background),
      backgroundSource: slide.backgroundSource,
      elements: resolveElements(slide.elements, assets, [...(normalizedMasters.get(normalizedLayouts.get(slide.layoutId)?.masterId ?? "")?.elements ?? []), ...(normalizedLayouts.get(slide.layoutId)?.elements ?? [])]),
      notes: structuredClone(slide.notes),
      ...(slide.transition !== undefined ? { transition: structuredClone(slide.transition) } : {}),
      ...(slide.readingOrder === undefined ? {} : { readingOrder: structuredClone(slide.readingOrder) }),
      ...(slide.eventTimeline !== undefined ? { eventTimeline: structuredClone(slide.eventTimeline) } : {}),
      hidden: slide.hidden,
      ...(slide.section !== undefined ? { section: slide.section } : {}),
      tags: [...slide.tags],
      customData: structuredClone(slide.customData),
    })),
    slideCount: normalized.slides.length,
    status: "resolved",
  };
  if (options.fieldRefresh !== undefined) refreshLayoutFields(result, options.fieldRefresh);
  if (sizingFields) assertSizingFieldsStable(result, sizingFields);
  function attachCellGeometry(elements: LayoutElement[], identity: { slideId?: string; layoutId?: string; masterId?: string }): void {
    for (const element of elements) {
      if (element.type === "group") attachCellGeometry(element.children, identity);
      if (element.type === "table") {
        const rotatedText = element.rows.some(row => row.cells.some(cell => cell.style.textDirection !== undefined && cell.style.textDirection !== "horizontal"));
        if (element.rows.some(row => row.cells.some(cell => cell.style.borders !== undefined || cell.style.stroke.dash !== "solid"))) {
          const resolved = resolveBorders(element, borderBudget); element.borderGrid = resolved.grid; element.cellLayout = resolved.cells;
          const width = element.columns.reduce((sum, value) => sum + value, 0), height = element.rows.reduce((sum, row) => sum + row.height, 0);
          if (Math.abs(width - element.box.width) > .01 || Math.abs(height - element.box.height) > .01) borderDiagnostics.push({ code: "table-border-frame-grid-mismatch", severity: "warning", message: "Border coordinates follow grid dimensions while the frame differs; native table resizing and preview clipping may differ. Fit the frame to the grid for aligned output.", path: `elements.${element.id}.box`, elementId: element.id, ...identity, bounds: { ...element.box, width, height }, available: { ...element.box } });
        } else if (tableSizing.resolved.has(element.id) || rotatedText) {
          const grid = cellGrid(element);
          if (grid.invalid.length === 0 || !rotatedText) element.cellLayout = grid.cells;
          else borderDiagnostics.push({ code: "table-text-direction-grid-invalid", severity: "warning", message: "Rotated table-cell preview could not resolve the merged grid; its layout may be inaccurate.", path: `elements.${element.id}.rows`, elementId: element.id, ...identity });
          if (rotatedText && grid.invalid.length === 0) {
            const width = element.columns.reduce((sum, value) => sum + value, 0), height = element.rows.reduce((sum, row) => sum + row.height, 0);
            if (Math.abs(width - element.box.width) > .01 || Math.abs(height - element.box.height) > .01) borderDiagnostics.push({ code: "table-text-direction-frame-grid-mismatch", severity: "warning", message: "Rotated cell preview follows grid dimensions while the table frame differs; native resizing and preview clipping may differ.", path: `elements.${element.id}.box`, elementId: element.id, ...identity, bounds: { ...element.box, width, height }, available: { ...element.box } });
          }
        }
      }
    }
  }
  for (const master of result.masters ?? []) attachCellGeometry(master.elements ?? [], { masterId: master.id });
  for (const layout of result.layouts) attachCellGeometry(layout.elements, { layoutId: layout.id });
  for (const slide of result.slides) attachCellGeometry(slide.elements, { slideId: slide.id });
  if (borderDiagnostics.length || options.textMeasurement !== undefined || options.diagnostics !== undefined || options.composition !== undefined || options.tableSizing !== undefined || options.intrinsicSizing !== undefined || options.tablePagination !== undefined) {
    result.analysis = analyzeResolvedLayout(result, options);
    result.analysis.diagnostics = [
      ...tableSizing.diagnostics, ...intrinsic.diagnostics, ...compositionDiagnostics,
      ...pagination.diagnostics, ...borderDiagnostics, ...result.analysis.diagnostics,
    ];
    if (options.tableSizing !== undefined || options.intrinsicSizing !== undefined) {
      const seen = new Set<string>();
      result.analysis.diagnostics = result.analysis.diagnostics.filter((diagnostic) => {
        const key = JSON.stringify(diagnostic);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    }
  }
  return result;
}
