import { ownerTheme } from "./owner-theme.js";
import { cellGrid } from "../table/cell-grid.js";
import type { Box, NormalizedTextFrameStyle, NormalizedTextParagraph } from "@pptkit/core";
import type { LayoutElement, LayoutResult } from "../types/layout.js";
import type { LayoutAnalysis, LayoutOptions } from "../types/analysis.js";
import { elementTransform, identity, multiply, overflows, transformedBounds, type Matrix } from "./geometry.js";
import { measureTextBody, type MeasurementContext } from "./measure-text.js";
import { resolveShapeTextBox } from "../shape/resolve-shape-text-box.js";

export function validateLayoutOptions(options: LayoutOptions): void {
  if (options.textMeasurement !== undefined) {
    if (typeof options.textMeasurement?.measure !== "function") throw new TypeError("textMeasurement.measure must be a synchronous function");
    if (options.textMeasurement.measureVertical !== undefined && typeof options.textMeasurement.measureVertical !== "function") throw new TypeError("textMeasurement.measureVertical must be a function");
    if (options.textMeasurement.measureEastAsianVertical !== undefined && typeof options.textMeasurement.measureEastAsianVertical !== "function") throw new TypeError("textMeasurement.measureEastAsianVertical must be a function");
    if (options.textMeasurement.resolveFont !== undefined && typeof options.textMeasurement.resolveFont !== "function") throw new TypeError("textMeasurement.resolveFont must be a function");
  }
  const diagnostics = options.diagnostics;
  if (diagnostics === null) throw new TypeError("diagnostics must be an options object");
  if (diagnostics?.tolerance !== undefined && (!Number.isFinite(diagnostics.tolerance) || diagnostics.tolerance < 0)) throw new RangeError("diagnostics.tolerance must be finite and nonnegative");
  for (const flag of [diagnostics?.textOverflow, diagnostics?.slideBounds, diagnostics?.shadowBounds, diagnostics?.includeHidden]) if (flag !== undefined && typeof flag !== "boolean") throw new TypeError("Diagnostic flags must be booleans");
  const allowed = diagnostics?.allowedOffSlideElementIds;
  if (allowed !== undefined && (!Array.isArray(allowed) || allowed.some((id) => typeof id !== "string"))) throw new TypeError("allowedOffSlideElementIds must contain strings");
}

export function analyzeResolvedLayout(layout: LayoutResult, options: LayoutOptions): LayoutAnalysis {
  const themeFor = ownerTheme(layout);
  const result: LayoutAnalysis = { measurements: [], diagnostics: [] };
  const tolerance = options.diagnostics?.tolerance ?? 0.01;
  const allowed = new Set(options.diagnostics?.allowedOffSlideElementIds ?? []);
  const slideBox = { x: 0, y: 0, width: layout.size.width, height: layout.size.height };

  function boundary(bounds: Box, context: MeasurementContext, suppressed: boolean): void {
    if (!options.diagnostics?.slideBounds || suppressed) return;
    if (![bounds.x, bounds.y, bounds.width, bounds.height, bounds.x + bounds.width, bounds.y + bounds.height].every(Number.isFinite)) {
      result.diagnostics.push({ ...context, code: "geometry-unmeasurable", severity: "warning", message: "Transformed bounds exceed finite numeric range. Reduce coordinates or nested scaling before checking slide boundaries." });
      return;
    }
    if (overflows(bounds, slideBox, tolerance)) {
      result.diagnostics.push({ ...context, code: "slide-overflow", severity: "warning", message: "Transformed bounds extend outside the slide. Move or resize the element, or explicitly allow this off-slide element.", bounds, available: { ...slideBox } });
    }
  }

  function text(content: readonly NormalizedTextParagraph[], frame: NormalizedTextFrameStyle, box: Box, element: LayoutElement, transform: Matrix, context: MeasurementContext, suppressed: boolean, direction?: "rotate90" | "rotate270"): void {
    if (!options.textMeasurement && !options.diagnostics?.textOverflow) return;
    const contentBox = { x: box.x + frame.margin.left, y: box.y + frame.margin.top,
      width: Math.max(0, box.width - frame.margin.left - frame.margin.right), height: Math.max(0, box.height - frame.margin.top - frame.margin.bottom) };
    if (!options.textMeasurement) {
      result.diagnostics.push({ ...context, code: "text-measurement-unavailable", severity: "warning", message: "Text overflow was not checked. Supply textMeasurement.measure with known font metrics." });
      return;
    }
    const frameDirection = frame.textDirection === "rotate90" || frame.textDirection === "rotate270" ? frame.textDirection : undefined;
    direction ??= frameDirection;
    const measurementBox = direction === undefined ? contentBox : { x: 0, y: 0, width: contentBox.height, height: contentBox.width };
    const measured = measureTextBody(content, frameDirection ? { ...frame, textDirection: "horizontal" } : frame, measurementBox, themeFor(context), options.textMeasurement, context, result.diagnostics);
    if (measured === undefined) return;
    const bounds = direction === "rotate90"
      ? { x: contentBox.x + contentBox.width - measured.y - measured.height, y: contentBox.y + measured.x, width: measured.height, height: measured.width }
      : direction === "rotate270"
        ? { x: contentBox.x + measured.y, y: contentBox.y + contentBox.height - measured.x - measured.width, width: measured.height, height: measured.width }
        : measured;
    if (direction !== undefined) result.diagnostics.push({ ...context, code: frameDirection ? "text-frame-direction-measurement-approximate" : "table-text-direction-measurement-approximate", severity: "warning", message: "Rotated text bounds use horizontal font measurements and a quarter-turn; native margins, wrapping and glyph orientation can differ." });
    result.measurements.push({ ...context, contentBox, bounds });
    if (options.diagnostics?.textOverflow && overflows(bounds, contentBox, tolerance)) result.diagnostics.push({ ...context, code: "text-overflow", severity: "warning", message: "Measured text exceeds its content frame and may be clipped. Increase the frame, shorten the text, or adjust its font/spacing.", bounds: { ...bounds }, available: { ...contentBox } });
    boundary(transformedBounds({ ...bounds, x: element.box.x + bounds.x, y: element.box.y + bounds.y }, transform), context, suppressed);
  }

  function elements(items: readonly LayoutElement[], parent: Matrix, path: string, owner: { slideId?: string; layoutId?: string; masterId?: string }, ancestorAllowed = false, ancestorOpacity = 1): void {
    items.forEach((element, index) => {
      if (element.hidden && !options.diagnostics?.includeHidden) return;
      const context = { ...owner, elementId: element.id, path: `${path}.${index}` };
      const transform = multiply(parent, elementTransform(element.box, element.transform));
      const suppressed = ancestorAllowed || allowed.has(element.id);
      boundary(transformedBounds(element.box, transform), { ...context, path: `${context.path}.box` }, suppressed);
      if (options.diagnostics?.shadowBounds && !suppressed && element.type === "shape" && element.style.shadow && element.style.shadow.opacity > 0 && element.opacity > 0 && ancestorOpacity > 0) {
        const shadow = element.style.shadow, angle = shadow.direction * Math.PI / 180, spread = shadow.blur * 1.5 + element.style.stroke.width / 2;
        const local = { x: element.box.x + shadow.distance * Math.cos(angle) - spread, y: element.box.y + shadow.distance * Math.sin(angle) - spread, width: element.box.width + spread * 2, height: element.box.height + spread * 2 };
        const bounds = transformedBounds(local, transform), shadowContext = { ...context, path: `${context.path}.style.shadow` };
        if (![bounds.x, bounds.y, bounds.width, bounds.height, bounds.x + bounds.width, bounds.y + bounds.height].every(Number.isFinite)) result.diagnostics.push({ ...shadowContext, code: "geometry-unmeasurable", severity: "warning", message: "Estimated shadow bounds exceed finite numeric range." });
        else if (overflows(bounds, slideBox, tolerance)) result.diagnostics.push({ ...shadowContext, code: "shadow-slide-overflow-approximate", severity: "warning", message: "Estimated SVG shadow bounds extend outside the slide; native blur and painted bounds can differ.", bounds, available: { ...slideBox } });
      }
      if (element.type === "group") {
        const childCoordinates: Matrix = [element.box.width / element.coordinateSize.width, 0, 0, element.box.height / element.coordinateSize.height, element.box.x, element.box.y];
        elements(element.children, multiply(transform, childCoordinates), `${context.path}.children`, owner, suppressed, ancestorOpacity * element.opacity);
      } else if (element.type === "text") {
        text(element.content, element.frame, { ...element.box, x: 0, y: 0 }, element, transform, { ...context, path: `${context.path}.content` }, suppressed);
      } else if (element.type === "shape" && element.text !== undefined) {
        const box = resolveShapeTextBox({ ...element, box: { ...element.box, x: 0, y: 0 } });
        text(element.text.content, element.text.frame, box, element, transform, { ...context, path: `${context.path}.text.content` }, suppressed);
      } else if (element.type === "table" && (options.textMeasurement || options.diagnostics?.textOverflow)) {
        const width = element.columns.reduce((sum, value) => sum + value, 0);
        const height = element.rows.reduce((sum, row) => sum + row.height, 0);
        if (Math.abs(width - element.box.width) > tolerance || Math.abs(height - element.box.height) > tolerance) {
          result.diagnostics.push({ ...context, code: "table-geometry-mismatch", severity: "warning", message: "Table grid dimensions differ from its frame. Cell measurements use explicit row/column dimensions; native applications may resize the grid." });
        }
        const grid = cellGrid(element);
        for (const { rowIndex, cellIndex } of grid.invalid) result.diagnostics.push({ ...context, path: `${context.path}.rows.${rowIndex}.cells.${cellIndex}.content`, code: "table-span-invalid", severity: "warning", message: "Cannot measure a cell whose span overlaps another cell or exceeds the table grid." });
        for (const { rowIndex, cellIndex, box } of grid.cells) {
          const cell = element.rows[rowIndex]!.cells[cellIndex]!;
          const cellContext = { ...context, path: `${context.path}.rows.${rowIndex}.cells.${cellIndex}.content` };
          const direction = cell.style.textDirection === "rotate90" || cell.style.textDirection === "rotate270" ? cell.style.textDirection : undefined;
          text(cell.content, { ...(cell.style.textDirection !== undefined && direction === undefined ? { textDirection: cell.style.textDirection } : {}), margin: cell.style.margin, verticalAlign: cell.style.verticalAlign, wrap: true, autoFit: { mode: "none" } }, box, element, transform, cellContext, suppressed, direction);
        }
      } else if (element.type === "chart" && options.diagnostics?.textOverflow) {
        result.diagnostics.push({ ...context, code: "chart-text-overflow-unchecked", severity: "warning", message: "Native chart label wrapping is application-dependent; this text-overflow check covers text frames, shape text and table cells." });
      }
    });
  }

  (layout.masters ?? []).forEach((master, index) => elements(master.elements ?? [], identity, `masters.${index}.elements`, { masterId: master.id }));
  layout.layouts.forEach((item, index) => elements(item.elements, identity, `layouts.${index}.elements`, { layoutId: item.id }));
  layout.slides.forEach((slide, index) => {
    if (!slide.hidden || options.diagnostics?.includeHidden) elements(slide.elements, identity, `slides.${index}.elements`, { slideId: slide.id });
  });
  return result;
}
