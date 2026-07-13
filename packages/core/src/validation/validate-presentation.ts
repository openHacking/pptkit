import type { PresentationAsset } from "../types/asset.js";
import type {
  ConnectorEndpointInput,
  NormalizedElementAccessibility,
  PresentationElement,
  TextContentInput,
} from "../types/element.js";
import type { PresentationDiagnostic } from "../types/diagnostic.js";
import type { Box } from "../types/geometry.js";
import type { PresentationDocument, PresentationSlide, SlideLayoutDefinition } from "../types/presentation.js";
import type { PaintInput, StrokeStyleInput, TextFrameStyleInput, TextParagraphStyleInput, TextRunStyleInput } from "../types/style.js";
import { duplicates } from "./collection.js";
import { isFiniteNumber, isValidBox, isValidPoint, isValidSize } from "./geometry.js";
import { isValidColor, isValidOpacity, isValidPaint } from "./style.js";

interface Context {
  diagnostics: PresentationDiagnostic[];
  slideIds: Set<string>;
  assetIds: Set<string>;
  globalElementIds: Set<string>;
}

function add(context: Context, diagnostic: Omit<PresentationDiagnostic, "severity"> & { severity?: "error" | "warning" }): void {
  context.diagnostics.push({ severity: diagnostic.severity ?? "error", ...diagnostic });
}

function validatePaint(context: Context, paint: PaintInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (!isValidPaint(paint)) add(context, { code: "invalid-paint", message: "Paint must use a valid color and opacity between 0 and 1.", path, ...ids });
}

function validateStroke(context: Context, stroke: StrokeStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (stroke === undefined) return;
  validatePaint(context, stroke.paint, `${path}.paint`, ids);
  if (stroke.width !== undefined && (!isFiniteNumber(stroke.width) || stroke.width < 0)) {
    add(context, { code: "invalid-stroke-width", message: "Stroke width must be a non-negative finite number.", path: `${path}.width`, ...ids });
  }
}

function validateFrame(context: Context, frame: TextFrameStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (frame === undefined) return;
  if (typeof frame.margin === "number" && (!isFiniteNumber(frame.margin) || frame.margin < 0)) {
    add(context, { code: "invalid-text-margin", message: "Text margin must be non-negative.", path: `${path}.margin`, ...ids });
  }
  if (typeof frame.margin === "object") {
    for (const [side, value] of Object.entries(frame.margin)) {
      if (!isFiniteNumber(value) || value < 0) add(context, { code: "invalid-text-margin", message: "Text margins must be non-negative finite numbers.", path: `${path}.margin.${side}`, ...ids });
    }
  }
  if (frame.autoFit?.mode === "shrink" && frame.autoFit.fontScale !== undefined && !isValidOpacity(frame.autoFit.fontScale)) {
    add(context, { code: "invalid-font-scale", message: "fontScale must be between 0 and 1.", path: `${path}.autoFit.fontScale`, ...ids });
  }
}

function validateParagraphStyle(context: Context, style: TextParagraphStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (style === undefined) return;
  for (const key of ["indent", "hanging", "spaceBefore", "spaceAfter"] as const) {
    const value = style[key];
    if (value !== undefined && (!isFiniteNumber(value) || value < 0)) add(context, { code: "invalid-paragraph-style", message: `${key} must be a non-negative finite number.`, path: `${path}.${key}`, ...ids });
  }
  if (style.lineSpacing !== undefined && (!isFiniteNumber(style.lineSpacing) || style.lineSpacing <= 0)) add(context, { code: "invalid-line-spacing", message: "lineSpacing must be positive.", path: `${path}.lineSpacing`, ...ids });
  if (style.bullet?.type === "number" && style.bullet.startAt !== undefined && (!Number.isInteger(style.bullet.startAt) || style.bullet.startAt < 1)) add(context, { code: "invalid-number-start", message: "Numbered lists must start at a positive integer.", path: `${path}.bullet.startAt`, ...ids });
}

function validateRunStyle(context: Context, style: TextRunStyleInput | undefined, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (style === undefined) return;
  if (style.fontSize !== undefined && (!isFiniteNumber(style.fontSize) || style.fontSize <= 0)) add(context, { code: "invalid-font-size", message: "fontSize must be positive.", path: `${path}.fontSize`, ...ids });
  if (style.color !== undefined && !isValidColor(style.color)) add(context, { code: "invalid-color", message: "Text color must be a six-digit RGB value or theme reference.", path: `${path}.color`, ...ids });
}

function validateAction(context: Context, action: PresentationElement["action"], path: string, ids: Partial<PresentationDiagnostic>): void {
  if (action === undefined) return;
  if (action.type === "url" && action.url.trim() === "") add(context, { code: "invalid-action-url", message: "URL actions require a non-empty URL.", path, ...ids });
  if (action.type === "slide" && !context.slideIds.has(action.slideId)) add(context, { code: "missing-action-slide", message: `Action references missing slide "${action.slideId}".`, path, ...ids });
}

function validateText(context: Context, content: TextContentInput, path: string, ids: Partial<PresentationDiagnostic>): void {
  if (typeof content === "string") return;
  if (!Array.isArray(content)) {
    add(context, { code: "invalid-text-content", message: "Text content must be a string or paragraph array.", path, ...ids });
    return;
  }
  content.forEach((paragraph, paragraphIndex) => {
    validateParagraphStyle(context, paragraph.style, `${path}.${paragraphIndex}.style`, ids);
    if (!Array.isArray(paragraph.runs)) add(context, { code: "invalid-text-runs", message: "Paragraph runs must be an array.", path: `${path}.${paragraphIndex}.runs`, ...ids });
    else paragraph.runs.forEach((run, runIndex) => {
      if (typeof run.text !== "string") add(context, { code: "invalid-run-text", message: "Text run content must be a string.", path: `${path}.${paragraphIndex}.runs.${runIndex}.text`, ...ids });
      validateRunStyle(context, run.style, `${path}.${paragraphIndex}.runs.${runIndex}.style`, ids);
      validateAction(context, run.action, `${path}.${paragraphIndex}.runs.${runIndex}.action`, ids);
    });
  });
}

function collectElementIds(elements: readonly PresentationElement[], destination: Set<string>): void {
  for (const element of elements) {
    destination.add(element.id);
    if (element.type === "group") collectElementIds(element.children, destination);
  }
}

function endpointReference(endpoint: ConnectorEndpointInput): string | undefined {
  return "elementId" in endpoint ? endpoint.elementId : undefined;
}

function validateElement(
  context: Context,
  element: PresentationElement,
  path: string,
  scopeIds: Set<string>,
  placeholderBoxes: ReadonlyMap<string, Box>,
  ids: Partial<PresentationDiagnostic>,
): void {
  const identity = { ...ids, elementId: element.id };
  if (context.globalElementIds.has(element.id)) add(context, { code: "duplicate-element-id", message: `Duplicate element id "${element.id}".`, path: `${path}.id`, ...identity });
  context.globalElementIds.add(element.id);
  if (element.box === undefined) {
    if (element.type !== "connector" && (element.placeholderKey === undefined || !placeholderBoxes.has(element.placeholderKey))) add(context, { code: "missing-element-box", message: "Elements require a box or a valid placeholder binding.", path: `${path}.box`, ...identity });
  } else if (!isValidBox(element.box)) add(context, { code: "invalid-element-box", message: "Element box must contain finite coordinates and non-negative dimensions.", path: `${path}.box`, ...identity });
  if (element.placeholderKey !== undefined && !placeholderBoxes.has(element.placeholderKey)) add(context, { code: "missing-placeholder", message: `Unknown placeholder key "${element.placeholderKey}".`, path: `${path}.placeholderKey`, ...identity });
  if (element.opacity !== undefined && !isValidOpacity(element.opacity)) add(context, { code: "invalid-opacity", message: "Element opacity must be between 0 and 1.", path: `${path}.opacity`, ...identity });
  if (element.transform?.rotation !== undefined && !isFiniteNumber(element.transform.rotation)) add(context, { code: "invalid-rotation", message: "Rotation must be finite.", path: `${path}.transform.rotation`, ...identity });
  validateAction(context, element.action, `${path}.action`, identity);

  if (element.type === "text") {
    validateText(context, element.content, `${path}.content`, identity);
    validateFrame(context, element.frame, `${path}.frame`, identity);
  } else if (element.type === "image") {
    if (!context.assetIds.has(element.assetId)) add(context, { code: "missing-image-asset", message: `Image references missing asset "${element.assetId}".`, path: `${path}.assetId`, assetId: element.assetId, ...identity });
    const crop = { left: 0, top: 0, right: 0, bottom: 0, ...element.crop };
    for (const [side, value] of Object.entries(crop)) if (!isValidOpacity(value)) add(context, { code: "invalid-image-crop", message: "Crop values must be between 0 and 1.", path: `${path}.crop.${side}`, ...identity });
    if (crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1) add(context, { code: "invalid-image-crop", message: "Opposing crop values must leave a visible area.", path: `${path}.crop`, ...identity });
  } else if (element.type === "shape") {
    validatePaint(context, element.style?.fill, `${path}.style.fill`, identity);
    validateStroke(context, element.style?.stroke, `${path}.style.stroke`, identity);
  } else if (element.type === "connector") {
    for (const [key, endpoint] of [["start", element.start], ["end", element.end]] as const) {
      const reference = endpointReference(endpoint);
      if (reference !== undefined && !scopeIds.has(reference)) add(context, { code: "missing-connector-target", message: `Connector references missing element "${reference}".`, path: `${path}.${key}`, ...identity });
      if (reference === undefined && !isValidPoint(endpoint as { x: number; y: number })) add(context, { code: "invalid-connector-point", message: "Connector points must contain finite coordinates.", path: `${path}.${key}`, ...identity });
    }
    element.route?.forEach((point, index) => { if (!isValidPoint(point)) add(context, { code: "invalid-connector-point", message: "Connector route points must contain finite coordinates.", path: `${path}.route.${index}`, ...identity }); });
    validateStroke(context, element.style, `${path}.style`, identity);
  } else if (element.type === "group") {
    if (!isValidSize(element.coordinateSize) || element.coordinateSize.width <= 0 || element.coordinateSize.height <= 0) add(context, { code: "invalid-group-coordinate-size", message: "Group coordinateSize must be positive.", path: `${path}.coordinateSize`, ...identity });
    const childIds = new Set<string>();
    collectElementIds(element.children, childIds);
    element.children.forEach((child, index) => validateElement(context, child, `${path}.children.${index}`, childIds, new Map(), ids));
  } else {
    if (element.columns.length === 0 || element.columns.some((width) => !isFiniteNumber(width) || width <= 0)) add(context, { code: "invalid-table-columns", message: "Table columns must contain positive widths.", path: `${path}.columns`, ...identity });
    element.rows.forEach((row, rowIndex) => {
      if (row.height !== undefined && (!isFiniteNumber(row.height) || row.height <= 0)) add(context, { code: "invalid-table-row-height", message: "Table row height must be positive.", path: `${path}.rows.${rowIndex}.height`, ...identity });
      const span = row.cells.reduce((total, cell) => total + (cell.colSpan ?? 1), 0);
      if (span > element.columns.length) add(context, { code: "table-row-overflow", message: "Table row spans more columns than the table defines.", path: `${path}.rows.${rowIndex}`, ...identity });
      row.cells.forEach((cell, cellIndex) => {
        if (!Number.isInteger(cell.rowSpan ?? 1) || (cell.rowSpan ?? 1) < 1 || !Number.isInteger(cell.colSpan ?? 1) || (cell.colSpan ?? 1) < 1) add(context, { code: "invalid-table-span", message: "rowSpan and colSpan must be positive integers.", path: `${path}.rows.${rowIndex}.cells.${cellIndex}`, ...identity });
        validateText(context, cell.content, `${path}.rows.${rowIndex}.cells.${cellIndex}.content`, identity);
        validatePaint(context, cell.style?.fill, `${path}.rows.${rowIndex}.cells.${cellIndex}.style.fill`, identity);
        validateStroke(context, cell.style?.stroke, `${path}.rows.${rowIndex}.cells.${cellIndex}.style.stroke`, identity);
      });
    });
  }
}

function validateLayout(context: Context, layout: SlideLayoutDefinition, index: number): void {
  const identity = { layoutId: layout.id };
  const keys = layout.placeholders.map((placeholder) => placeholder.key);
  for (const key of duplicates(keys)) add(context, { code: "duplicate-placeholder-key", message: `Duplicate placeholder key "${key}".`, path: `layouts.${index}.placeholders`, ...identity });
  const boxes = new Map(layout.placeholders.map((placeholder) => [placeholder.key, placeholder.box]));
  layout.placeholders.forEach((placeholder, placeholderIndex) => {
    if (!isValidBox(placeholder.box)) add(context, { code: "invalid-placeholder-box", message: "Placeholder box is invalid.", path: `layouts.${index}.placeholders.${placeholderIndex}.box`, ...identity });
    validateFrame(context, placeholder.textStyle?.frame, `layouts.${index}.placeholders.${placeholderIndex}.textStyle.frame`, identity);
    validateParagraphStyle(context, placeholder.textStyle?.paragraph, `layouts.${index}.placeholders.${placeholderIndex}.textStyle.paragraph`, identity);
    validateRunStyle(context, placeholder.textStyle?.run, `layouts.${index}.placeholders.${placeholderIndex}.textStyle.run`, identity);
  });
  validatePaint(context, layout.background, `layouts.${index}.background`, identity);
  const scopeIds = new Set<string>();
  collectElementIds(layout.elements, scopeIds);
  layout.elements.forEach((element, elementIndex) => validateElement(context, element, `layouts.${index}.elements.${elementIndex}`, scopeIds, boxes, identity));
}

function validateAsset(context: Context, asset: PresentationAsset, index: number): void {
  if (asset.source.value.trim() === "") add(context, { code: "invalid-asset-source", message: "Asset source value cannot be empty.", path: `assets.${index}.source.value`, assetId: asset.id });
  for (const key of ["width", "height"] as const) {
    const value = asset[key];
    if (value !== undefined && (!isFiniteNumber(value) || value <= 0)) add(context, { code: "invalid-asset-dimension", message: `Asset ${key} must be positive.`, path: `assets.${index}.${key}`, assetId: asset.id });
  }
}

export function validatePresentation(document: PresentationDocument): PresentationDiagnostic[] {
  const diagnostics: PresentationDiagnostic[] = [];
  const slideIds = new Set(document.slides.map((slide) => slide.id));
  const assetIds = new Set(document.assets.map((asset) => asset.id));
  const context: Context = { diagnostics, slideIds, assetIds, globalElementIds: new Set() };

  for (const id of duplicates(document.slides.map((slide) => slide.id))) add(context, { code: "duplicate-slide-id", message: `Duplicate slide id "${id}".`, path: "slides" });
  for (const id of duplicates(document.assets.map((asset) => asset.id))) add(context, { code: "duplicate-asset-id", message: `Duplicate asset id "${id}".`, path: "assets", assetId: id });
  for (const id of duplicates(document.layouts.map((layout) => layout.id))) add(context, { code: "duplicate-layout-id", message: `Duplicate layout id "${id}".`, path: "layouts", layoutId: id });

  for (const [role, color] of Object.entries(document.theme.colors ?? {})) if (!/^#?[0-9a-f]{6}$/i.test(color)) add(context, { code: "invalid-theme-color", message: `Theme color "${role}" must be six-digit RGB.`, path: `theme.colors.${role}` });
  document.assets.forEach((asset, index) => validateAsset(context, asset, index));
  document.layouts.forEach((layout, index) => validateLayout(context, layout, index));

  const layoutMap = new Map(document.layouts.map((layout) => [layout.id, layout]));
  document.slides.forEach((slide: PresentationSlide, slideIndex) => {
    const identity = { slideId: slide.id };
    const layout = slide.layoutId === undefined ? undefined : layoutMap.get(slide.layoutId);
    if (slide.layoutId !== undefined && layout === undefined) add(context, { code: "missing-slide-layout", message: `Slide references missing layout "${slide.layoutId}".`, path: `slides.${slideIndex}.layoutId`, ...identity });
    validatePaint(context, slide.background, `slides.${slideIndex}.background`, identity);
    if (slide.notes !== undefined) validateText(context, slide.notes, `slides.${slideIndex}.notes`, identity);
    const placeholderBoxes = new Map((layout?.placeholders ?? []).map((placeholder) => [placeholder.key, placeholder.box]));
    const scopeIds = new Set<string>();
    collectElementIds(slide.elements, scopeIds);
    if (layout !== undefined) collectElementIds(layout.elements, scopeIds);
    slide.elements.forEach((element, elementIndex) => validateElement(context, element, `slides.${slideIndex}.elements.${elementIndex}`, scopeIds, placeholderBoxes, identity));
  });
  return diagnostics;
}
