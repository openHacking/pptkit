import type { ArrowType, Box, NormalizedTransform, Point } from "@pptkit/core";
import type { LayoutElement, LayoutResult, LayoutSlide } from "@pptkit/layout";
import { shapePolygonPoints } from "./shape-polygon.js";
import { dashLengths } from "./style.js";

/** A transformed authoring frame, not an outline or a pixel-accurate painted region. */
export interface SvgSelectionFrame {
  readonly elementId: string;
  readonly ancestorIds: readonly string[];
  /** Clockwise in the element's untransformed coordinate system, in slide coordinates. */
  readonly corners: readonly [Readonly<Point>, Readonly<Point>, Readonly<Point>, Readonly<Point>];
}

export interface SvgSelectionOwner { readonly type: "slide" | "layout" | "master"; readonly id: string }
export interface SvgOwnedSelectionFrame extends SvgSelectionFrame { readonly owner: SvgSelectionOwner }
export interface SvgSelectionConnector {
  readonly kind: "connector";
  readonly elementId: string;
  readonly ancestorIds: readonly string[];
  readonly owner: SvgSelectionOwner;
  readonly points: readonly Readonly<Point>[];
  readonly strokeGeometry?: SvgConnectorStrokeGeometry;
  readonly markerGeometry?: SvgConnectorMarkerGeometry;
}
export interface SvgConnectorMarkerGeometry {
  readonly points: readonly Readonly<Point>[];
  readonly transform: readonly [number, number, number, number, number, number];
  readonly width: number;
  readonly beginArrow: ArrowType;
  readonly endArrow: ArrowType;
}
export interface SvgConnectorStrokeGeometry {
  readonly points: readonly Readonly<Point>[];
  readonly transform: readonly [number, number, number, number, number, number];
  readonly width: number;
  readonly dash?: readonly number[];
}
export interface SvgRoundedShapeStrokeGeometry {
  readonly kind: "rounded-rectangle";
  readonly box: Readonly<Box>;
  readonly radius: number;
  readonly transform: readonly [number, number, number, number, number, number];
  readonly width: number;
}
export type SvgCurvedShapeStrokeGeometry = SvgRoundedShapeStrokeGeometry | (Omit<SvgRoundedShapeStrokeGeometry, "kind"> & { readonly kind: "ellipse" });
export type SvgShapeOutline = { readonly type: "ellipse" } | { readonly type: "rounded-rectangle"; readonly radiusX: number; readonly radiusY: number } | { readonly type: "polygon"; readonly points: readonly Readonly<Point>[] };
export interface SvgShapePaintSummary { readonly fillMayPaint: boolean; readonly strokeMayPaint: boolean; readonly textOrShadow: boolean }
export type SvgSelectionGeometry = (SvgOwnedSelectionFrame & { readonly kind: "frame"; readonly outline?: SvgShapeOutline; readonly shapeStrokeGeometry?: SvgConnectorStrokeGeometry | SvgCurvedShapeStrokeGeometry; readonly shapePaint?: SvgShapePaintSummary }) | SvgSelectionConnector;
/** Optional coarse selection policies; these do not perform painted-pixel tests. */
export interface SvgSelectionGeometryOptions {
  readonly excludeZeroOpacity?: boolean;
  readonly excludeUnpaintedShapes?: boolean;
  readonly excludeUnpaintedConnectors?: boolean;
  readonly includeConnectorStrokeGeometry?: boolean;
  readonly includeShapeStrokeGeometry?: boolean;
  readonly includeShapePaintSummary?: boolean;
}

function hasNoConnectorPaint(element: Extract<LayoutElement, { type: "connector" }>): boolean {
  const { paint, width, beginArrow, endArrow } = element.style;
  if (paint.type === "none") return true;
  // Markers use independent opaque paint in the SVG renderer; do not infer their
  // visibility from the line's opacity or width.
  return beginArrow === "none" && endArrow === "none" && (width === 0 || paint.opacity === 0);
}

function hasNoShapePaint(element: Extract<LayoutElement, { type: "shape" }>): boolean {
  const summary = shapePaintSummary(element);
  return !summary.fillMayPaint && !summary.strokeMayPaint && !summary.textOrShadow;
}

function shapePaintSummary(element: Extract<LayoutElement, { type: "shape" }>): SvgShapePaintSummary {
  const { fill, stroke } = element.style;
  const fillInvisible = fill.type === "none" || (fill.type === "solid" && fill.opacity === 0)
    || (fill.type === "linearGradient" && fill.stops.every(stop => stop.opacity === 0))
    || (fill.type === "pattern" && fill.foregroundOpacity === 0 && fill.backgroundOpacity === 0);
  const strokeInvisible = stroke.width === 0 || stroke.paint.type === "none" || stroke.paint.opacity === 0;
  return { fillMayPaint: !fillInvisible, strokeMayPaint: !strokeInvisible, textOrShadow: element.text !== undefined || element.style.shadow !== undefined };
}

function shapeOutline(element: Extract<LayoutElement, { type: "shape" }>): SvgShapeOutline | undefined {
  const b = element.box;
  if (element.shape === "rect") return undefined;
  if (element.shape === "ellipse") return Object.freeze({ type: "ellipse" });
  if (element.shape === "roundRect") {
    const radius = Math.min(b.width, b.height) * 0.08;
    const radiusX = radius / b.width, radiusY = radius / b.height;
    if (radiusX === 0 || radiusY === 0) return undefined;
    return Object.freeze({ type: "rounded-rectangle", radiusX, radiusY });
  }
  const points = shapePolygonPoints(b, element.shape, element.slant).map(point => Object.freeze({ x: (point.x - b.x) / b.width, y: (point.y - b.y) / b.height }));
  if (!points.every(point => Number.isFinite(point.x) && Number.isFinite(point.y))) throw new TypeError("Selection outline overflowed.");
  return Object.freeze({ type: "polygon", points: Object.freeze(points) });
}

type Matrix = readonly [number, number, number, number, number, number];
const identity: Matrix = [1, 0, 0, 1, 0, 0];
const multiply = (a: Matrix, b: Matrix): Matrix => [
  a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
  a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
  a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
];
const translate = (x: number, y: number): Matrix => [1, 0, 0, 1, x, y];
const apply = (m: Matrix, x: number, y: number): Readonly<Point> => Object.freeze({ x: m[0] * x + m[2] * y + m[4], y: m[1] * x + m[3] * y + m[5] });

function transform(box: Box, t: NormalizedTransform): Matrix {
  if (![box.x, box.y, box.width, box.height, t.rotation].every(Number.isFinite) || box.width < 0 || box.height < 0) throw new TypeError("Selection geometry requires finite, nonnegative frames and finite rotations.");
  const angle = t.rotation * Math.PI / 180, c = Math.cos(angle), s = Math.sin(angle);
  const cx = box.x + box.width / 2, cy = box.y + box.height / 2;
  return multiply(multiply(translate(cx, cy), [c * (t.flipH ? -1 : 1), s * (t.flipH ? -1 : 1), -s * (t.flipV ? -1 : 1), c * (t.flipV ? -1 : 1), 0, 0]), translate(-cx, -cy));
}

function collect(layers: readonly { elements: readonly LayoutElement[]; owner?: SvgSelectionOwner }[], includeConnectors = false, options: SvgSelectionGeometryOptions = {}): readonly (SvgSelectionFrame | SvgSelectionGeometry)[] {
  const frames: (SvgSelectionFrame | SvgSelectionGeometry)[] = [];
  let visited = 0, connectorPoints = 0;
  const visit = (elements: readonly LayoutElement[], parent: Matrix, ancestors: readonly string[], depth: number, owner?: SvgSelectionOwner): void => {
    if (depth > 64) throw new RangeError("Selection geometry exceeds maximum group depth.");
    for (const element of elements) {
      if (++visited > 100000) throw new RangeError("Selection geometry exceeds maximum element count.");
      if (element.hidden || (options.excludeZeroOpacity && element.opacity === 0) || (options.excludeUnpaintedShapes && element.type === "shape" && hasNoShapePaint(element)) || (options.excludeUnpaintedConnectors && element.type === "connector" && hasNoConnectorPaint(element))) continue;
      const b = element.box, matrix = multiply(parent, transform(b, element.transform));
      if (element.type === "group") {
        if (![element.coordinateSize.width, element.coordinateSize.height].every(v => Number.isFinite(v) && v > 0)) throw new TypeError("Selection geometry requires positive group coordinate dimensions.");
        if (b.width === 0 || b.height === 0) continue;
        const child = multiply(matrix, [b.width / element.coordinateSize.width, 0, 0, b.height / element.coordinateSize.height, b.x, b.y]);
        visit(element.children, child, [...ancestors, element.id], depth + 1, owner);
      } else if (element.type === "connector" && includeConnectors && owner) {
        connectorPoints += element.route.length + 2;
        if (connectorPoints > 1000000) throw new RangeError("Selection geometry exceeds maximum connector points.");
        const rawPoints = [element.start, ...element.route, element.end];
        const points = Object.freeze(rawPoints.map(point => apply(matrix, point.x, point.y)));
        if (!points.every(p => Number.isFinite(p.x) && Number.isFinite(p.y))) throw new TypeError("Selection geometry overflowed.");
        const style = element.style;
        const dash = dashLengths(style);
        if (options.includeConnectorStrokeGeometry && dash && !Number.isFinite(dash.reduce((sum, value) => sum + value, 0))) throw new RangeError("Connector dash pattern exceeds finite hit-test range.");
        const strokeGeometry = options.includeConnectorStrokeGeometry && style.paint.type === "solid" && style.paint.opacity > 0 && style.width > 0
          ? Object.freeze({ points: Object.freeze(rawPoints.map(point => Object.freeze({ ...point }))), transform: Object.freeze([...matrix]) as SvgConnectorStrokeGeometry["transform"], width: style.width, ...(dash ? { dash: Object.freeze([...dash]) } : {}) })
          : undefined;
        const markerGeometry = options.includeConnectorStrokeGeometry && style.paint.type !== "none" && style.width > 0 && (style.beginArrow !== "none" || style.endArrow !== "none")
          ? Object.freeze({ points: Object.freeze(rawPoints.map(point => Object.freeze({ ...point }))), transform: Object.freeze([...matrix]) as SvgConnectorMarkerGeometry["transform"], width: style.width, beginArrow: style.beginArrow, endArrow: style.endArrow })
          : undefined;
        frames.push(Object.freeze({ kind: "connector", elementId: element.id, ancestorIds: Object.freeze([...ancestors]), owner, points, ...(strokeGeometry ? { strokeGeometry } : {}), ...(markerGeometry ? { markerGeometry } : {}) }));
      } else if (element.type !== "connector" && b.width > 0 && b.height > 0) {
        const corners = Object.freeze([apply(matrix, b.x, b.y), apply(matrix, b.x + b.width, b.y), apply(matrix, b.x + b.width, b.y + b.height), apply(matrix, b.x, b.y + b.height)]) as SvgSelectionFrame["corners"];
        if (!corners.every(p => Number.isFinite(p.x) && Number.isFinite(p.y))) throw new TypeError("Selection geometry overflowed.");
        const outline = includeConnectors && element.type === "shape" ? shapeOutline(element) : undefined;
        const stroke = element.type === "shape" ? element.style.stroke : undefined;
        let shapeStrokeGeometry: SvgConnectorStrokeGeometry | SvgCurvedShapeStrokeGeometry | undefined;
        if (includeConnectors && options.includeShapeStrokeGeometry && element.type === "shape" && element.shape !== "ellipse" && element.shape !== "roundRect" && stroke && stroke.paint.type === "solid" && stroke.paint.opacity > 0 && stroke.width > 0) {
          const raw = element.shape === "rect" ? [{ x: b.x, y: b.y }, { x: b.x + b.width, y: b.y }, { x: b.x + b.width, y: b.y + b.height }, { x: b.x, y: b.y + b.height }] : shapePolygonPoints(b, element.shape, element.slant);
          const dash = dashLengths(stroke);
          if (dash && !Number.isFinite(dash.reduce((sum, value) => sum + value, 0))) throw new RangeError("Shape dash pattern exceeds finite hit-test range.");
          shapeStrokeGeometry = Object.freeze({ points: Object.freeze([...raw, raw[0]!].map(point => Object.freeze(point))), transform: Object.freeze([...matrix]) as SvgConnectorStrokeGeometry["transform"], width: stroke.width, ...(dash ? { dash: Object.freeze([...dash]) } : {}) });
        }
        if (includeConnectors && options.includeShapeStrokeGeometry && element.type === "shape" && element.shape === "roundRect" && stroke && stroke.paint.type === "solid" && stroke.paint.opacity > 0 && stroke.width > 0 && stroke.width < Math.min(b.width, b.height) && stroke.dash === "solid") {
          shapeStrokeGeometry = Object.freeze({ kind: "rounded-rectangle", box: Object.freeze({ ...b }), radius: Math.min(b.width, b.height) * .08, transform: Object.freeze([...matrix]) as SvgRoundedShapeStrokeGeometry["transform"], width: stroke.width });
        }
        if (includeConnectors && options.includeShapeStrokeGeometry && element.type === "shape" && element.shape === "ellipse" && stroke && stroke.paint.type === "solid" && stroke.paint.opacity > 0 && stroke.width > 0 && stroke.width < Math.min(b.width, b.height) && stroke.dash === "solid") {
          shapeStrokeGeometry = Object.freeze({ kind: "ellipse", box: Object.freeze({ ...b }), radius: 0, transform: Object.freeze([...matrix]) as SvgCurvedShapeStrokeGeometry["transform"], width: stroke.width });
        }
        const shapePaint = includeConnectors && options.includeShapePaintSummary && element.type === "shape" ? Object.freeze(shapePaintSummary(element)) : undefined;
        frames.push(Object.freeze({ elementId: element.id, ancestorIds: Object.freeze([...ancestors]), corners, ...(owner ? { owner } : {}), ...(includeConnectors ? { kind: "frame" as const } : {}), ...(outline ? { outline } : {}), ...(shapeStrokeGeometry ? { shapeStrokeGeometry } : {}), ...(shapePaint ? { shapePaint } : {}) }));
      }
    }
  };
  for (const layer of layers) visit(layer.elements, identity, [], 0, layer.owner);
  return Object.freeze(frames);
}

/** Return detached, ordered slide-owned non-connector frames; inherited artwork is excluded. */
export function getSvgSelectionFrames(slide: LayoutSlide): readonly SvgSelectionFrame[] {
  return collect([{ elements: slide.elements }]) as readonly SvgSelectionFrame[];
}

function slideLayers(layout: LayoutResult, slideIndex: number): { elements: readonly LayoutElement[]; owner: SvgSelectionOwner }[] {
  if (!Number.isSafeInteger(slideIndex) || slideIndex < 0 || slideIndex >= layout.slides.length) throw new RangeError("slideIndex must identify a resolved slide.");
  const slide = layout.slides[slideIndex]!;
  const slideLayout = new Map(layout.layouts.map(item => [item.id, item])).get(slide.layoutId);
  if (slideLayout?.masterId !== undefined && !slideLayout.theme) throw new TypeError(`Layout "${slideLayout.id}" has no resolved master theme.`);
  const layers: { elements: readonly LayoutElement[]; owner: SvgSelectionOwner }[] = [];
  if (slide.showInheritedShapes !== false) {
    const master = slideLayout?.masterId === undefined ? undefined : new Map((layout.masters ?? []).map(item => [item.id, item])).get(slideLayout.masterId);
    if (master && slideLayout?.showMasterShapes !== false) layers.push({ elements: master.elements ?? [], owner: Object.freeze({ type: "master", id: master.id }) });
    if (slideLayout) layers.push({ elements: slideLayout.elements, owner: Object.freeze({ type: "layout", id: slideLayout.id }) });
  }
  layers.push({ elements: slide.elements, owner: Object.freeze({ type: "slide", id: slide.id }) });
  return layers;
}

/** Include visible master and layout artwork in the same layer order as the SVG renderer. */
export function getSvgSlideSelectionFrames(layout: LayoutResult, slideIndex: number): readonly SvgOwnedSelectionFrame[] {
  return collect(slideLayers(layout, slideIndex)) as readonly SvgOwnedSelectionFrame[];
}

/** Include transformed connector polylines alongside frames in SVG paint order. */
export function getSvgSlideSelectionGeometry(layout: LayoutResult, slideIndex: number, options?: SvgSelectionGeometryOptions): readonly SvgSelectionGeometry[] {
  if (options !== undefined && (options === null || typeof options !== "object" || (options.excludeZeroOpacity !== undefined && typeof options.excludeZeroOpacity !== "boolean") || (options.excludeUnpaintedShapes !== undefined && typeof options.excludeUnpaintedShapes !== "boolean") || (options.excludeUnpaintedConnectors !== undefined && typeof options.excludeUnpaintedConnectors !== "boolean") || (options.includeConnectorStrokeGeometry !== undefined && typeof options.includeConnectorStrokeGeometry !== "boolean") || (options.includeShapeStrokeGeometry !== undefined && typeof options.includeShapeStrokeGeometry !== "boolean") || (options.includeShapePaintSummary !== undefined && typeof options.includeShapePaintSummary !== "boolean"))) throw new TypeError("Selection geometry options must be booleans.");
  return collect(slideLayers(layout, slideIndex), true, options) as readonly SvgSelectionGeometry[];
}

/** Hit connector stroke bodies in paint order; arrows and other objects are excluded. */
export function hitTestSvgConnectorStrokes(items: readonly SvgSelectionGeometry[], point: Point): SvgSelectionConnector | undefined {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("Hit-test point must be finite.");
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i]!;
    if (item.kind !== "connector" || !item.strokeGeometry) continue;
    const { transform: [a, b, c, d, e, f], points, width, dash } = item.strokeGeometry;
    const scale = Math.max(Math.abs(a), Math.abs(b), Math.abs(c), Math.abs(d));
    if (!Number.isFinite(scale) || scale === 0) throw new RangeError("Connector transform is singular or nonfinite.");
    const aa = a / scale, bb = b / scale, cc = c / scale, dd = d / scale, determinant = aa * dd - bb * cc;
    const x = (point.x - e) / scale, y = (point.y - f) / scale;
    if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(determinant) || determinant === 0) throw new RangeError("Connector transform is singular or outside finite hit-test range.");
    const localX = (x * dd - y * cc) / determinant, localY = (y * aa - x * bb) / determinant;
    if (!Number.isFinite(localX) || !Number.isFinite(localY)) throw new RangeError("Connector hit-test exceeds finite range.");
    const period = dash?.reduce((sum, value) => sum + value, 0) ?? 0;
    let phase = 0;
    for (let j = 1; j < points.length; j++) {
      const p = points[j - 1]!, q = points[j]!;
      const dx = q.x - p.x, dy = q.y - p.y, length = Math.hypot(dx, dy), px = localX - p.x, py = localY - p.y;
      if (![length, px, py].every(Number.isFinite)) throw new RangeError("Connector geometry exceeds finite hit-test range.");
      const projection = length === 0 ? 0 : px * (dx / length) + py * (dy / length);
      if (!Number.isFinite(projection)) throw new RangeError("Connector geometry exceeds finite hit-test range.");
      const along = Math.max(0, Math.min(length, projection));
      const ux = length === 0 ? 0 : dx / length, uy = length === 0 ? 0 : dy / length;
      if (!dash) {
        if (Math.hypot(px - ux * along, py - uy * along) <= width / 2) return item;
      } else {
        const shift = along % period, at = phase >= period - shift ? phase - (period - shift) : phase + shift;
        let start = 0;
        for (let k = 0; k < dash.length; k += 2) {
          const end = start + dash[k]!;
          for (const cycle of [-1, 0, 1]) {
            const left = Math.max(-along, start - at + cycle * period), right = Math.min(length - along, end - at + cycle * period);
            if (left <= right) {
              const offset = Math.max(left, Math.min(right, 0)), position = along + offset;
              if (Math.hypot(px - ux * position, py - uy * position) <= width / 2) return item;
            }
          }
          start = end + dash[k + 1]!;
        }
        const advance = length % period;
        phase = phase >= period - advance ? phase - (period - advance) : phase + advance;
      }
    }
  }
  return undefined;
}

/** Hit captured polygonal shape strokes, without testing their interiors. */
export function hitTestSvgShapeStrokes(items: readonly SvgSelectionGeometry[], point: Point): Extract<SvgSelectionGeometry, { kind: "frame" }> | undefined {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("Hit-test point must be finite.");
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i]!;
    if (item.kind !== "frame" || !item.shapeStrokeGeometry) continue;
    if ("kind" in item.shapeStrokeGeometry) {
      const { kind, box, radius, width, transform: [a, b, c, d, e, f] } = item.shapeStrokeGeometry;
      const scale = Math.max(Math.abs(a), Math.abs(b), Math.abs(c), Math.abs(d));
      if (!Number.isFinite(scale) || scale === 0) throw new RangeError("Shape transform is singular or nonfinite.");
      const aa = a / scale, bb = b / scale, cc = c / scale, dd = d / scale, det = aa * dd - bb * cc;
      const x = (point.x - e) / scale, y = (point.y - f) / scale;
      if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(det) || det === 0) throw new RangeError("Shape transform is singular or outside finite hit-test range.");
      const lx = (x * dd - y * cc) / det, ly = (y * aa - x * bb) / det;
      if (kind === "ellipse") {
        const distance = ellipseDistance(lx - (box.x + box.width / 2), ly - (box.y + box.height / 2), box.width / 2, box.height / 2);
        if (distance <= width / 2) return item;
        continue;
      }
      const qx = Math.abs(lx - (box.x + box.width / 2)) - (box.width / 2 - radius);
      const qy = Math.abs(ly - (box.y + box.height / 2)) - (box.height / 2 - radius);
      const distance = Math.hypot(Math.max(qx, 0), Math.max(qy, 0)) + Math.min(Math.max(qx, qy), 0) - radius;
      if (!Number.isFinite(distance)) throw new RangeError("Shape hit-test exceeds finite range.");
      if (Math.abs(distance) <= width / 2) return item;
      continue;
    }
    const connector: SvgSelectionConnector = { kind: "connector", elementId: item.elementId, ancestorIds: item.ancestorIds, owner: item.owner, points: [], strokeGeometry: item.shapeStrokeGeometry };
    if (hitTestSvgConnectorStrokes([connector], point)) return item;
  }
  return undefined;
}

/** Euclidean distance to an ellipse via its closest-point Lagrange multiplier. */
function ellipseDistance(x: number, y: number, rx: number, ry: number): number {
  if (![x, y, rx, ry].every(Number.isFinite) || rx <= 0 || ry <= 0) throw new RangeError("Ellipse geometry exceeds finite hit-test range.");
  const major = Math.max(rx, ry), minor = Math.min(rx, ry) / major;
  const u = Math.abs(rx >= ry ? x : y) / major, v = Math.abs(rx >= ry ? y : x) / major;
  const b2 = minor * minor;
  if (![u, v, b2].every(Number.isFinite) || b2 === 0) throw new RangeError("Ellipse geometry exceeds finite hit-test range.");
  let px: number, py: number;
  // The inner multiplier approaches -b² at the major axis and cannot be
  // represented for tiny nonzero v. The axis limit changes distance by at
  // most v, avoiding unstable division while keeping a bounded error.
  if (v <= minor * 1e-12) {
    const difference = 1 - b2;
    if (difference > 0 && u < difference) { px = u / difference; py = minor * Math.sqrt(Math.max(0, 1 - px * px)); }
    else { px = 1; py = 0; }
  } else if (u === 0) { px = 0; py = minor; }
  else {
    const value = (t: number): number => Math.hypot(u / (t + 1), minor * v / (t + b2));
    const outside = Math.hypot(u, v / minor) >= 1;
    let low = outside ? 0 : -b2, high = outside ? Math.max(1, u, v) : 0;
    while (outside && value(high) > 1) { high *= 2; if (!Number.isFinite(high)) throw new RangeError("Ellipse geometry exceeds finite hit-test range."); }
    for (let i = 0; i < 80; i++) {
      const mid = low + (high - low) / 2;
      if (mid === low || mid === high) break;
      if (value(mid) > 1) low = mid; else high = mid;
    }
    const t = high;
    px = u / (t + 1); py = b2 * v / (t + b2);
  }
  const result = Math.hypot(u - px, v - py) * major;
  if (!Number.isFinite(result)) throw new RangeError("Ellipse geometry exceeds finite hit-test range.");
  return result;
}

/** Hit supported SVG connector markers, including zero-length paths. */
export function hitTestSvgConnectorMarkers(items: readonly SvgSelectionGeometry[], point: Point): SvgSelectionConnector | undefined {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("Hit-test point must be finite.");
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i]!;
    if (item.kind !== "connector" || !item.markerGeometry) continue;
    const { transform: [a, b, c, d, e, f], points, width, beginArrow, endArrow } = item.markerGeometry;
    const scale = Math.max(Math.abs(a), Math.abs(b), Math.abs(c), Math.abs(d));
    if (!Number.isFinite(scale) || scale === 0) throw new RangeError("Connector transform is singular or nonfinite.");
    const aa = a / scale, bb = b / scale, cc = c / scale, dd = d / scale, det = aa * dd - bb * cc;
    const x = (point.x - e) / scale, y = (point.y - f) / scale;
    if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(det) || det === 0) throw new RangeError("Connector transform is singular or outside finite hit-test range.");
    const localX = (x * dd - y * cc) / det, localY = (y * aa - x * bb) / det;
    if (!Number.isFinite(localX) || !Number.isFinite(localY)) throw new RangeError("Connector hit-test exceeds finite range.");
    for (const [type, start] of [[beginArrow, true], [endArrow, false]] as const) {
      if (type === "none") continue;
      const endpoint = points[start ? 0 : points.length - 1]!;
      let neighbor: Readonly<Point> | undefined;
      for (let j = start ? 1 : points.length - 2; start ? j < points.length : j >= 0; j += start ? 1 : -1) {
        const candidate = points[j]!;
        if (candidate.x !== endpoint.x || candidate.y !== endpoint.y) { neighbor = candidate; break; }
      }
      // SVG uses a zero-degree tangent for a wholly degenerate path; the
      // start marker is then reversed by auto-start-reverse.
      const dx = neighbor ? endpoint.x - neighbor.x : start ? -1 : 1;
      const dy = neighbor ? endpoint.y - neighbor.y : 0, length = Math.hypot(dx, dy);
      if (!Number.isFinite(length) || length === 0) throw new RangeError("Connector marker tangent exceeds finite range.");
      const px = localX - endpoint.x, py = localY - endpoint.y;
      const mx = (px * (dx / length) + py * (dy / length)) / width + (start ? 1 : 9);
      const my = (-px * (dy / length) + py * (dx / length)) / width + 5;
      if (!Number.isFinite(mx) || !Number.isFinite(my)) throw new RangeError("Connector marker hit-test exceeds finite range.");
      if (mx < 0 || mx > 10 || my < 0 || my > 10) continue;
      if (type === "oval") { if ((mx - 5) ** 2 + (my - 5) ** 2 <= 3.5 ** 2) return item; continue; }
      // The open arrow has a 1.5-unit butt-capped, miter-joined stroke.
      const n = .75 / Math.sqrt(5), miter = .75 * Math.sqrt(5);
      const polygon = type === "arrow" ? [[1 - n, 1 + 2 * n], [9 - miter, 5], [1 - n, 9 - 2 * n], [1 + n, 9 + 2 * n], [9 + miter, 5], [1 + n, 1 - 2 * n]]
        : type === "diamond" ? [[1, 5], [5, 1], [9, 5], [5, 9]] : type === "stealth" ? [[1, 1], [9, 5], [1, 9], [4, 5]] : [[1, 1], [9, 5], [1, 9]];
      let inside = false;
      for (let j = 0, k = polygon.length - 1; j < polygon.length; k = j++) {
        const [ax, ay] = polygon[j]!, [bx, by] = polygon[k]!;
        const cross = (mx - ax!) * (by! - ay!) - (my - ay!) * (bx! - ax!);
        if (Math.abs(cross) <= 1e-12 && mx >= Math.min(ax!, bx!) - 1e-12 && mx <= Math.max(ax!, bx!) + 1e-12 && my >= Math.min(ay!, by!) - 1e-12 && my <= Math.max(ay!, by!) + 1e-12) return item;
        if ((ay! > my) !== (by! > my) && mx < (bx! - ax!) * (my - ay!) / (by! - ay!) + ax!) inside = !inside;
      }
      if (inside) return item;
    }
  }
  return undefined;
}

/** Topmost frame containing a slide-space point; transparent areas inside frames still count. */
export function hitTestSvgSelectionFrames<T extends SvgSelectionFrame>(frames: readonly T[], point: Point): T | undefined {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("Hit-test point must be finite.");
  for (let i = frames.length - 1; i >= 0; i--) {
    const frame = frames[i]!, corners = frame.corners;
    let positive = false, negative = false;
    for (let j = 0; j < 4; j++) {
      const a = corners[j]!, b = corners[(j + 1) % 4]!;
      const dx = b.x - a.x, dy = b.y - a.y, px = point.x - a.x, py = point.y - a.y;
      if (![dx, dy, px, py].every(Number.isFinite)) throw new RangeError("Frame geometry exceeds finite hit-test range.");
      // Scale the two vectors independently: their cross-product sign is unchanged,
      // while multiplication of otherwise finite coordinates cannot overflow.
      const edgeScale = Math.max(Math.abs(dx), Math.abs(dy));
      const pointScale = Math.max(Math.abs(px), Math.abs(py));
      const first = edgeScale === 0 || pointScale === 0 ? 0 : (dx / edgeScale) * (py / pointScale);
      const second = edgeScale === 0 || pointScale === 0 ? 0 : (dy / edgeScale) * (px / pointScale);
      const cross = first - second, tolerance = 16 * Number.EPSILON * (Math.abs(first) + Math.abs(second));
      if (cross > tolerance) positive = true; else if (cross < -tolerance) negative = true;
    }
    if (!(positive && negative) && (positive || negative)) return frame;
  }
  return undefined;
}

/** Hit the topmost frame or connector centerline within a slide-space tolerance. */
export function hitTestSvgSelectionGeometry(items: readonly SvgSelectionGeometry[], point: Point, connectorTolerance: number): SvgSelectionGeometry | undefined {
  return hitGeometry(items, point, connectorTolerance, false);
}

/** Use shape outlines for shapes, retaining frame and connector-centerline behavior for other objects. */
export function hitTestSvgSelectionOutlines(items: readonly SvgSelectionGeometry[], point: Point, connectorTolerance: number): SvgSelectionGeometry | undefined {
  return hitGeometry(items, point, connectorTolerance, true);
}

/** Hit shape interiors and captured connector strokes together in paint order. */
export function hitTestSvgSelectionOutlinesAndStrokes(items: readonly SvgSelectionGeometry[], point: Point): SvgSelectionGeometry | undefined {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("Hit-test point must be finite.");
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i]!;
    if (item.kind === "connector") {
      if (hitTestSvgConnectorStrokes([item], point) || hitTestSvgConnectorMarkers([item], point)) return item;
    } else if ((hitTestSvgSelectionFrames([item], point) && (!item.outline || insideOutline(item, point))) || hitTestSvgShapeStrokes([item], point)) return item;
  }
  return undefined;
}

/** Select using coarse shape paint summaries, with frame fallbacks for unknown paint. */
export function hitTestSvgSelectionPaintAware(items: readonly SvgSelectionGeometry[], point: Point): SvgSelectionGeometry | undefined {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("Hit-test point must be finite.");
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i]!;
    if (item.kind === "connector") {
      if (hitTestSvgConnectorStrokes([item], point) || hitTestSvgConnectorMarkers([item], point)) return item;
      continue;
    }
    if (hitTestSvgShapeStrokes([item], point)) return item;
    const paint = item.shapePaint;
    if (paint && !paint.fillMayPaint && !paint.textOrShadow && (!paint.strokeMayPaint || item.shapeStrokeGeometry)) continue;
    if (!hitTestSvgSelectionFrames([item], point)) continue;
    if (!paint || paint.textOrShadow || (paint.strokeMayPaint && !item.shapeStrokeGeometry) || !item.outline || insideOutline(item, point)) return item;
  }
  return undefined;
}

function hitGeometry(items: readonly SvgSelectionGeometry[], point: Point, connectorTolerance: number, useOutlines: boolean): SvgSelectionGeometry | undefined {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) throw new TypeError("Hit-test point must be finite.");
  if (!Number.isFinite(connectorTolerance) || connectorTolerance < 0) throw new RangeError("Connector tolerance must be finite and nonnegative.");
  for (let i = items.length - 1; i >= 0; i--) {
    const item = items[i]!;
    if (item.kind === "frame") {
      if (hitTestSvgSelectionFrames([item], point) && (!useOutlines || !item.outline || insideOutline(item, point))) return item;
    } else {
      for (let j = 1; j < item.points.length; j++) {
        const a = item.points[j - 1]!, b = item.points[j]!;
        const dx = b.x - a.x, dy = b.y - a.y, length = Math.hypot(dx, dy);
        const px = point.x - a.x, py = point.y - a.y;
        if (![length, px, py].every(Number.isFinite)) throw new RangeError("Connector geometry exceeds finite hit-test range.");
        const projection = length === 0 ? 0 : px * (dx / length) + py * (dy / length);
        if (!Number.isFinite(projection)) throw new RangeError("Connector geometry exceeds finite hit-test range.");
        const along = Math.max(0, Math.min(length, projection));
        if (Math.hypot(point.x - (a.x + (length === 0 ? 0 : dx / length) * along), point.y - (a.y + (length === 0 ? 0 : dy / length) * along)) <= connectorTolerance) return item;
      }
    }
  }
  return undefined;
}

function insideOutline(frame: SvgOwnedSelectionFrame & { readonly outline?: SvgShapeOutline }, point: Point): boolean {
  const [a, b, , d] = frame.corners, ux = b.x - a.x, uy = b.y - a.y, vx = d.x - a.x, vy = d.y - a.y;
  const scale = Math.max(Math.abs(ux), Math.abs(uy), Math.abs(vx), Math.abs(vy));
  if (!Number.isFinite(scale) || scale === 0) return false;
  const u0 = ux / scale, u1 = uy / scale, v0 = vx / scale, v1 = vy / scale;
  const px = (point.x - a.x) / scale, py = (point.y - a.y) / scale, det = u0 * v1 - u1 * v0;
  if (!Number.isFinite(px) || !Number.isFinite(py) || !Number.isFinite(det) || det === 0) return false;
  const x = (px * v1 - py * v0) / det, y = (u0 * py - u1 * px) / det;
  const outline = frame.outline!;
  if (outline.type === "ellipse") return ((x - 0.5) * 2) ** 2 + ((y - 0.5) * 2) ** 2 <= 1 + 1e-12;
  if (outline.type === "rounded-rectangle") {
    const dx = Math.max(outline.radiusX - x, 0, x - (1 - outline.radiusX)) / outline.radiusX;
    const dy = Math.max(outline.radiusY - y, 0, y - (1 - outline.radiusY)) / outline.radiusY;
    return dx * dx + dy * dy <= 1 + 1e-12;
  }
  let inside = false;
  for (let i = 0, j = outline.points.length - 1; i < outline.points.length; j = i++) {
    const p = outline.points[i]!, q = outline.points[j]!;
    const cross = (x - p.x) * (q.y - p.y) - (y - p.y) * (q.x - p.x);
    if (Math.abs(cross) <= 1e-12 && x >= Math.min(p.x, q.x) - 1e-12 && x <= Math.max(p.x, q.x) + 1e-12 && y >= Math.min(p.y, q.y) - 1e-12 && y <= Math.max(p.y, q.y) + 1e-12) return true;
    if ((p.y > y) !== (q.y > y) && x < (q.x - p.x) * (y - p.y) / (q.y - p.y) + p.x) inside = !inside;
  }
  return inside;
}
