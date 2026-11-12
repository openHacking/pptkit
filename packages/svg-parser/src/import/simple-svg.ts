import { SaxesParser } from "saxes";
import type { ConnectorElementInput, ShapeElementInput } from "@pptkit/core";
import { linearPath } from "./linear-path.js";
import { namedColors } from "./named-colors.js";

export interface SvgImportIssue { readonly code: "unsupported-svg" | "flattened-groups" | "path-converted"; readonly message: string }
export type SvgImportResult =
  | { readonly status: "editable"; readonly width: number; readonly height: number; readonly elements: readonly (ShapeElementInput | ConnectorElementInput)[]; readonly issues: readonly SvgImportIssue[] }
  | { readonly status: "unsupported"; readonly issues: readonly SvgImportIssue[] };

export class SvgImportError extends Error {
  constructor(readonly code: "invalid-xml" | "input-limit" | "invalid-input", message: string) { super(message); this.name = "SvgImportError"; }
}

const SVG = "http://www.w3.org/2000/svg", XMLNS = "http://www.w3.org/2000/xmlns/", SCALE = 72 / 96;
const numeric = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?(?:px)?$/;
const svgNumber = '[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?';
const opacityPattern = new RegExp(`^${svgNumber}%?$`);
const commaRgbPattern = new RegExp(`^rgba?\\([ \\t\\r\\n]*(\\d{1,3})[ \\t\\r\\n]*,[ \\t\\r\\n]*(\\d{1,3})[ \\t\\r\\n]*,[ \\t\\r\\n]*(\\d{1,3})(?:[ \\t\\r\\n]*,[ \\t\\r\\n]*(${svgNumber}%?))?[ \\t\\r\\n]*\\)$`);
const spaceRgbPattern = new RegExp(`^rgba?\\([ \\t\\r\\n]*(\\d{1,3})[ \\t\\r\\n]+(\\d{1,3})[ \\t\\r\\n]+(\\d{1,3})(?:[ \\t\\r\\n]*/[ \\t\\r\\n]*(${svgNumber}%?))?[ \\t\\r\\n]*\\)$`);
const viewBoxPattern = new RegExp(`^[ \\t\\r\\n]*(${svgNumber})(?:[ \\t\\r\\n]*,[ \\t\\r\\n]*|[ \\t\\r\\n]+)(${svgNumber})(?:[ \\t\\r\\n]*,[ \\t\\r\\n]*|[ \\t\\r\\n]+)(${svgNumber})(?:[ \\t\\r\\n]*,[ \\t\\r\\n]*|[ \\t\\r\\n]+)(${svgNumber})[ \\t\\r\\n]*$`);
const transformPattern = new RegExp(`^(translate|scale)[ \\t\\r\\n]*\\([ \\t\\r\\n]*(${svgNumber})(?:(?:[ \\t\\r\\n]*,[ \\t\\r\\n]*|[ \\t\\r\\n]+)(${svgNumber}))?[ \\t\\r\\n]*\\)`);
function scalar(value: string | undefined): number | undefined {
  if (value === undefined || !numeric.test(value)) return undefined;
  const result = Number(value.endsWith("px") ? value.slice(0, -2) : value);
  return Number.isFinite(result) && Math.abs(result) <= 1_000_000 ? result : undefined;
}

type Matrix = { sx: number; sy: number; tx: number; ty: number };
function transform(value: string | undefined): Matrix | undefined {
  const result = { sx: 1, sy: 1, tx: 0, ty: 0 };
  if (value === undefined) return result;
  if (value.length > 4096) return undefined;
  let remaining = value.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, ""), count = 0;
  while (remaining) {
    if (++count > 64) return undefined;
    const match = transformPattern.exec(remaining);
    if (!match) return undefined;
    const first = scalar(match[2]), second = match[3] === undefined ? undefined : scalar(match[3]);
    if (first === undefined || (match[3] !== undefined && second === undefined)) return undefined;
    if (match[1] === "translate") { result.tx += result.sx * first; result.ty += result.sy * (second ?? 0); }
    else { result.sx *= first; result.sy *= second ?? first; }
    if (![result.sx, result.sy, result.tx, result.ty].every(Number.isFinite) || result.sx === 0 || result.sy === 0) return undefined;
    remaining = remaining.slice(match[0].length);
    if (remaining) {
      const separator = /^(?:[ \t\r\n]*,[ \t\r\n]*|[ \t\r\n]+)/.exec(remaining);
      if (!separator) return undefined;
      remaining = remaining.slice(separator[0].length);
      if (!remaining) return undefined;
    }
  }
  return count ? result : undefined;
}

function compose(parent: Matrix, local: Matrix): Matrix | undefined {
  const result = { sx: parent.sx * local.sx, sy: parent.sy * local.sy, tx: parent.tx + parent.sx * local.tx, ty: parent.ty + parent.sy * local.ty };
  return Object.values(result).every(Number.isFinite) && result.sx !== 0 && result.sy !== 0 ? result : undefined;
}

type LiteralColor = string | { color: string; opacity: number };
function solid(color: LiteralColor, opacity: number): { type: "solid"; color: string; opacity?: number } {
  const combined = opacity * (typeof color === "string" ? 1 : color.opacity);
  return { type: "solid", color: typeof color === "string" ? color : color.color, ...(combined === 1 ? {} : { opacity: combined }) };
}

function paintOpacity(value: string | undefined, inherited: number): number | undefined {
  if (value === undefined) return inherited;
  const token = value.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, "").toLowerCase();
  if (token === "inherit") return inherited;
  if (!opacityPattern.test(token)) return undefined;
  const percent = token.endsWith("%"), result = Number(percent ? token.slice(0, -1) : token) / (percent ? 100 : 1);
  return Number.isFinite(result) && result >= 0 && result <= 1 ? result : undefined;
}

function fill(value: string | undefined, inherited: LiteralColor | null): LiteralColor | null | undefined {
  if (value === undefined) return inherited;
  const token = value.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, "").toLowerCase();
  if (token === "inherit") return inherited;
  if (token === "none") return null;
  if (token === "transparent") return { color: "000000", opacity: 0 };
  if (/^#[0-9a-f]{8}$/.test(token)) return { color: token.slice(1, 7).toUpperCase(), opacity: parseInt(token.slice(7, 9), 16) / 255 };
  if (/^#[0-9a-f]{4}$/.test(token)) return { color: [...token.slice(1, 4)].map(c => c + c).join("").toUpperCase(), opacity: parseInt(token.slice(4).repeat(2), 16) / 255 };
  if (/^#[0-9a-f]{6}$/.test(token)) return token.slice(1).toUpperCase();
  if (/^#[0-9a-f]{3}$/.test(token)) return [...token.slice(1)].map(c => c + c).join("").toUpperCase();
  if (Object.hasOwn(namedColors, token)) return namedColors[token];
  const rgb = commaRgbPattern.exec(token) ?? spaceRgbPattern.exec(token);
  if (rgb) {
    const channels = rgb.slice(1, 4).map(Number), alpha = rgb[4] === undefined ? 1 : paintOpacity(rgb[4], 1);
    if (channels.every(channel => channel <= 255) && alpha !== undefined) {
      const color = channels.map(channel => channel.toString(16).padStart(2, "0")).join("").toUpperCase();
      return alpha === 1 ? color : { color, opacity: alpha };
    }
  }
  return undefined;
}

function resolveStrokeWidth(value: string | undefined, inherited: number): number | undefined {
  if (value === undefined) return inherited;
  const token = value.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, "").toLowerCase();
  return token === "inherit" ? inherited : scalar(token);
}

type LineCap = "butt" | "round" | "square";
type LineJoin = "miter" | "round" | "bevel";
function lineStyle<T extends string>(value: string | undefined, inherited: T, allowed: readonly T[]): T | undefined {
  if (value === undefined) return inherited;
  const token = value.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, "").toLowerCase();
  if (token === "inherit") return inherited;
  return allowed.find(item => item === token);
}

/** Experimental all-or-nothing import: unsupported sources return no objects. */
export function importSimpleSvg(input: string): SvgImportResult {
  if (typeof input !== "string") throw new SvgImportError("invalid-input", "SVG input must be a string.");
  if (input.length > 1_000_000) throw new SvgImportError("input-limit", "SVG input exceeds the character limit.");
  const parser = new SaxesParser({ xmlns: true }), issues: SvgImportIssue[] = [], warnings: SvgImportIssue[] = [], elements: (ShapeElementInput | ConnectorElementInput)[] = [];
  const contexts: { kind: "group" | "shape" | "invalid"; matrix: Matrix; fill: LiteralColor | null; stroke: LiteralColor | null; fillOpacity: number; strokeOpacity: number; strokeWidth: number; cap: LineCap; join: LineJoin }[] = [];
  let depth = 0, count = 0, width: number | undefined, height: number | undefined, root = false;
  let viewport: { x: number; y: number; width: number; height: number; sx: number; sy: number; dx: number; dy: number } | undefined;
  const unsupported = (message: string): void => { if (issues.length < 100) issues.push(Object.freeze({ code: "unsupported-svg", message })); };
  parser.on("doctype", () => { throw new SvgImportError("invalid-xml", "DTD declarations are not supported."); });
  parser.on("processinginstruction", () => unsupported("Processing instructions are unsupported."));
  parser.on("text", text => { if (text.trim()) unsupported("Text content is unsupported."); });
  parser.on("cdata", () => unsupported("CDATA is unsupported."));
  parser.on("opentag", tag => {
    depth++; count++;
    if (depth > 64 || count > 10000 || Object.keys(tag.attributes).length > 64) throw new SvgImportError("input-limit", "SVG element, depth or attribute limit exceeded.");
    if (tag.uri !== SVG || tag.prefix) unsupported("Elements must use the default SVG namespace.");
    const values: Record<string, string> = Object.create(null) as Record<string, string>;
    for (const attr of Object.values(tag.attributes)) {
      if (attr.uri === XMLNS && attr.name === "xmlns" && depth === 1 && attr.value === SVG) continue;
      if (attr.uri !== "" || attr.prefix) { unsupported("Namespaced attributes are unsupported."); continue; }
      values[attr.local] = attr.value;
    }
    if (depth === 1) {
      root = tag.local === "svg" && tag.uri === SVG && !tag.prefix;
      if (!root) unsupported("Root must be an SVG element.");
      if (Object.keys(values).some(key => !["width", "height", "viewBox", "preserveAspectRatio", "fill", "stroke", "fill-opacity", "stroke-opacity", "stroke-width", "stroke-linecap", "stroke-linejoin"].includes(key))) unsupported("Unsupported root attribute.");
      const rootFill = fill(values.fill, "000000");
      const rootStroke = fill(values.stroke, null), rootStrokeWidth = resolveStrokeWidth(values["stroke-width"], 1);
      const cap = lineStyle<LineCap>(values["stroke-linecap"], "butt", ["butt", "round", "square"]), join = lineStyle<LineJoin>(values["stroke-linejoin"], "miter", ["miter", "round", "bevel"]);
      const fillOpacity = paintOpacity(values["fill-opacity"], 1), strokeOpacity = paintOpacity(values["stroke-opacity"], 1);
      if (rootFill === undefined || rootStroke === undefined || rootStrokeWidth === undefined || rootStrokeWidth < 0 || cap === undefined || join === undefined || fillOpacity === undefined || strokeOpacity === undefined) unsupported("Unsupported root paint or stroke style.");
      contexts.push({ kind: "group", matrix: { sx: 1, sy: 1, tx: 0, ty: 0 }, fill: rootFill === undefined ? "000000" : rootFill, stroke: rootStroke ?? null, fillOpacity: fillOpacity ?? 1, strokeOpacity: strokeOpacity ?? 1, strokeWidth: rootStrokeWidth ?? 1, cap: cap ?? "butt", join: join ?? "miter" });
      width = scalar(values.width); height = scalar(values.height);
      if (width === undefined || height === undefined || width <= 0 || height <= 0) unsupported("Root requires positive explicit pixel width and height.");
      const aspect = values.preserveAspectRatio?.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, "").replace(/[ \t\r\n]+/g, " ") ?? "xMidYMid meet";
      const alignment = /^(?:x(Min|Mid|Max)Y(Min|Mid|Max))(?: (meet|slice))?$/.exec(aspect);
      if (aspect !== "none" && !alignment) unsupported("Unsupported preserveAspectRatio.");
      if (width !== undefined && height !== undefined && width > 0 && height > 0) {
        let x = 0, y = 0, vw = width, vh = height;
        if (values.viewBox !== undefined) {
          const match = viewBoxPattern.exec(values.viewBox);
          const numbers = match ? match.slice(1).map(value => scalar(value)) : [];
          if (numbers.length !== 4 || numbers.some(value => value === undefined) || numbers[2]! <= 0 || numbers[3]! <= 0) unsupported("Invalid or unsupported viewBox.");
          else [x, y, vw, vh] = numbers as [number, number, number, number];
        }
        let sx = width / vw, sy = height / vh;
        if (aspect !== "none") sx = sy = alignment?.[3] === "slice" ? Math.max(sx, sy) : Math.min(sx, sy);
        const fraction = (value: string | undefined): number => value === "Min" ? 0 : value === "Max" ? 1 : 0.5;
        const dx = (width - vw * sx) * fraction(alignment?.[1]), dy = (height - vh * sy) * fraction(alignment?.[2]);
        if (![sx, sy, dx, dy].every(Number.isFinite)) unsupported("ViewBox scale is not finite.");
        else viewport = { x, y, width: vw, height: vh, sx, sy, dx, dy };
      }
      return;
    }
    const parent = contexts[depth - 2];
    if (!parent || parent.kind !== "group") { contexts.push({ kind: "invalid", matrix: { sx: 1, sy: 1, tx: 0, ty: 0 }, fill: null, stroke: null, fillOpacity: 1, strokeOpacity: 1, strokeWidth: 1, cap: "butt", join: "miter" }); unsupported("Unsupported nesting."); return; }
    if (tag.local === "g") {
      if (Object.keys(values).some(key => !["fill", "stroke", "fill-opacity", "stroke-opacity", "stroke-width", "stroke-linecap", "stroke-linejoin", "transform"].includes(key))) unsupported("Unsupported group attribute.");
      const local = transform(values.transform), matrix = local && compose(parent.matrix, local), color = fill(values.fill, parent.fill);
      const stroke = fill(values.stroke, parent.stroke), strokeWidth = resolveStrokeWidth(values["stroke-width"], parent.strokeWidth);
      const cap = lineStyle(values["stroke-linecap"], parent.cap, ["butt", "round", "square"]), join = lineStyle(values["stroke-linejoin"], parent.join, ["miter", "round", "bevel"]);
      const fillOpacity = paintOpacity(values["fill-opacity"], parent.fillOpacity), strokeOpacity = paintOpacity(values["stroke-opacity"], parent.strokeOpacity);
      if (!matrix || color === undefined || stroke === undefined || strokeWidth === undefined || strokeWidth < 0 || cap === undefined || join === undefined || fillOpacity === undefined || strokeOpacity === undefined) { contexts.push({ ...parent, kind: "invalid" }); unsupported("Unsupported group transform, paint or stroke style."); return; }
      contexts.push({ kind: "group", matrix, fill: color, stroke, fillOpacity, strokeOpacity, strokeWidth, cap, join });
      if (!warnings.some(warning => warning.code === "flattened-groups")) warnings.push(Object.freeze({ code: "flattened-groups", message: "SVG group structure is flattened; resulting objects are individually editable." }));
      return;
    }
    contexts.push({ ...parent, kind: "shape" });
    if (["path", "polyline", "polygon"].includes(tag.local)) {
      const geometryKey = tag.local === "path" ? "d" : "points";
      if (Object.keys(values).some(key => ![geometryKey, "fill", "stroke", "fill-opacity", "stroke-opacity", "stroke-width", "stroke-linecap", "stroke-linejoin", "transform"].includes(key))) { unsupported("Unsupported linear geometry attribute."); return; }
      const raw = values[geometryKey];
      const data = tag.local === "path" ? raw : raw && /^[\dEe+.,\- \t\r\n]+$/.test(raw) ? `M ${raw}${tag.local === "polygon" ? " Z" : ""}` : undefined;
      const parsedPath = linearPath(data), color = fill(values.fill, parent.fill), stroke = fill(values.stroke, parent.stroke);
      const strokeWidth = resolveStrokeWidth(values["stroke-width"], parent.strokeWidth);
      const local = transform(values.transform), matrix = local && compose(parent.matrix, local);
      const cap = lineStyle(values["stroke-linecap"], parent.cap, ["butt", "round", "square"]), join = lineStyle(values["stroke-linejoin"], parent.join, ["miter", "round", "bevel"]);
      const fillOpacity = paintOpacity(values["fill-opacity"], parent.fillOpacity), strokeOpacity = paintOpacity(values["stroke-opacity"], parent.strokeOpacity);
      if (!parsedPath || color !== null || !stroke || strokeWidth === undefined || strokeWidth <= 0 || !matrix || !viewport || width === undefined || height === undefined || !cap || join !== "round" || (!parsedPath.closed && cap !== "round") || fillOpacity === undefined || strokeOpacity === undefined) { unsupported("Linear geometry requires no fill, a positive solid stroke, round joins and round caps for open paths."); return; }
      const sx = Math.abs(matrix.sx) * viewport.sx, sy = Math.abs(matrix.sy) * viewport.sy;
      if (sx !== sy || !Number.isFinite(sx * strokeWidth)) { unsupported("Nonuniform or invalid path stroke scaling is unsupported."); return; }
      const hx = strokeWidth * Math.abs(matrix.sx) / 2, hy = strokeWidth * Math.abs(matrix.sy) / 2, hm = strokeWidth * sx / 2;
      const mapped: { x: number; y: number }[] = [];
      for (const point of parsedPath.points) {
        const x = point.x * matrix.sx + matrix.tx, y = point.y * matrix.sy + matrix.ty;
        const mx = (x - viewport.x) * viewport.sx + viewport.dx, my = (y - viewport.y) * viewport.sy + viewport.dy;
        if (![x, y, mx, my].every(Number.isFinite) || x - hx < viewport.x || x + hx > viewport.x + viewport.width || y - hy < viewport.y || y + hy > viewport.y + viewport.height || mx - hm < 0 || mx + hm > width || my - hm < 0 || my + hm > height) { unsupported("Path stroke crosses the viewBox or viewport."); return; }
        mapped.push({ x: mx * SCALE, y: my * SCALE });
      }
      elements.push({ type: "connector", start: mapped[0]!, end: mapped[mapped.length - 1]!, route: mapped.slice(1, -1), pathMode: parsedPath.closed ? "closed-round" : "open-round", style: { paint: solid(stroke, strokeOpacity), width: strokeWidth * sx * SCALE } });
      if (!warnings.some(warning => warning.code === "path-converted")) warnings.push(Object.freeze({ code: "path-converted", message: "Linear SVG geometry is converted to an editable connector; its original element and commands are not retained." }));
      return;
    }
    if (!["rect", "circle", "ellipse"].includes(tag.local)) { unsupported("Unsupported element."); return; }
    const allowed = tag.local === "rect" ? ["x", "y", "width", "height", "fill", "transform", "stroke", "fill-opacity", "stroke-opacity", "stroke-width"] : tag.local === "circle" ? ["cx", "cy", "r", "fill", "transform", "stroke", "fill-opacity", "stroke-opacity", "stroke-width"] : ["cx", "cy", "rx", "ry", "fill", "transform", "stroke", "fill-opacity", "stroke-opacity", "stroke-width"];
    if (Object.keys(values).some(key => !allowed.includes(key))) unsupported("Unsupported shape attribute.");
    const color = fill(values.fill, parent.fill);
    if (color === undefined) unsupported("Unsupported fill.");
    const stroke = fill(values.stroke, parent.stroke), strokeWidth = resolveStrokeWidth(values["stroke-width"], parent.strokeWidth);
    const fillOpacity = paintOpacity(values["fill-opacity"], parent.fillOpacity), strokeOpacity = paintOpacity(values["stroke-opacity"], parent.strokeOpacity);
    if (stroke === undefined || strokeWidth === undefined || strokeWidth < 0 || fillOpacity === undefined || strokeOpacity === undefined) { unsupported("Unsupported paint opacity, stroke or stroke width."); return; }
    let x: number | undefined, y: number | undefined, w: number | undefined, h: number | undefined;
    if (tag.local === "rect") { x = scalar(values.x ?? "0"); y = scalar(values.y ?? "0"); w = scalar(values.width); h = scalar(values.height); }
    else {
      const cx = scalar(values.cx ?? "0"), cy = scalar(values.cy ?? "0"), rx = scalar(values[tag.local === "circle" ? "r" : "rx"]), ry = scalar(values[tag.local === "circle" ? "r" : "ry"]);
      if (cx !== undefined && cy !== undefined && rx !== undefined && ry !== undefined) { x = cx - rx; y = cy - ry; w = rx * 2; h = ry * 2; }
    }
    const local = transform(values.transform), matrix = local && compose(parent.matrix, local);
    if (!matrix) { unsupported("Unsupported or invalid shape transform."); return; }
    if (x === undefined || y === undefined || w === undefined || h === undefined || w <= 0 || h <= 0 || !viewport) { unsupported("Shape geometry is invalid."); return; }
    const x1 = x * matrix.sx + matrix.tx, x2 = (x + w) * matrix.sx + matrix.tx, y1 = y * matrix.sy + matrix.ty, y2 = (y + h) * matrix.sy + matrix.ty;
    x = Math.min(x1, x2); y = Math.min(y1, y2); w = Math.abs(x2 - x1); h = Math.abs(y2 - y1);
    if (![x, y, w, h].every(Number.isFinite) || w <= 0 || h <= 0 || x < viewport.x || y < viewport.y || x + w > viewport.x + viewport.width || y + h > viewport.y + viewport.height) { unsupported("Transformed shape crosses the viewBox or is invalid."); return; }
    const mappedX = (x - viewport.x) * viewport.sx + viewport.dx, mappedY = (y - viewport.y) * viewport.sy + viewport.dy;
    const mappedWidth = w * viewport.sx, mappedHeight = h * viewport.sy;
    if (width === undefined || height === undefined || ![mappedX, mappedY, mappedWidth, mappedHeight].every(Number.isFinite) || mappedX < 0 || mappedY < 0 || mappedX + mappedWidth > width || mappedY + mappedHeight > height) { unsupported("Shape geometry crosses the viewport."); return; }
    const visibleStroke = stroke !== null && strokeWidth > 0;
    if (visibleStroke && tag.local === "rect") { unsupported("Visible rectangle strokes are unsupported."); return; }
    const strokeScaleX = Math.abs(matrix.sx) * viewport.sx, strokeScaleY = Math.abs(matrix.sy) * viewport.sy;
    if (visibleStroke && (strokeScaleX !== strokeScaleY || !Number.isFinite(strokeScaleX * strokeWidth))) { unsupported("Nonuniform or invalid stroke scaling is unsupported."); return; }
    const halfX = visibleStroke ? strokeWidth * Math.abs(matrix.sx) / 2 : 0, halfY = visibleStroke ? strokeWidth * Math.abs(matrix.sy) / 2 : 0;
    const halfMapped = visibleStroke ? strokeWidth * strokeScaleX / 2 : 0;
    if (x - halfX < viewport.x || y - halfY < viewport.y || x + w + halfX > viewport.x + viewport.width || y + h + halfY > viewport.y + viewport.height || mappedX - halfMapped < 0 || mappedY - halfMapped < 0 || mappedX + mappedWidth + halfMapped > width || mappedY + mappedHeight + halfMapped > height) { unsupported("Shape stroke crosses the viewBox or viewport."); return; }
    elements.push({ type: "shape", shape: tag.local === "rect" ? "rect" : "ellipse", box: { x: mappedX * SCALE, y: mappedY * SCALE, width: mappedWidth * SCALE, height: mappedHeight * SCALE }, style: { fill: color == null ? { type: "none" } : solid(color, fillOpacity), stroke: visibleStroke ? { paint: solid(stroke, strokeOpacity), width: strokeWidth * strokeScaleX * SCALE } : { paint: { type: "none" } } } });
  });
  parser.on("closetag", () => { depth--; contexts.pop(); });
  try { parser.write(input).close(); } catch (error) {
    if (error instanceof SvgImportError) throw error;
    throw new SvgImportError("invalid-xml", "SVG XML is malformed or contains unsupported entity references.");
  }
  if (!root || width === undefined || height === undefined || issues.length) return Object.freeze({ status: "unsupported", issues: Object.freeze(issues) });
  return Object.freeze({ status: "editable", width: width * SCALE, height: height * SCALE, elements: Object.freeze(elements), issues: Object.freeze(warnings) });
}
