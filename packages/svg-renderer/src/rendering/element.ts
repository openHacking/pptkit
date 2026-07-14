import type {
  ArrowType,
  Box,
  ElementAction,
  NormalizedAsset,
  NormalizedPaint,
  NormalizedPresentationTheme,
  NormalizedStrokeStyle,
} from "@pptkit/core";
import type { LayoutElement, LayoutGroupElement } from "@pptkit/layout";
import type { SvgAssetResolver, SvgRenderWarning } from "../types/public.js";
import { escapeXml, safeActionUrl, safeId, safeUrl } from "./escape.js";
import { paintAttributes, strokeAttributes, transformAttribute } from "./style.js";
import { canRenderTextNatively, textHtml, textSvg } from "./text.js";

export interface RenderContext {
  slideId: string;
  theme: NormalizedPresentationTheme;
  assets: ReadonlyMap<string, NormalizedAsset>;
  resolveAsset?: SvgAssetResolver;
  assetCache: Map<string, Promise<string | undefined>>;
  warnings: SvgRenderWarning[];
  defs: string[];
}

function accessibility(element: LayoutElement): string {
  const identity = ` data-pptkit-element-id="${escapeXml(element.id)}"`;
  if (element.accessibility.decorative) return `${identity} aria-hidden="true"`;
  return `${identity} role="img" aria-label="${escapeXml(element.accessibility.description ?? element.name)}"`;
}

function actionWrapper(element: LayoutElement, body: string, context: RenderContext): string {
  const action: ElementAction | undefined = element.action;
  if (action === undefined) return body;
  if (action.type === "slide") {
    context.warnings.push({
      code: "slide-action-not-linked",
      message: `Slide action targeting ${action.slideId} is metadata-only in standalone SVG output.`,
      slideId: context.slideId,
      elementId: element.id,
    });
    return `<g data-pptkit-slide-target="${escapeXml(action.slideId)}">${body}</g>`;
  }
  const url = safeActionUrl(action.url);
  if (url === undefined) {
    context.warnings.push({ code: "unsafe-action-url", message: "An action URL with an unsupported protocol was omitted.", slideId: context.slideId, elementId: element.id });
    return body;
  }
  return `<a href="${escapeXml(url)}" rel="noopener noreferrer"${action.tooltip === undefined ? "" : ` aria-label="${escapeXml(action.tooltip)}"`}>${body}</a>`;
}

function geometry(element: Extract<LayoutElement, { type: "shape" }>, context: RenderContext): string {
  const { x, y, width, height } = element.box;
  const attrs = `${paintAttributes(element.style.fill, context.theme, "fill")} ${strokeAttributes(element.style.stroke, context.theme)}`;
  if (element.shape === "rect") return `<rect x="${x}" y="${y}" width="${width}" height="${height}" ${attrs}/>`;
  if (element.shape === "roundRect") return `<rect x="${x}" y="${y}" width="${width}" height="${height}" rx="${Math.min(width, height) * 0.08}" ${attrs}/>`;
  if (element.shape === "ellipse") return `<ellipse cx="${x + width / 2}" cy="${y + height / 2}" rx="${width / 2}" ry="${height / 2}" ${attrs}/>`;
  const points = element.shape === "triangle"
    ? [[x + width / 2, y], [x + width, y + height], [x, y + height]]
    : element.shape === "diamond"
      ? [[x + width / 2, y], [x + width, y + height / 2], [x + width / 2, y + height], [x, y + height / 2]]
      : element.shape === "arrow"
        ? [[x, y + height * 0.25], [x + width * 0.62, y + height * 0.25], [x + width * 0.62, y], [x + width, y + height / 2], [x + width * 0.62, y + height], [x + width * 0.62, y + height * 0.75], [x, y + height * 0.75]]
        : [[x, y], [x + width * 0.65, y], [x + width, y + height / 2], [x + width * 0.65, y + height], [x, y + height], [x + width * 0.35, y + height / 2]];
  return `<polygon points="${points.map((point) => point.join(",")).join(" ")}" ${attrs}/>`;
}

function foreignObject(box: Box, html: string, transform: string, opacity: number, accessibilityValue: string, overflowVisible = false): string {
  return `<foreignObject x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}"${overflowVisible ? ' overflow="visible"' : ""}${transform} opacity="${opacity}"${accessibilityValue}>${html}</foreignObject>`;
}

function markerShape(type: ArrowType, color: string): string {
  if (type === "oval") return `<circle cx="5" cy="5" r="3.5" fill="${color}"/>`;
  if (type === "diamond") return `<path d="M1 5 L5 1 L9 5 L5 9 Z" fill="${color}"/>`;
  if (type === "stealth") return `<path d="M1 1 L9 5 L1 9 L4 5 Z" fill="${color}"/>`;
  if (type === "arrow") return `<path d="M1 1 L9 5 L1 9" fill="none" stroke="${color}" stroke-width="1.5"/>`;
  return `<path d="M1 1 L9 5 L1 9 Z" fill="${color}"/>`;
}

function marker(type: ArrowType, end: "start" | "end", elementId: string, style: NormalizedStrokeStyle, context: RenderContext): string | undefined {
  if (type === "none" || style.paint.type === "none") return undefined;
  const id = safeId(`${context.slideId}-${elementId}-${end}-${type}`);
  const color = style.paint.type === "solid" ? (typeof style.paint.color === "string" ? style.paint.color : context.theme.colors[style.paint.color.theme]) : "000000";
  const cssColor = color.startsWith("#") ? color : `#${color}`;
  context.defs.push(`<marker id="${id}" markerWidth="10" markerHeight="10" refX="${end === "end" ? 9 : 1}" refY="5" orient="auto-start-reverse" markerUnits="strokeWidth" viewBox="0 0 10 10">${markerShape(type, cssColor)}</marker>`);
  return id;
}

async function resolveImage(asset: NormalizedAsset, context: RenderContext): Promise<string | undefined> {
  let pending = context.assetCache.get(asset.id);
  if (pending === undefined) {
    pending = (async () => {
      if (context.resolveAsset !== undefined) {
        try {
          const resolved = await context.resolveAsset(asset);
          if (resolved === undefined) context.warnings.push({ code: "asset-unresolved", message: "The asset resolver did not return an image URL.", slideId: context.slideId, assetId: asset.id });
          return resolved;
        } catch (error) {
          context.warnings.push({ code: "asset-resolver-failed", message: `The asset resolver failed: ${error instanceof Error ? error.message : String(error)}`, slideId: context.slideId, assetId: asset.id });
          return undefined;
        }
      }
      if (asset.source.type === "url") return asset.source.value;
      context.warnings.push({ code: "asset-path-unsupported", message: "Local path assets require a custom SvgAssetResolver.", slideId: context.slideId, assetId: asset.id });
      return undefined;
    })();
    context.assetCache.set(asset.id, pending);
  }
  return pending;
}

function missingImage(element: Extract<LayoutElement, { type: "image" }>, context: RenderContext): string {
  const { x, y, width, height } = element.box;
  return `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${accessibility(element)}><rect x="${x}" y="${y}" width="${width}" height="${height}" fill="#F3F4F6" stroke="#9CA3AF" stroke-dasharray="6 4"/><path d="M${x + width * 0.15} ${y + height * 0.75} L${x + width * 0.42} ${y + height * 0.45} L${x + width * 0.6} ${y + height * 0.62} L${x + width * 0.78} ${y + height * 0.35} L${x + width * 0.9} ${y + height * 0.75}" fill="none" stroke="#9CA3AF" stroke-width="2"/></g>`;
}

async function imageElement(element: Extract<LayoutElement, { type: "image" }>, context: RenderContext): Promise<string> {
  const asset = context.assets.get(element.assetId);
  if (asset === undefined) {
    context.warnings.push({ code: "asset-missing", message: `Image asset ${element.assetId} was not found.`, slideId: context.slideId, elementId: element.id, assetId: element.assetId });
    return missingImage(element, context);
  }
  const href = await resolveImage(asset, context);
  if (href === undefined || safeUrl(href) === undefined) {
    if (href !== undefined) context.warnings.push({ code: "unsafe-asset-url", message: "An image URL with an unsupported protocol was omitted.", slideId: context.slideId, elementId: element.id, assetId: asset.id });
    return missingImage(element, context);
  }
  const clipId = safeId(`${context.slideId}-${element.id}-clip`);
  const box = element.box;
  context.defs.push(`<clipPath id="${clipId}"><rect x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}"/></clipPath>`);
  let image: string;
  if (element.fit === "crop" && asset.width !== undefined && asset.height !== undefined) {
    const visibleWidth = Math.max(0.0001, 1 - element.crop.left - element.crop.right);
    const visibleHeight = Math.max(0.0001, 1 - element.crop.top - element.crop.bottom);
    const width = box.width / visibleWidth;
    const height = box.height / visibleHeight;
    image = `<image href="${escapeXml(href)}" x="${box.x - width * element.crop.left}" y="${box.y - height * element.crop.top}" width="${width}" height="${height}" preserveAspectRatio="none"/>`;
  } else {
    if (element.fit === "crop") context.warnings.push({ code: "image-crop-degraded", message: "Image crop requires intrinsic asset dimensions; preview fell back to centered cover.", slideId: context.slideId, elementId: element.id, assetId: asset.id });
    image = `<image href="${escapeXml(href)}" x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}" preserveAspectRatio="${element.fit === "stretch" ? "none" : "xMidYMid slice"}"/>`;
  }
  return actionWrapper(element, `<g${transformAttribute(box, element.transform)} opacity="${element.opacity}" clip-path="url(#${clipId})"${accessibility(element)}>${image}</g>`, context);
}

function tableHtml(element: Extract<LayoutElement, { type: "table" }>, context: RenderContext): string {
  const columns = element.columns.map((width) => `<col style="width:${width}px"/>`).join("");
  const rows = element.rows.map((row) => {
    const cells = row.cells.map((cell) => {
      const fill = cell.style.fill.type === "none" ? "transparent" : (typeof cell.style.fill.color === "string" ? cell.style.fill.color : context.theme.colors[cell.style.fill.color.theme]);
      const fillColor = fill === "transparent" || fill.startsWith("#") ? fill : `#${fill}`;
      const stroke = cell.style.stroke.paint.type === "none" ? "transparent" : (typeof cell.style.stroke.paint.color === "string" ? cell.style.stroke.paint.color : context.theme.colors[cell.style.stroke.paint.color.theme]);
      const strokeColor = stroke === "transparent" || stroke.startsWith("#") ? stroke : `#${stroke}`;
      const style = `box-sizing:border-box;background:${fillColor};border:${cell.style.stroke.width}px solid ${strokeColor};padding:${cell.style.margin.top}px ${cell.style.margin.right}px ${cell.style.margin.bottom}px ${cell.style.margin.left}px;vertical-align:${cell.style.verticalAlign};overflow:hidden`;
      const frame = { margin: { top: 0, right: 0, bottom: 0, left: 0 }, verticalAlign: cell.style.verticalAlign, wrap: true, autoFit: { mode: "none" as const } };
      return `<td rowspan="${cell.rowSpan}" colspan="${cell.colSpan}" style="${style}">${textHtml(cell.content, frame, { theme: context.theme, warnings: context.warnings, slideId: context.slideId, elementId: element.id })}</td>`;
    }).join("");
    return `<tr style="height:${row.height}px">${cells}</tr>`;
  }).join("");
  return `<div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;overflow:hidden"><table style="width:100%;height:100%;table-layout:fixed;border-collapse:collapse;border-spacing:0"><colgroup>${columns}</colgroup><tbody>${rows}</tbody></table></div>`;
}

function groupElement(element: LayoutGroupElement, context: RenderContext): Promise<string> {
  return (async () => {
    const renderedChildren: string[] = [];
    for (const child of element.children) renderedChildren.push(await renderElement(child, context));
    const children = renderedChildren.join("");
    const box = element.box;
    const cx = box.x + box.width / 2;
    const cy = box.y + box.height / 2;
    const outer = `translate(${cx} ${cy}) rotate(${element.transform.rotation}) scale(${element.transform.flipH ? -1 : 1} ${element.transform.flipV ? -1 : 1}) translate(${-cx} ${-cy}) translate(${box.x} ${box.y}) scale(${box.width / element.coordinateSize.width} ${box.height / element.coordinateSize.height})`;
    return actionWrapper(element, `<g transform="${outer}" opacity="${element.opacity}"${accessibility(element)}>${children}</g>`, context);
  })();
}

export async function renderElement(element: LayoutElement, context: RenderContext): Promise<string> {
  if (element.hidden) return "";
  if (element.type === "image") return imageElement(element, context);
  if (element.type === "group") return groupElement(element, context);
  if (element.type === "text") {
    const textContext = { theme: context.theme, warnings: context.warnings, slideId: context.slideId, elementId: element.id };
    if (canRenderTextNatively(element.content)) {
      const content = textSvg(element.content, element.frame, element.box, textContext);
      return actionWrapper(element, `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${accessibility(element)}>${content}</g>`, context);
    }
    const html = textHtml(element.content, element.frame, textContext);
    return actionWrapper(element, foreignObject(element.box, html, transformAttribute(element.box, element.transform), element.opacity, accessibility(element), true), context);
  }
  if (element.type === "shape") {
    const transform = transformAttribute(element.box, element.transform);
    const shape = geometry(element, context);
    const text = element.text === undefined ? "" : foreignObject(element.box, textHtml(element.text.content, element.text.frame, { theme: context.theme, warnings: context.warnings, slideId: context.slideId, elementId: element.id }), "", 1, "", true);
    return actionWrapper(element, `<g${transform} opacity="${element.opacity}"${accessibility(element)}>${shape}${text}</g>`, context);
  }
  if (element.type === "connector") {
    const points = [element.start, ...element.route, element.end];
    const start = marker(element.style.beginArrow, "start", element.id, element.style, context);
    const end = marker(element.style.endArrow, "end", element.id, element.style, context);
    const markers = `${start === undefined ? "" : ` marker-start="url(#${start})"`}${end === undefined ? "" : ` marker-end="url(#${end})"`}`;
    const path = `<polyline points="${points.map((point) => `${point.x},${point.y}`).join(" ")}" fill="none" ${strokeAttributes(element.style, context.theme)}${markers}/>`;
    return actionWrapper(element, `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${accessibility(element)}>${path}</g>`, context);
  }
  if (element.type === "table") {
    return actionWrapper(element, foreignObject(element.box, tableHtml(element, context), transformAttribute(element.box, element.transform), element.opacity, accessibility(element)), context);
  }
  return "";
}

export function backgroundRect(box: Box, paint: NormalizedPaint, theme: NormalizedPresentationTheme): string {
  return `<rect x="0" y="0" width="${box.width}" height="${box.height}" ${paintAttributes(paint, theme, "fill")}/>`;
}
