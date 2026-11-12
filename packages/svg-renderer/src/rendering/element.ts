import type {
  ArrowType,
  Box,
  ElementAction,
  NormalizedAsset,
  NormalizedTextFrameStyle,
  NormalizedTextParagraph,
  NormalizedPaint,
  NormalizedBackgroundPaint,
  NormalizedPresentationTheme,
  NormalizedStrokeStyle,
} from "@pptkit/core";
import type { LayoutElement, LayoutGroupElement } from "@pptkit/layout";
import { formatListMarker } from "@pptkit/core";
import { resolveShapeTextBox, resolveTableBorders, TableBorderError } from "@pptkit/layout";
import { escapeXml, safeActionUrl, safeUrl } from "./escape.js";
import { chartElement } from "./chart.js";
import { colorValue, paintAttributes, strokeAttributes, transformAttribute } from "./style.js";
import { canRenderMixedNoWrapNatively, canRenderNativeParagraphs, canRenderShapeTabTextNatively, canRenderTextNatively, textHtml, textSvg } from "./text.js";
import { readingOrderOwnership, wrapReadingOrder } from "./reading-order.js";
import { accessibility, nextDefinitionId, recordFont, type RenderContext } from "./context.js";
import { probePngAnimation } from "./png-animation-probe.js";
import { probeJpegPrecision } from "./jpeg-precision-probe.js";
import { checkRenderAbort, waitForResolver } from "./abort.js";
import { SvgRenderAbortError } from "../types/public.js";
import { shapePolygonPoints } from "./shape-polygon.js";

export type { RenderContext };

function actionWrapper(element: LayoutElement, body: string, context: RenderContext): string {
  const action: ElementAction | undefined = element.action;
  if (action === undefined) return body;
  if (action.type === "slide") {
    const target = context.slideActionTargets?.get(action.slideId);
    if (target !== undefined) return `<a href="#${target}" data-pptkit-slide-target="${escapeXml(action.slideId)}"${context.insideDecorativeGroup || element.accessibility.decorative ? ' tabindex="-1"' : ""}${action.tooltip === undefined ? "" : ` aria-label="${escapeXml(action.tooltip)}"`}>${body}</a>`;
    context.warnings.push({
      code: "slide-action-not-linked",
      message: context.slideActionTargets ? `Slide action targeting ${action.slideId} has no selected document target.` : `Slide action targeting ${action.slideId} is metadata-only in standalone SVG output.`,
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
  return `<a href="${escapeXml(url)}" rel="noopener noreferrer"${context.insideDecorativeGroup ? ' tabindex="-1"' : ""}${action.tooltip === undefined ? "" : ` aria-label="${escapeXml(action.tooltip)}"`}>${body}</a>`;
}

function geometry(element: Extract<LayoutElement, { type: "shape" }>, context: RenderContext): string {
  const { x, y, width, height } = element.box;
  let fill: string;
  if (element.style.fill.type === "linearGradient") {
    const id = nextDefinitionId(context);
    context.defs.push(linearGradientDefinition(element.box, element.style.fill, context.theme, id));
    context.warnings.push({ code: "shape-gradient-approximation", message: "SVG gradient interpolation can differ from PowerPoint.", slideId: context.slideId, elementId: element.id });
    fill = `fill="url(#${id})"`;
  } else if (element.style.fill.type === "pattern") {
    const paint = element.style.fill, id = nextDefinitionId(context);
    const fg = colorValue(paint.foreground, context.theme), bg = colorValue(paint.background, context.theme);
    // Use one filled outline for the cross: separate strokes compound opacity.
    const down = "M0 0H.7L8 7.3V8H7.3L0 .7Z M7.3 0H8V.7Z M0 7.3L.7 8H0Z", up = "M7.3 0H8V.7L.7 8H0V7.3Z M0 0H.7L0 .7Z M8 7.3V8H7.3Z";
    const paths = { horizontal: "M0 0H8V.5H0Z", vertical: "M0 0H.5V8H0Z", cross: "M0 0H8V.5H.5V8H0Z", downwardDiagonal: down, upwardDiagonal: up, diagonalCross: `${down} ${up}` } as const;
    const percent = paint.pattern.startsWith("percent") ? Number(paint.pattern.slice(7)) : undefined;
    // A single dot (or its inverse) gives the requested nominal tile coverage.
    // Office preset textures are different; the approximation warning applies.
    const radius = percent === undefined ? 0 : Math.sqrt(64 * Math.min(percent, 100 - percent) / (100 * Math.PI));
    const circle = `M4 ${4 - radius}a${radius} ${radius} 0 1 0 0 ${2 * radius}a${radius} ${radius} 0 1 0 0 ${-2 * radius}Z`;
    const path = percent === undefined ? paths[paint.pattern as keyof typeof paths] : percent <= 50 ? circle : `M0 0H8V8H0Z ${circle}`;
    const lines = `<path d="${path}" fill="${fg}" fill-opacity="${paint.foregroundOpacity}"${percent !== undefined && percent > 50 ? ' fill-rule="evenodd"' : ""}/>`;
    context.defs.push(`<pattern id="${id}" patternUnits="userSpaceOnUse" x="${x}" y="${y}" width="8" height="8"><rect width="8" height="8" fill="${bg}" fill-opacity="${paint.backgroundOpacity}"/>${lines}</pattern>`);
    context.warnings.push({ code: "shape-pattern-approximation", message: "SVG pattern spacing and rendering are approximate and can differ from PowerPoint.", slideId: context.slideId, elementId: element.id });
    fill = `fill="url(#${id})"`;
  } else fill = paintAttributes(element.style.fill, context.theme, "fill");
  const attrs = `${fill} ${strokeAttributes(element.style.stroke, context.theme)}`;
  if (element.shape === "rect") return `<rect x="${x}" y="${y}" width="${width}" height="${height}" ${attrs}/>`;
  if (element.shape === "roundRect") return `<rect x="${x}" y="${y}" width="${width}" height="${height}" rx="${Math.min(width, height) * 0.08}" ${attrs}/>`;
  if (element.shape === "ellipse") return `<ellipse cx="${x + width / 2}" cy="${y + height / 2}" rx="${width / 2}" ry="${height / 2}" ${attrs}/>`;
  const points = shapePolygonPoints(element.box, element.shape, element.slant);
  return `<polygon points="${points.map((point) => `${point.x},${point.y}`).join(" ")}" ${attrs}/>`;
}

function foreignObject(box: Box, html: string, transform: string, opacity: number, accessibilityValue: string, overflowVisible = false): string {
  return `<foreignObject x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}"${overflowVisible ? ' overflow="visible"' : ""}${transform} opacity="${opacity}"${accessibilityValue}>${html}</foreignObject>`;
}

function warnNativeMixedText(element: LayoutElement, context: RenderContext): void {
  context.warnings.push({ code: "svg-native-mixed-text-approximation", message: "Native SVG uses approximate line breaks, baselines and font metrics and may differ from browser XHTML or PowerPoint.", slideId: context.slideId, elementId: element.id });
}

function nativeTextOptIn(content: NormalizedTextParagraph[], frame: NormalizedTextFrameStyle, context: RenderContext): boolean {
  if (context.mixedTextMode === "native") return canRenderNativeParagraphs(content);
  return context.mixedTextMode === "native-no-wrap" && canRenderMixedNoWrapNatively(content, frame);
}

function markerShape(type: ArrowType, color: string): string {
  if (type === "oval") return `<circle cx="5" cy="5" r="3.5" fill="${color}"/>`;
  if (type === "diamond") return `<path d="M1 5 L5 1 L9 5 L5 9 Z" fill="${color}"/>`;
  if (type === "stealth") return `<path d="M1 1 L9 5 L1 9 L4 5 Z" fill="${color}"/>`;
  if (type === "arrow") return `<path d="M1 1 L9 5 L1 9" fill="none" stroke="${color}" stroke-width="1.5"/>`;
  return `<path d="M1 1 L9 5 L1 9 Z" fill="${color}"/>`;
}

function marker(type: ArrowType, end: "start" | "end", style: NormalizedStrokeStyle, context: RenderContext): string | undefined {
  if (type === "none" || style.paint.type === "none") return undefined;
  const id = nextDefinitionId(context);
  const color = style.paint.type === "solid" ? (typeof style.paint.color === "string" ? style.paint.color : context.theme.colors[style.paint.color.theme]) : "000000";
  const cssColor = color.startsWith("#") ? color : `#${color}`;
  context.defs.push(`<marker id="${id}" markerWidth="10" markerHeight="10" refX="${end === "end" ? 9 : 1}" refY="5" orient="auto-start-reverse" markerUnits="strokeWidth" viewBox="0 0 10 10">${markerShape(type, cssColor)}</marker>`);
  return id;
}

async function resolveImage(asset: NormalizedAsset, context: RenderContext): Promise<string | undefined> {
  let pending = context.assetCache.get(asset.id);
  if (pending === undefined) {
    pending = (async () => {
      if (asset.mimeType?.trim().toLowerCase() === "image/apng") context.warnings.push({ code: "apng-preview-playback-unverified", message: "The asset is declared as APNG; SVG image animation and timing depend on the viewing environment and any resolver.", slideId: context.slideId, assetId: asset.id });
      if (context.pngAnimationDiagnostics === "warn") {
        const declared = asset.mimeType?.trim().toLowerCase();
        if (context.resolveAsset === undefined && asset.source.type === "bytes" && (declared === undefined || declared === "image/png" || declared === "image/apng")) {
          const header = atob(asset.source.value.slice(0,12));
          if (header.startsWith("\x89PNG\r\n\x1a\n")) {
            const result = probePngAnimation(asset.source.value);
            if (result === "declared" && declared !== "image/apng") context.warnings.push({ code: "apng-preview-playback-unverified", message: "PNG contains an animation declaration; its frames were not validated and SVG playback depends on the viewing environment.", slideId: context.slideId, assetId: asset.id });
            if (result === "incomplete") context.warnings.push({ code: "png-preview-inspection-incomplete", message: "The bounded PNG animation declaration scan was incomplete; animation status is unknown.", slideId: context.slideId, assetId: asset.id });
          }
        } else if ((declared === "image/png" || declared === "image/apng") && (context.resolveAsset !== undefined || asset.source.type !== "bytes")) context.warnings.push({ code: "png-preview-inspection-incomplete", message: "PNG animation cannot be inspected from resolver output or non-byte sources; animation status is unknown.", slideId: context.slideId, assetId: asset.id });
      }
      if (context.resolveAsset !== undefined) {
        try {
          checkRenderAbort(context.signal);
          const resolved = await waitForResolver(context.resolveAsset(asset, context.signal ? { signal: context.signal } : {}), context.signal);
          checkRenderAbort(context.signal);
          if (resolved === undefined) context.warnings.push({ code: "asset-unresolved", message: "The asset resolver did not return an image URL.", slideId: context.slideId, assetId: asset.id });
          return resolved;
        } catch (error) {
          if (error instanceof SvgRenderAbortError || context.signal?.aborted) throw new SvgRenderAbortError();
          context.warnings.push({ code: "asset-resolver-failed", message: `The asset resolver failed: ${error instanceof Error ? error.message : String(error)}`, slideId: context.slideId, assetId: asset.id });
          return undefined;
        }
      }
      if (asset.source.type === "url") return asset.source.value;
      if (asset.source.type === "bytes") {
        // Preview serializes a self-contained URL; decoding/rendering remains browser-owned.
        const header = atob(asset.source.value.slice(0, 12));
        const mime = header.startsWith("\x89PNG\r\n\x1a\n") ? "image/png" : header.startsWith("\xff\xd8\xff") ? "image/jpeg"
          : header.startsWith("GIF87a") || header.startsWith("GIF89a") ? "image/gif" : undefined;
        const declared = asset.mimeType?.trim().toLowerCase();
        if (mime && (declared === undefined || declared === mime || (mime === "image/png" && declared === "image/apng"))) {
          if (mime === "image/jpeg") {
            const precision = probeJpegPrecision(asset.source.value);
            if (precision === "incomplete") context.warnings.push({ code: "jpeg-preview-inspection-incomplete", message: "The bounded JPEG header scan ended before sample precision could be determined; SVG image decoding and display depend on the viewing environment.", slideId: context.slideId, assetId: asset.id });
            else if (precision !== undefined && precision !== 8) context.warnings.push({ code: "jpeg-preview-precision-unverified", message: `JPEG declares ${precision}-bit sample precision; SVG image decoding and display depend on the viewing environment.`, slideId: context.slideId, assetId: asset.id });
          }
          return `data:${mime};base64,${asset.source.value}`;
        }
        context.warnings.push({ code: "asset-bytes-unsupported", message: "Byte image signature is unsupported or differs from its declared MIME type.", slideId: context.slideId, assetId: asset.id });
        return undefined;
      }
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
  if (context.portabilityDiagnostics === "warn" && !/^data:/i.test(href)) context.warnings.push({ code: "svg-external-image-dependency", message: "This image URL depends on resources outside the SVG; blob URLs also depend on the originating browser session.", slideId: context.slideId, elementId: element.id, assetId: asset.id });
  if (context.portabilityDiagnostics === "warn" && /^data:image\/svg\+xml(?:;|,)/i.test(href)) context.warnings.push({ code: "svg-inline-svg-image-unverified", message: "An embedded SVG image may contain nested external resources or unsupported content and has not been inspected.", slideId: context.slideId, elementId: element.id, assetId: asset.id });
  const clipId = nextDefinitionId(context);
  const box = element.box;
  context.defs.push(`<clipPath id="${clipId}"><rect x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}"/></clipPath>`);
  let image: string;
  if (element.fit === "crop") {
    const visibleWidth = 1 - (element.crop.left + element.crop.right);
    const visibleHeight = 1 - (element.crop.top + element.crop.bottom);
    const width = box.width / visibleWidth;
    const height = box.height / visibleHeight;
    const x = box.x - width * element.crop.left;
    const y = box.y - height * element.crop.top;
    if (![x, y, width, height].every(Number.isFinite)) throw new RangeError(`Image crop geometry for "${element.id}" cannot be represented as finite SVG coordinates.`);
    image = `<image href="${escapeXml(href)}" x="${x}" y="${y}" width="${width}" height="${height}" preserveAspectRatio="none"/>`;
  } else {
    image = `<image href="${escapeXml(href)}" x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}" preserveAspectRatio="${element.fit === "stretch" ? "none" : element.fit === "contain" ? "xMidYMid meet" : "xMidYMid slice"}"/>`;
  }
  return actionWrapper(element, `<g${transformAttribute(box, element.transform)} opacity="${element.opacity}" clip-path="url(#${clipId})"${accessibility(element)}>${image}</g>`, context);
}

function verticalCssDirection(direction: NormalizedTextFrameStyle["textDirection"]) {
  return direction !== undefined && direction !== "horizontal" && direction !== "rotate90" && direction !== "rotate270";
}

function textRenderContext(element: LayoutElement, context: RenderContext): Parameters<typeof textHtml>[2] {
  return {
    theme: context.theme, warnings: context.warnings, slideId: context.slideId, elementId: element.id,
    ...(context.slideActionTargets === undefined ? {} : { slideActionTargets: context.slideActionTargets }),
    ...(context.fontDependencies === undefined && context.fontCoverage === undefined ? {} : { recordFont: (family: string, bold: boolean, italic: boolean, text: string, fontSize: number, language?: string) => recordFont(context, family, bold, italic, text, fontSize, language, element.id) }),
    skipLinkTabStop: context.insideDecorativeGroup === true || element.accessibility.decorative === true,
  };
}

function tableHtml(element: Extract<LayoutElement, { type: "table" }>, context: RenderContext): string {
  const textContext = textRenderContext(element, context);
  const label = element.accessibility.decorative ? "" : ` aria-label="${escapeXml(element.accessibility.description ?? element.name)}"`;
  const cellPaint = (paint: typeof element.rows[number]["cells"][number]["style"]["fill"]): string => {
    if (paint.type === "none") return "transparent";
    const color = colorValue(paint.color, context.theme);
    if (paint.opacity === 1) return color;
    return `rgba(${parseInt(color.slice(1, 3), 16)},${parseInt(color.slice(3, 5), 16)},${parseInt(color.slice(5, 7), 16)},${paint.opacity})`;
  };
  const directed = (html: string, direction: NormalizedTextFrameStyle["textDirection"]): string => {
    if (direction === undefined || direction === "horizontal" || verticalCssDirection(direction)) return html;
    context.warnings.push({ code: "table-text-direction-approximation", message: "SVG rotated table text uses CSS writing modes; native Office glyph orientation, margins and wrapping may differ.", slideId: context.slideId, elementId: element.id });
    const mode = direction === "rotate90" ? "sideways-rl" : "sideways-lr";
    return `<div style="width:100%;height:100%;writing-mode:${mode}">${html}</div>`;
  };
  if (element.cellLayout !== undefined) {
    const rows: string[][] = element.rows.map(() => []);
    for (const { rowIndex, cellIndex, columnIndex, box } of element.cellLayout) {
      const cell = element.rows[rowIndex]!.cells[cellIndex]!;
      const fill = cellPaint(cell.style.fill);
      const stroke = cellPaint(cell.style.stroke.paint);
      const frame = { ...(verticalCssDirection(cell.style.textDirection) && cell.style.textDirection !== undefined ? { textDirection: cell.style.textDirection } : {}), margin: cell.style.margin, verticalAlign: cell.style.verticalAlign, wrap: true, autoFit: { mode: "none" as const } };
      const content = directed(textHtml(cell.content, frame, textContext), cell.style.textDirection);
      const border = element.borderGrid ? "" : `<div aria-hidden="true" style="position:absolute;inset:0;pointer-events:none;box-sizing:border-box;border:${cell.style.stroke.width}px solid ${stroke}"></div>`;
      const role = cell.headerScope === "row" ? "rowheader" : cell.headerScope === "column" ? "columnheader" : "cell";
      rows[rowIndex]!.push(`<div role="${role}" aria-colindex="${columnIndex + 1}" aria-rowindex="${rowIndex + 1}" aria-colspan="${cell.colSpan}" aria-rowspan="${cell.rowSpan}" data-pptkit-cell="${rowIndex}:${cellIndex}" style="position:absolute;left:${box.x}px;top:${box.y}px;width:${box.width}px;height:${box.height}px;overflow:hidden;background:${fill}">${content}${border}</div>`);
    }
    const cells = rows.map((row) => `<div role="row" style="display:contents">${row.join("")}</div>`).join("");
    return `<div xmlns="http://www.w3.org/1999/xhtml" role="table"${label} aria-rowcount="${element.rows.length}" aria-colcount="${element.columns.length}" style="position:relative;width:100%;height:100%;overflow:hidden">${cells}</div>`;
  }
  const columns = element.columns.map((width) => `<col style="width:${width}px"/>`).join("");
  const rows = element.rows.map((row) => {
    const cells = row.cells.map((cell) => {
      const fillColor = cellPaint(cell.style.fill);
      const strokeColor = cellPaint(cell.style.stroke.paint);
      const style = `box-sizing:border-box;background:${fillColor};border:${cell.style.stroke.width}px solid ${strokeColor};padding:${cell.style.margin.top}px ${cell.style.margin.right}px ${cell.style.margin.bottom}px ${cell.style.margin.left}px;vertical-align:${cell.style.verticalAlign};overflow:hidden`;
      const frame = { ...(verticalCssDirection(cell.style.textDirection) && cell.style.textDirection !== undefined ? { textDirection: cell.style.textDirection } : {}), margin: { top: 0, right: 0, bottom: 0, left: 0 }, verticalAlign: cell.style.verticalAlign, wrap: true, autoFit: { mode: "none" as const } };
      const scope = cell.headerScope === "row" ? ' scope="row"' : cell.headerScope === "column" ? ' scope="col"' : "";
      const tag = scope ? "th" : "td";
      return `<${tag}${scope} rowspan="${cell.rowSpan}" colspan="${cell.colSpan}" style="${style}${scope ? ";font-weight:inherit;text-align:inherit" : ""}">${directed(textHtml(cell.content, frame, textContext), cell.style.textDirection)}</${tag}>`;
    }).join("");
    return `<tr style="height:${row.height}px">${cells}</tr>`;
  }).join("");
  return `<div xmlns="http://www.w3.org/1999/xhtml" style="width:100%;height:100%;overflow:hidden"><table${label} style="width:100%;height:100%;table-layout:fixed;border-collapse:collapse;border-spacing:0"><colgroup>${columns}</colgroup><tbody>${rows}</tbody></table></div>`;
}

function tableBorderSvg(element: Extract<LayoutElement, { type: "table" }>, context: RenderContext): string {
  const grid = element.borderGrid!;
  if (grid.schemaVersion !== "pptkit-table-border-grid-v1" || grid.rowCount !== element.rows.length || grid.columnCount !== element.columns.length || !Array.isArray(grid.horizontal) || !Array.isArray(grid.vertical) || grid.horizontal.length !== (grid.rowCount + 1) * grid.columnCount || grid.vertical.length !== grid.rowCount * (grid.columnCount + 1) || !Array.isArray(element.cellLayout) || element.cellLayout.length !== element.rows.reduce((sum, row) => sum + row.cells.length, 0)) throw new TypeError("Unsupported or inconsistent table border Layout metadata.");
  const xs = [element.box.x], ys = [element.box.y];
  for (const width of element.columns) xs.push(xs[xs.length - 1]! + width);
  for (const row of element.rows) ys.push(ys[ys.length - 1]! + row.height);
  const paths: string[] = [];
  const line = (style: NormalizedStrokeStyle, x1: number, y1: number, x2: number, y2: number, at: string) => {
    if (style.paint.type === "none" || style.width === 0 || style.paint.opacity === 0) return;
    paths.push(`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" ${strokeAttributes(style, context.theme)} data-pptkit-table-border="${at}"/>`);
  };
  for (let row = 0; row <= grid.rowCount; row++) for (let column = 0; column < grid.columnCount; column++) line(grid.horizontal[row * grid.columnCount + column]!, xs[column]!, ys[row]!, xs[column + 1]!, ys[row]!, `h:${row}:${column}`);
  for (let row = 0; row < grid.rowCount; row++) for (let column = 0; column <= grid.columnCount; column++) line(grid.vertical[row * (grid.columnCount + 1) + column]!, xs[column]!, ys[row]!, xs[column]!, ys[row + 1]!, `v:${row}:${column}`);
  return `<g fill="none" pointer-events="none" aria-hidden="true">${paths.join("")}</g>`;
}

function nativeTableSvg(element: Extract<LayoutElement, { type: "table" }>, context: RenderContext): string | undefined {
  if (element.rows.some(row => row.cells.some(cell => ![undefined, "horizontal", "rotate90", "rotate270"].includes(cell.style.textDirection) || !canRenderNativeParagraphs(cell.content)))) return undefined;
  let resolved: ReturnType<typeof resolveTableBorders>;
  try { resolved = resolveTableBorders(element); }
  catch (error) { if (error instanceof TableBorderError) return undefined; throw error; }
  const rows: string[][] = element.rows.map(() => []);
  const textContext = textRenderContext(element, context);
  let hasRotatedText = false;
  for (const { rowIndex, cellIndex, columnIndex, box } of resolved.cells) {
    const cell = element.rows[rowIndex]!.cells[cellIndex]!;
    const absolute = { x: element.box.x + box.x, y: element.box.y + box.y, width: box.width, height: box.height };
    const clip = nextDefinitionId(context);
    context.defs.push(`<clipPath id="${clip}"><rect x="${absolute.x}" y="${absolute.y}" width="${absolute.width}" height="${absolute.height}"/></clipPath>`);
    const frame = { margin: cell.style.margin, verticalAlign: cell.style.verticalAlign, wrap: true, autoFit: { mode: "none" as const } };
    let content: string;
    if (cell.style.textDirection === "rotate90" || cell.style.textDirection === "rotate270") {
      hasRotatedText = true;
      const clockwise = cell.style.textDirection === "rotate90";
      const cx = absolute.x + absolute.width / 2, cy = absolute.y + absolute.height / 2;
      const logical = { x: cx - absolute.height / 2, y: cy - absolute.width / 2, width: absolute.height, height: absolute.width };
      const { top, right, bottom, left } = cell.style.margin;
      const margin = clockwise ? { top: right, right: bottom, bottom: left, left: top } : { top: left, right: top, bottom: right, left: bottom };
      content = `<g transform="rotate(${clockwise ? 90 : 270} ${cx} ${cy})">${textSvg(cell.content, { ...frame, margin }, logical, textContext)}</g>`;
    } else content = textSvg(cell.content, frame, absolute, textContext);
    const role = cell.headerScope === "row" ? "rowheader" : cell.headerScope === "column" ? "columnheader" : "cell";
    const name = cell.content.map(paragraph => formatListMarker(paragraph.style.bullet) + paragraph.runs.map(run => run.text).join("")).join("\n");
    const label = escapeXml(name).replace(/\n/g, "&#10;");
    rows[rowIndex]!.push(`<g role="${role}" aria-colindex="${columnIndex + 1}" aria-rowindex="${rowIndex + 1}" aria-colspan="${cell.colSpan}" aria-rowspan="${cell.rowSpan}" aria-label="${label}" data-pptkit-cell="${rowIndex}:${cellIndex}" clip-path="url(#${clip})"><rect x="${absolute.x}" y="${absolute.y}" width="${absolute.width}" height="${absolute.height}" ${paintAttributes(cell.style.fill, context.theme, "fill")}/>${content}</g>`);
  }
  const clip = nextDefinitionId(context), box = element.box;
  context.defs.push(`<clipPath id="${clip}"><rect x="${box.x}" y="${box.y}" width="${box.width}" height="${box.height}"/></clipPath>`);
  const label = element.accessibility.decorative ? "" : ` aria-label="${escapeXml(element.accessibility.description ?? element.name)}"`;
  const cells = rows.map(row => `<g role="row">${row.join("")}</g>`).join("");
  const borders = tableBorderSvg({ ...element, borderGrid: resolved.grid, cellLayout: resolved.cells }, context);
  if (hasRotatedText) context.warnings.push({ code: "table-text-direction-approximation", message: "Native SVG rotates horizontal text layout; Office glyph positioning, wrapping and margins may differ.", slideId: context.slideId, elementId: element.id });
  context.warnings.push({ code: "svg-native-table-approximation", message: "Native SVG table text wrapping, font metrics and shared borders are approximate and may differ from browser XHTML or PowerPoint.", slideId: context.slideId, elementId: element.id });
  return `<g${transformAttribute(box, element.transform)} opacity="${element.opacity}"${accessibility(element, "child")}><g role="table"${label} aria-rowcount="${element.rows.length}" aria-colcount="${element.columns.length}" clip-path="url(#${clip})">${cells}${borders}</g></g>`;
}

function groupElement(element: LayoutGroupElement, context: RenderContext): Promise<string> {
  return (async () => {
    const renderedChildren: string[] = [];
    const childContext = element.accessibility.decorative && !context.insideDecorativeGroup ? { ...context, insideDecorativeGroup: true as const } : context;
    const readingOrder = context.insideDecorativeGroup || element.accessibility.decorative ? undefined : context.groupReadingOrders?.get(element.id);
    for (const child of element.children) renderedChildren.push(wrapReadingOrder(child, await renderElement(child, childContext), readingOrder));
    const children = renderedChildren.join("");
    const box = element.box;
    const cx = box.x + box.width / 2;
    const cy = box.y + box.height / 2;
    const outer = `translate(${cx} ${cy}) rotate(${element.transform.rotation}) scale(${element.transform.flipH ? -1 : 1} ${element.transform.flipV ? -1 : 1}) translate(${-cx} ${-cy}) translate(${box.x} ${box.y}) scale(${box.width / element.coordinateSize.width} ${box.height / element.coordinateSize.height})`;
    return actionWrapper(element, `<g transform="${outer}" opacity="${element.opacity}"${accessibility(element)}${readingOrderOwnership(readingOrder)}>${children}</g>`, context);
  })();
}

function rotatedText(content: NormalizedTextParagraph[], frame: NormalizedTextFrameStyle, box: Box, context: Parameters<typeof textHtml>[2]): string | undefined {
  if (verticalCssDirection(frame.textDirection)) return foreignObject(box, textHtml(content, frame, context), "", 1, "", true);
  if (frame.textDirection !== "rotate90" && frame.textDirection !== "rotate270") return undefined;
  const clockwise = frame.textDirection === "rotate90";
  const cx = box.x + box.width / 2, cy = box.y + box.height / 2;
  const logical = { x: cx - box.height / 2, y: cy - box.width / 2, width: box.height, height: box.width };
  const { top, right, bottom, left } = frame.margin;
  const margin = clockwise ? { top: right, right: bottom, bottom: left, left: top } : { top: left, right: top, bottom: right, left: bottom };
  context.warnings.push({ code: "text-frame-direction-approximation", message: "Preview rotates a horizontal text layout; native wrapping and font metrics may differ.", slideId: context.slideId, elementId: context.elementId });
  const html = textHtml(content, { ...frame, margin, textDirection: "horizontal" }, context);
  return foreignObject(logical, html, ` transform="rotate(${clockwise ? 90 : 270} ${cx} ${cy})"`, 1, "", true);
}

export async function renderElement(element: LayoutElement, context: RenderContext): Promise<string> {
  const fontDependencies = context.portabilityDiagnostics === "warn" && element.type !== "group" ? new Map<string, { family: string; bold: boolean; italic: boolean }>() : undefined;
  const local = fontDependencies === undefined ? context : { ...context, fontDependencies };
  const output = await renderElementBody(element, local);
  if (context.portabilityDiagnostics === "warn" && element.type !== "group") {
    if (output.includes("<foreignObject ")) context.warnings.push({ code: "svg-foreign-object-dependency", message: "This object uses XHTML foreignObject, which may be omitted by standalone SVG viewers.", slideId: context.slideId, elementId: element.id });
    if (fontDependencies && fontDependencies.size > 0) context.warnings.push({ code: "svg-system-font-dependency", message: "This object requests fonts whose availability, coverage and viewer support are not verified.", slideId: context.slideId, elementId: element.id, fontDependencies: [...fontDependencies.values()] });
  }
  return output;
}

async function renderElementBody(element: LayoutElement, context: RenderContext): Promise<string> {
  checkRenderAbort(context.signal);
  if (element.hidden) return "";
  if (element.type === "image") return imageElement(element, context);
  if (element.type === "group") return groupElement(element, context);
  if (element.type === "text") {
    const textContext = textRenderContext(element, context);
    const rotated = rotatedText(element.content, element.frame, element.box, textContext);
    if (rotated !== undefined) return actionWrapper(element, `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${accessibility(element)}>${rotated}</g>`, context);
    const native = canRenderTextNatively(element.content, element.frame);
    const mixed = !native && nativeTextOptIn(element.content, element.frame, context);
    if (native || mixed) {
      if (mixed) warnNativeMixedText(element, context);
      const content = textSvg(element.content, element.frame, element.box, textContext);
      return actionWrapper(element, `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${accessibility(element)}>${content}</g>`, context);
    }
    const html = textHtml(element.content, element.frame, textContext);
    return actionWrapper(element, foreignObject(element.box, html, transformAttribute(element.box, element.transform), element.opacity, accessibility(element), true), context);
  }
  if (element.type === "shape") {
    const transform = transformAttribute(element.box, element.transform);
    let shape = geometry(element, context);
    const shadow = element.style.shadow;
    if (shadow !== undefined) {
      const id = nextDefinitionId(context), radians = shadow.direction * Math.PI / 180;
      const dx = shadow.distance * Math.cos(radians), dy = shadow.distance * Math.sin(radians);
      const margin = shadow.distance + shadow.blur * 3 + element.style.stroke.width + 1;
      const { x, y, width, height } = element.box;
      context.defs.push(`<filter id="${id}" filterUnits="userSpaceOnUse" x="${x - margin}" y="${y - margin}" width="${width + margin * 2}" height="${height + margin * 2}"><feGaussianBlur in="SourceAlpha" stdDeviation="${shadow.blur / 2}" result="blur"/><feOffset in="blur" dx="${dx}" dy="${dy}" result="offset"/><feFlood flood-color="${colorValue(shadow.color, context.theme)}" flood-opacity="${shadow.opacity}" result="color"/><feComposite in="color" in2="offset" operator="in" result="shadow"/><feMerge><feMergeNode in="shadow"/><feMergeNode in="SourceGraphic"/></feMerge></filter>`);
      shape = `<g filter="url(#${id})">${shape}</g>`;
      context.warnings.push({ code: "shape-shadow-approximation", message: "SVG shadow blur and text coverage can differ from PowerPoint.", slideId: context.slideId, elementId: element.id });
    }
    let text = "";
    if (element.text !== undefined) {
      const textBox = resolveShapeTextBox(element);
      if (element.shape === "parallelogram") {
        context.warnings.push({ code: "shape-text-frame-approximation", message: "SVG uses the preset parallelogram text rectangle; text layout and wrapping can differ from PowerPoint.", slideId: context.slideId, elementId: element.id });
      }
      const textContext = textRenderContext(element, context);
      const rotated = rotatedText(element.text.content, element.text.frame, textBox, textContext);
      if (rotated !== undefined) text = rotated;
      else {
        const nativeTab = canRenderShapeTabTextNatively(element.text.content, element.text.frame);
        const mixed = !nativeTab && nativeTextOptIn(element.text.content, element.text.frame, context);
        if (nativeTab || mixed) {
          if (mixed) warnNativeMixedText(element, context);
          text = textSvg(element.text.content, element.text.frame, textBox, textContext);
        } else text = foreignObject(textBox, textHtml(element.text.content, element.text.frame, textContext), "", 1, "", true);
      }
    }
    return actionWrapper(element, `<g${transform} opacity="${element.opacity}"${accessibility(element)}>${shape}${text}</g>`, context);
  }
  if (element.type === "connector") {
    const points = [element.start, ...element.route, element.end];
    const start = marker(element.style.beginArrow, "start", element.style, context);
    const end = marker(element.style.endArrow, "end", element.style, context);
    const markers = `${start === undefined ? "" : ` marker-start="url(#${start})"`}${end === undefined ? "" : ` marker-end="url(#${end})"`}`;
    const path = element.pathMode === "closed-round" ? `<polygon points="${points.slice(0, -1).map((point) => `${point.x},${point.y}`).join(" ")}" fill="none" ${strokeAttributes(element.style, context.theme)}${markers}/>` : `<polyline points="${points.map((point) => `${point.x},${point.y}`).join(" ")}" fill="none" ${strokeAttributes(element.style, context.theme)}${markers}/>`;
    return actionWrapper(element, `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${accessibility(element)}>${path}</g>`, context);
  }
  if (element.type === "chart") {
    return actionWrapper(element, chartElement(element, context), context);
  }
  if (element.type === "table") {
    const nativeFlags = element.nativeTableStyleFlags;
    if (element.nativeFirstRowStyle === false || nativeFlags?.bandedRows === false ||
        nativeFlags?.firstColumn === true || nativeFlags?.lastColumn === true ||
        nativeFlags?.lastRow === true || nativeFlags?.bandedColumns === true) {
      context.warnings.push({
        code: "native-table-style-flags-not-rendered", message: "SVG uses effective Core cell formatting and does not apply the selected native table style flags.",
        slideId: context.slideId, elementId: element.id,
      });
    }
    if (element.nativeTableStylePreset === "medium-2-accent-2") context.warnings.push({
      code: "native-table-preset-not-rendered", message: "SVG uses effective Core cell formatting and does not render the native table preset.",
      slideId: context.slideId, elementId: element.id,
    });
    if (context.tableMode === "native") {
      const rendered = nativeTableSvg(element, context);
      if (rendered !== undefined) return actionWrapper(element, rendered, context);
      context.warnings.push({ code: "svg-native-table-fallback", message: "The table uses unsupported text direction, actions, tabs, bidirectional text, invalid geometry or exceeds native table grid limits; XHTML was retained.", slideId: context.slideId, elementId: element.id });
    }
    const frameAccessibility = accessibility(element, "child");
    if (element.borderGrid) {
      const borders = tableBorderSvg(element, context);
      return actionWrapper(element, `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${frameAccessibility}>${foreignObject(element.box, tableHtml(element, context), "", 1, "")}${borders}</g>`, context);
    }
    return actionWrapper(element, foreignObject(element.box, tableHtml(element, context), transformAttribute(element.box, element.transform), element.opacity, frameAccessibility), context);
  }
  return "";
}

function linearGradientDefinition(box: Box, paint: Extract<NormalizedBackgroundPaint, { type: "linearGradient" }>, theme: NormalizedPresentationTheme, id: string): string {
  const radians = paint.angle * Math.PI / 180;
  let dx = Math.cos(radians), dy = Math.sin(radians);
  if (paint.scaled && box.width > 0 && box.height > 0) { dx /= box.width; dy /= box.height; }
  const length = Math.hypot(dx, dy); dx /= length; dy /= length;
  const half = (Math.abs(box.width * dx) + Math.abs(box.height * dy)) / 2, cx = box.x + box.width / 2, cy = box.y + box.height / 2;
  const stops = paint.stops.map(stop => `<stop offset="${stop.position}" stop-color="${colorValue(stop.color, theme)}" stop-opacity="${stop.opacity}"/>`).join("");
  return `<linearGradient id="${id}" gradientUnits="userSpaceOnUse" x1="${cx - dx * half}" y1="${cy - dy * half}" x2="${cx + dx * half}" y2="${cy + dy * half}">${stops}</linearGradient>`;
}

export function backgroundRect(box: Box, paint: NormalizedBackgroundPaint, theme: NormalizedPresentationTheme, id: string): string {
  if (paint.type === "linearGradient") return `<defs>${linearGradientDefinition(box, paint, theme, id)}</defs><rect x="0" y="0" width="${box.width}" height="${box.height}" fill="url(#${id})"/>`;
  return `<rect x="0" y="0" width="${box.width}" height="${box.height}" ${paintAttributes(paint, theme, "fill")}/>`;
}
