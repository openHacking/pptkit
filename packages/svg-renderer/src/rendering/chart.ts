import type { NormalizedChartMarkerStyle, NormalizedChartTextStyle, NormalizedPaint, Point } from "@pptkit/core";
import type { LayoutChartElement } from "@pptkit/layout";
import { accessibility, nextDefinitionId, recordFont, type RenderContext } from "./context.js";
import { cssFontFamily, escapeXml } from "./escape.js";
import { colorValue, paintAttributes, strokeAttributes, transformAttribute } from "./style.js";

function n(value: number): number { return Number.isFinite(value * 1000) ? Math.round(value * 1000) / 1000 : value; }
function axisLabel(value: number, majorUnit: number): number { return Math.abs(majorUnit) < 0.001 || (value !== 0 && Math.abs(value) < 0.001) ? value : n(value); }
function percentLabel(value: number, majorUnit: number): string {
  const scaled = value * 100;
  if (Number.isFinite(scaled)) return `${axisLabel(scaled, majorUnit * 100)}%`;
  const [mantissa, exponent] = value.toExponential().split("e");
  return `${mantissa}e+${Number(exponent) + 2}%`;
}
function axisRatio(value: number, min: number, max: number): number {
  const span = max - min;
  return Number.isFinite(span) ? (value - min) / span : (value / 2 - min / 2) / (max / 2 - min / 2);
}
const invalidXmlCharacters = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/gu;
function chartText(value: string): string {
  return escapeXml(value.replace(invalidXmlCharacters, "\ufffd")).replace(/\r/g, "&#13;");
}

function textStyle(style: NormalizedChartTextStyle, context: RenderContext, text: string | number, elementId: string): string {
  recordFont(context, style.fontFamily, style.bold, style.italic, String(text), n(style.fontSize), undefined, elementId);
  return `font-family="${cssFontFamily(style.fontFamily)}" font-size="${n(style.fontSize)}" font-weight="${style.bold ? "700" : "400"}" font-style="${style.italic ? "italic" : "normal"}" fill="${style.color}"`;
}

function paintRect(box: { x: number; y: number; width: number; height: number }, paint: NormalizedPaint, context: RenderContext, role: string): string {
  return `<rect x="${n(box.x)}" y="${n(box.y)}" width="${n(box.width)}" height="${n(box.height)}" ${paintAttributes(paint, context.theme, "fill")} data-chart-role="${role}"/>`;
}

function markerShape(marker: NormalizedChartMarkerStyle, point: Point, context: RenderContext, role: string): string {
  const half = marker.size / 2;
  const fill = paintAttributes(marker.fill, context.theme, "fill");
  const stroke = strokeAttributes(marker.stroke, context.theme);
  const attrs = `${fill} ${stroke} data-chart-role="${role}"`;
  if (marker.shape === "circle") return `<circle cx="${n(point.x)}" cy="${n(point.y)}" r="${n(half)}" ${attrs}/>`;
  if (marker.shape === "diamond") return `<path d="M${n(point.x)} ${n(point.y - half)} L${n(point.x + half)} ${n(point.y)} L${n(point.x)} ${n(point.y + half)} L${n(point.x - half)} ${n(point.y)} Z" ${attrs}/>`;
  if (marker.shape === "triangle") return `<path d="M${n(point.x)} ${n(point.y - half)} L${n(point.x + half)} ${n(point.y + half)} L${n(point.x - half)} ${n(point.y + half)} Z" ${attrs}/>`;
  if (marker.shape === "star") {
    const points: string[] = [];
    for (let index = 0; index < 10; index += 1) {
      const radius = index % 2 === 0 ? half : half * 0.45;
      const angle = -Math.PI / 2 + index * Math.PI / 5;
      points.push(`${n(point.x + Math.cos(angle) * radius)},${n(point.y + Math.sin(angle) * radius)}`);
    }
    return `<polygon points="${points.join(" ")}" ${attrs}/>`;
  }
  if (marker.shape === "x" || marker.shape === "plus") {
    const d = marker.shape === "x"
      ? `M${n(point.x - half)} ${n(point.y - half)} L${n(point.x + half)} ${n(point.y + half)} M${n(point.x + half)} ${n(point.y - half)} L${n(point.x - half)} ${n(point.y + half)}`
      : `M${n(point.x - half)} ${n(point.y)} L${n(point.x + half)} ${n(point.y)} M${n(point.x)} ${n(point.y - half)} L${n(point.x)} ${n(point.y + half)}`;
    return `<path d="${d}" fill="none" ${stroke} data-chart-role="${role}"/>`;
  }
  if (marker.shape === "dash") return `<line x1="${n(point.x - half)}" y1="${n(point.y)}" x2="${n(point.x + half)}" y2="${n(point.y)}" ${stroke} data-chart-role="${role}"/>`;
  return `<rect x="${n(point.x - half)}" y="${n(point.y - half)}" width="${n(marker.size)}" height="${n(marker.size)}" ${attrs}/>`;
}

function piePath(center: Point, radius: number, start: number, end: number): string {
  const span = end - start;
  if (span >= Math.PI * 2 - 1e-6) return `M${n(center.x)} ${n(center.y - radius)} A${n(radius)} ${n(radius)} 0 1 1 ${n(center.x - 0.001)} ${n(center.y - radius)} Z`;
  const x1 = center.x + radius * Math.cos(start);
  const y1 = center.y + radius * Math.sin(start);
  const x2 = center.x + radius * Math.cos(end);
  const y2 = center.y + radius * Math.sin(end);
  return `M${n(center.x)} ${n(center.y)} L${n(x1)} ${n(y1)} A${n(radius)} ${n(radius)} 0 ${span > Math.PI ? 1 : 0} 1 ${n(x2)} ${n(y2)} Z`;
}

function ringPath(center: Point, inner: number, outer: number, start: number, end: number): string {
  const point = (radius: number, angle: number) => `${n(center.x + radius * Math.cos(angle))} ${n(center.y + radius * Math.sin(angle))}`;
  const span = end - start;
  if (span >= Math.PI * 2 - 1e-6) {
    const middle = start + Math.PI;
    return `M${point(outer, start)} A${n(outer)} ${n(outer)} 0 1 1 ${point(outer, middle)} A${n(outer)} ${n(outer)} 0 1 1 ${point(outer, start)} L${point(inner, start)} A${n(inner)} ${n(inner)} 0 1 0 ${point(inner, middle)} A${n(inner)} ${n(inner)} 0 1 0 ${point(inner, start)} Z`;
  }
  const large = span > Math.PI ? 1 : 0;
  return `M${point(outer, start)} A${n(outer)} ${n(outer)} 0 ${large} 1 ${point(outer, end)} L${point(inner, end)} A${n(inner)} ${n(inner)} 0 ${large} 0 ${point(inner, start)} Z`;
}

function tickLength(mark: "none" | "inside" | "outside" | "cross"): { before: number; after: number } {
  if (mark === "none") return { before: 0, after: 0 };
  if (mark === "inside") return { before: 3, after: 0 };
  if (mark === "cross") return { before: 3, after: 3 };
  return { before: 0, after: 3 };
}

function axesSvg(element: Exclude<LayoutChartElement, { chartType: "pie" | "doughnut" | "scatter" | "bubble" }>, context: RenderContext): string {
  const { plotBox: plot, valueScale: scale, categoryPositions, categoryTickPositions } = element.chartLayout;
  const horizontal = element.chartType === "bar" && element.orientation === "horizontal";
  const pieces: string[] = [];
  if (element.axes.category.visible) {
    pieces.push(horizontal
      ? `<line x1="${n(plot.x)}" y1="${n(plot.y)}" x2="${n(plot.x)}" y2="${n(plot.y + plot.height)}" ${strokeAttributes(element.axes.category.line, context.theme)} data-chart-role="category-axis"/>`
      : `<line x1="${n(plot.x)}" y1="${n(plot.y + plot.height)}" x2="${n(plot.x + plot.width)}" y2="${n(plot.y + plot.height)}" ${strokeAttributes(element.axes.category.line, context.theme)} data-chart-role="category-axis"/>`);
    const length = tickLength(element.axes.category.majorTick);
    for (const position of categoryTickPositions) {
      pieces.push(horizontal
        ? `<line x1="${n(plot.x - length.after)}" y1="${n(position.y)}" x2="${n(plot.x + length.before)}" y2="${n(position.y)}" ${strokeAttributes(element.axes.category.line, context.theme)} data-chart-role="category-tick"/>`
        : `<line x1="${n(position.x)}" y1="${n(plot.y + plot.height - length.before)}" x2="${n(position.x)}" y2="${n(plot.y + plot.height + length.after)}" ${strokeAttributes(element.axes.category.line, context.theme)} data-chart-role="category-tick"/>`);
    }
    for (const [index, position] of categoryPositions.entries()) {
      if (element.axes.category.labels) pieces.push(horizontal
        ? `<text x="${n(plot.x - length.after - 4)}" y="${n(position.y + element.axes.category.labelStyle.fontSize * 0.35)}" ${textStyle(element.axes.category.labelStyle, context, element.categories[index] ?? "", element.id)} text-anchor="end" data-chart-role="category-label">${chartText(element.categories[index] ?? "")}</text>`
        : `<text x="${n(position.x)}" y="${n(plot.y + plot.height + length.after + element.axes.category.labelStyle.fontSize + 3)}" ${textStyle(element.axes.category.labelStyle, context, element.categories[index] ?? "", element.id)} text-anchor="middle" data-chart-role="category-label">${chartText(element.categories[index] ?? "")}</text>`);
    }
  }
  if (element.axes.value.visible) {
    pieces.push(horizontal
      ? `<line x1="${n(plot.x)}" y1="${n(plot.y + plot.height)}" x2="${n(plot.x + plot.width)}" y2="${n(plot.y + plot.height)}" ${strokeAttributes(element.axes.value.line, context.theme)} data-chart-role="value-axis"/>`
      : `<line x1="${n(plot.x)}" y1="${n(plot.y)}" x2="${n(plot.x)}" y2="${n(plot.y + plot.height)}" ${strokeAttributes(element.axes.value.line, context.theme)} data-chart-role="value-axis"/>`);
    const length = tickLength(element.axes.value.majorTick);
    for (const value of scale.ticks) {
      const ratio = axisRatio(value, scale.min, scale.max);
      const point = horizontal ? { x: plot.x + ratio * plot.width, y: plot.y + plot.height } : { x: plot.x, y: plot.y + plot.height - ratio * plot.height };
      pieces.push(horizontal
        ? `<line x1="${n(point.x)}" y1="${n(point.y - length.before)}" x2="${n(point.x)}" y2="${n(point.y + length.after)}" ${strokeAttributes(element.axes.value.line, context.theme)} data-chart-role="value-tick"/>`
        : `<line x1="${n(point.x - length.after)}" y1="${n(point.y)}" x2="${n(point.x + length.before)}" y2="${n(point.y)}" ${strokeAttributes(element.axes.value.line, context.theme)} data-chart-role="value-tick"/>`);
      const label = element.chartType === "area" && element.grouping === "percentStacked" ? percentLabel(value, scale.majorUnit) : axisLabel(value, scale.majorUnit);
      if (element.axes.value.labels) pieces.push(horizontal
        ? `<text x="${n(point.x)}" y="${n(point.y + length.after + element.axes.value.labelStyle.fontSize + 3)}" ${textStyle(element.axes.value.labelStyle, context, label, element.id)} text-anchor="middle" data-chart-role="value-label">${label}</text>`
        : `<text x="${n(point.x - length.after - 4)}" y="${n(point.y + element.axes.value.labelStyle.fontSize * 0.35)}" ${textStyle(element.axes.value.labelStyle, context, label, element.id)} text-anchor="end" data-chart-role="value-label">${label}</text>`);
    }
  }
  if (element.chartType === "combo" && element.axes.secondaryValue.visible && element.chartLayout.secondaryValueScale) {
    const axis = element.axes.secondaryValue, secondary = element.chartLayout.secondaryValueScale;
    const x = plot.x + plot.width, length = tickLength(axis.majorTick);
    pieces.push(`<line x1="${n(x)}" y1="${n(plot.y)}" x2="${n(x)}" y2="${n(plot.y + plot.height)}" ${strokeAttributes(axis.line, context.theme)} data-chart-role="secondary-axis"/>`);
    for (const value of secondary.ticks) {
      const y = plot.y + plot.height * (1 - axisRatio(value, secondary.min, secondary.max));
      pieces.push(`<line x1="${n(x - length.before)}" y1="${n(y)}" x2="${n(x + length.after)}" y2="${n(y)}" ${strokeAttributes(axis.line, context.theme)} data-chart-role="secondary-tick"/>`);
      if (axis.labels) pieces.push(`<text x="${n(x + length.after + 4)}" y="${n(y + axis.labelStyle.fontSize * .35)}" ${textStyle(axis.labelStyle, context, axisLabel(value, secondary.majorUnit), element.id)} text-anchor="start" data-chart-role="secondary-label">${axisLabel(value, secondary.majorUnit)}</text>`);
    }
  }
  return pieces.join("");
}

function legendSvg(element: LayoutChartElement, context: RenderContext): string {
  const box = element.chartLayout.legendBox;
  if (!element.legend.visible || box === undefined) return "";
  const items = element.chartLayout.legendItems;
  return items.map((item) => {
    const x = item.box.x + 2;
    const y = item.box.y + (element.legend.position === "right" ? 0 : 3);
    let glyph: string;
    const series = element.series[item.seriesIndex ?? 0];
    if (series && "line" in series) {
      const line = series.line;
      glyph = `${line === false || line === undefined ? "" : `<line x1="${n(x)}" y1="${n(y + 5)}" x2="${n(x + 16)}" y2="${n(y + 5)}" ${strokeAttributes(line, context.theme)} data-chart-role="legend-line"/>`}${item.marker === false ? "" : markerShape(item.marker, { x: x + 8, y: y + 5 }, context, "legend-marker")}`;
    } else if (element.chartType === "bubble") glyph = `<circle cx="${n(x + 7)}" cy="${n(y + 5)}" r="4" fill="${colorValue(item.color, context.theme)}" data-chart-role="legend-swatch"/>`;
    else glyph = `<rect x="${n(x + 3)}" y="${n(y + 1)}" width="8" height="8" fill="${colorValue(item.color, context.theme)}" data-chart-role="legend-swatch"/>`;
    return `${glyph}<text x="${n(x + 20)}" y="${n(y + 9)}" ${textStyle(element.legend.textStyle, context, item.label, element.id)} text-anchor="start" data-chart-role="legend-label">${chartText(item.label)}</text>`;
  }).join("");
}

function xyChartElement(element: Extract<LayoutChartElement, { chartType: "scatter" | "bubble" }>, context: RenderContext): string {
  const layout = element.chartLayout, plot = layout.plotBox, xs = layout.xScale!, ys = layout.valueScale;
  const text = [...(element.title === undefined ? [] : [element.title]), ...layout.legendItems.map(item => item.label)];
  if (text.some(value => { invalidXmlCharacters.lastIndex = 0; return invalidXmlCharacters.test(value); })) context.warnings.push({ code: "chart-text-xml-replacement", slideId: context.slideId, elementId: element.id, message: "Characters forbidden by XML 1.0 were replaced with U+FFFD in the SVG chart preview." });
  const clipId = nextDefinitionId(context);context.defs.push(`<clipPath id="${clipId}"><rect x="${n(plot.x)}" y="${n(plot.y)}" width="${n(plot.width)}" height="${n(plot.height)}"/></clipPath>`);
  const background = paintRect({ x: 0, y: 0, width: element.box.width, height: element.box.height }, element.style.chartArea, context, "chart-area") + paintRect(plot, element.style.plotArea, context, "plot-area");
  const xAxisY = plot.y + plot.height - Math.max(0, Math.min(1, axisRatio(0, ys.min, ys.max))) * plot.height;
  const yAxisX = plot.x + Math.max(0, Math.min(1, axisRatio(0, xs.min, xs.max))) * plot.width;
  const grid: string[] = [], axes: string[] = [], data: string[] = [];
  for (const [name, scale, axis] of [["x", xs, element.axes.x], ["y", ys, element.axes.y]] as const) {
    const horizontal = name === "x";
    if (axis.visible) axes.push(horizontal ? `<line x1="${n(plot.x)}" y1="${n(xAxisY)}" x2="${n(plot.x+plot.width)}" y2="${n(xAxisY)}" ${strokeAttributes(axis.line,context.theme)} data-chart-role="x-axis"/>` : `<line x1="${n(yAxisX)}" y1="${n(plot.y)}" x2="${n(yAxisX)}" y2="${n(plot.y+plot.height)}" ${strokeAttributes(axis.line,context.theme)} data-chart-role="y-axis"/>`);
    for (const value of scale.ticks) {
      const ratio = axisRatio(value,scale.min,scale.max), x = horizontal ? plot.x+ratio*plot.width : yAxisX, y = horizontal ? xAxisY : plot.y+plot.height-ratio*plot.height;
      if (axis.majorGridlines !== false) grid.push(horizontal ? `<line x1="${n(x)}" y1="${n(plot.y)}" x2="${n(x)}" y2="${n(plot.y+plot.height)}" ${strokeAttributes(axis.majorGridlines,context.theme)}/>` : `<line x1="${n(plot.x)}" y1="${n(y)}" x2="${n(plot.x+plot.width)}" y2="${n(y)}" ${strokeAttributes(axis.majorGridlines,context.theme)}/>`);
      const length = tickLength(axis.majorTick);
      if (axis.visible) axes.push(horizontal ? `<line x1="${n(x)}" y1="${n(y-length.before)}" x2="${n(x)}" y2="${n(y+length.after)}" ${strokeAttributes(axis.line,context.theme)} data-chart-role="x-tick"/>` : `<line x1="${n(x-length.after)}" y1="${n(y)}" x2="${n(x+length.before)}" y2="${n(y)}" ${strokeAttributes(axis.line,context.theme)} data-chart-role="y-tick"/>`);
      if (axis.visible && axis.labels) axes.push(horizontal ? `<text x="${n(x)}" y="${n(y+length.after+axis.labelStyle.fontSize+4)}" ${textStyle(axis.labelStyle, context, axisLabel(value,scale.majorUnit), element.id)} text-anchor="middle" data-chart-role="x-label">${axisLabel(value,scale.majorUnit)}</text>` : `<text x="${n(x-length.after-4)}" y="${n(y+axis.labelStyle.fontSize*.35)}" ${textStyle(axis.labelStyle, context, axisLabel(value,scale.majorUnit), element.id)} text-anchor="end" data-chart-role="y-label">${axisLabel(value,scale.majorUnit)}</text>`);
    }
  }
  if (element.chartType === "bubble") {
    for (const item of layout.bubbles ?? []) {
      const series = element.series[item.seriesIndex];
      if (!series) continue;
      for (const point of item.points) data.push(`<circle cx="${n(point.x)}" cy="${n(point.y)}" r="${n(point.radius)}" fill="${colorValue(series.color, context.theme)}" data-chart-role="bubble"/>`);
    }
  } else for (const line of layout.scatter ?? []) {
    const series = element.series[line.seriesIndex];if (!series) continue;
    if (series.line !== false) for (const segment of line.segments) if (segment.length > 1) data.push(`<polyline points="${segment.map(p=>`${n(p.x)},${n(p.y)}`).join(" ")}" fill="none" ${strokeAttributes(series.line,context.theme)} data-chart-role="scatter-line"/>`);
    if (series.marker !== false) for (const point of line.points) data.push(markerShape(series.marker,point,context,"scatter-marker"));
  }
  const title = element.title === undefined || layout.titleBox === undefined ? "" : `<text x="${n(layout.titleBox.x+layout.titleBox.width/2)}" y="${n(layout.titleBox.y+element.titleStyle.fontSize+2)}" ${textStyle(element.titleStyle, context, element.title, element.id)} text-anchor="middle" data-chart-role="title">${chartText(element.title)}</text>`;
  return `<g${transformAttribute(element.box,element.transform)} opacity="${element.opacity}"${accessibility(element)}><g transform="translate(${n(element.box.x)} ${n(element.box.y)})">${background}<g data-chart-layer="grid">${grid.join("")}</g><g clip-path="url(#${clipId})" data-chart-layer="series">${data.join("")}</g><g data-chart-layer="axes">${axes.join("")}</g>${title}<g data-chart-layer="legend">${legendSvg(element,context)}</g></g></g>`;
}

export function chartElement(element: LayoutChartElement, context: RenderContext): string {
  if (element.chartType === "scatter" || element.chartType === "bubble") return xyChartElement(element, context);
  const layout = element.chartLayout;
  const text = [
    ...(element.title !== undefined && layout.titleBox !== undefined ? [element.title] : []),
    ...(element.chartType !== "pie" && element.chartType !== "doughnut" && element.axes.category.visible && element.axes.category.labels ? element.categories : []),
    ...(element.legend.visible && layout.legendBox !== undefined ? layout.legendItems.map(item => item.label) : []),
  ];
  if (text.some(value => { invalidXmlCharacters.lastIndex = 0; return invalidXmlCharacters.test(value); })) context.warnings.push({ code: "chart-text-xml-replacement", slideId: context.slideId, elementId: element.id, message: "Characters forbidden by XML 1.0 were replaced with U+FFFD in the SVG chart preview." });
  const clipId = nextDefinitionId(context);
  context.defs.push(`<clipPath id="${clipId}"><rect x="${n(layout.plotBox.x)}" y="${n(layout.plotBox.y)}" width="${n(layout.plotBox.width)}" height="${n(layout.plotBox.height)}"/></clipPath>`);
  const background = paintRect({ x: 0, y: 0, width: element.box.width, height: element.box.height }, element.style.chartArea, context, "chart-area");
  const plotBackground = paintRect(layout.plotBox, element.style.plotArea, context, "plot-area");
  const grid: string[] = [];
  if (element.chartType !== "pie" && element.chartType !== "doughnut" && element.axes.category.majorGridlines !== false) {
    const horizontal = element.chartType === "bar" && element.orientation === "horizontal";
    for (const position of layout.categoryTickPositions) grid.push(horizontal
      ? `<line x1="${n(layout.plotBox.x)}" y1="${n(position.y)}" x2="${n(layout.plotBox.x + layout.plotBox.width)}" y2="${n(position.y)}" ${strokeAttributes(element.axes.category.majorGridlines, context.theme)} data-chart-role="category-gridline"/>`
      : `<line x1="${n(position.x)}" y1="${n(layout.plotBox.y)}" x2="${n(position.x)}" y2="${n(layout.plotBox.y + layout.plotBox.height)}" ${strokeAttributes(element.axes.category.majorGridlines, context.theme)} data-chart-role="category-gridline"/>`);
  }
  if (element.chartType !== "pie" && element.chartType !== "doughnut" && element.axes.value.majorGridlines !== false) {
    const horizontal = element.chartType === "bar" && element.orientation === "horizontal";
    for (const value of layout.valueScale.ticks) {
      const ratio = axisRatio(value, layout.valueScale.min, layout.valueScale.max);
      grid.push(horizontal
        ? `<line x1="${n(layout.plotBox.x + ratio * layout.plotBox.width)}" y1="${n(layout.plotBox.y)}" x2="${n(layout.plotBox.x + ratio * layout.plotBox.width)}" y2="${n(layout.plotBox.y + layout.plotBox.height)}" ${strokeAttributes(element.axes.value.majorGridlines, context.theme)} data-chart-role="gridline"/>`
        : `<line x1="${n(layout.plotBox.x)}" y1="${n(layout.plotBox.y + layout.plotBox.height - ratio * layout.plotBox.height)}" x2="${n(layout.plotBox.x + layout.plotBox.width)}" y2="${n(layout.plotBox.y + layout.plotBox.height - ratio * layout.plotBox.height)}" ${strokeAttributes(element.axes.value.majorGridlines, context.theme)} data-chart-role="gridline"/>`);
    }
  }
  if (element.chartType === "combo" && layout.secondaryValueScale && element.axes.secondaryValue.majorGridlines !== false) {
    const scale = layout.secondaryValueScale, plot = layout.plotBox;
    for (const value of scale.ticks) {
      const y = plot.y + plot.height * (1 - axisRatio(value, scale.min, scale.max));
      grid.push(`<line x1="${n(plot.x)}" y1="${n(y)}" x2="${n(plot.x + plot.width)}" y2="${n(y)}" ${strokeAttributes(element.axes.secondaryValue.majorGridlines, context.theme)} data-chart-role="secondary-gridline"/>`);
    }
  }
  const data: string[] = [];
  if (element.chartType === "area") for (const area of layout.areas ?? []) for (const polygon of area.polygons) if (polygon.length >= 4) data.push(`<polygon points="${polygon.map(point => `${n(point.x)},${n(point.y)}`).join(" ")}" fill="${colorValue(element.series[area.seriesIndex]?.color ?? "#000000", context.theme)}" data-chart-role="area"/>`);
  if (element.chartType === "bar" || element.chartType === "combo") for (const bar of layout.bars) data.push(`<rect x="${n(bar.box.x)}" y="${n(bar.box.y)}" width="${n(bar.box.width)}" height="${n(bar.box.height)}" fill="${colorValue(element.series[bar.seriesIndex]?.color ?? "#000000", context.theme)}" data-chart-role="bar"/>`);
  if (element.chartType === "line" || element.chartType === "combo") for (const line of layout.lines) {
    const series = element.series[line.seriesIndex];
    if (!series || !("line" in series)) continue;
    if (series.line !== false) for (const segment of line.segments ?? [line.points]) data.push(`<polyline points="${segment.map((point) => `${n(point.x)},${n(point.y)}`).join(" ")}" fill="none" ${strokeAttributes(series.line, context.theme)} data-chart-role="line"/>`);
    if (series.marker !== false) for (const point of line.points) data.push(markerShape(series.marker, point, context, "series-marker"));
  } else if (layout.pie !== undefined) for (const slice of layout.pie.slices) if (slice.endAngle - slice.startAngle > 1e-9) data.push(`<path d="${piePath(layout.pie.center, layout.pie.radius, slice.startAngle, slice.endAngle)}" fill="${colorValue(slice.color, context.theme)}" data-chart-role="pie-slice"/>`);
  if (element.chartType === "doughnut" && layout.doughnut !== undefined) for (const ring of layout.doughnut.rings) for (const slice of ring.slices) if (slice.endAngle - slice.startAngle > 1e-9) data.push(`<path d="${ringPath(layout.doughnut.center, ring.innerRadius, ring.outerRadius, slice.startAngle, slice.endAngle)}" fill="${colorValue(slice.color, context.theme)}" data-chart-role="doughnut-slice"/>`);
  const axes = element.chartType === "pie" || element.chartType === "doughnut" ? "" : axesSvg(element, context);
  const title = element.title === undefined || layout.titleBox === undefined ? "" : `<text x="${n(layout.titleBox.x + layout.titleBox.width / 2)}" y="${n(layout.titleBox.y + element.titleStyle.fontSize + 2)}" ${textStyle(element.titleStyle, context, element.title, element.id)} text-anchor="middle" data-chart-role="title">${chartText(element.title)}</text>`;
  const local = `${background}${plotBackground}<g data-chart-layer="grid">${grid.join("")}</g><g clip-path="url(#${clipId})" data-chart-layer="series">${data.join("")}</g><g data-chart-layer="axes">${axes}</g>${title}<g data-chart-layer="legend">${legendSvg(element, context)}</g>`;
  return `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${accessibility(element)}><g transform="translate(${n(element.box.x)} ${n(element.box.y)})">${local}</g></g>`;
}
