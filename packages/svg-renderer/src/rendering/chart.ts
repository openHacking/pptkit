import type { NormalizedChartMarkerStyle, NormalizedChartTextStyle, NormalizedPaint, Point } from "@pptkit/core";
import type { LayoutChartElement } from "@pptkit/layout";
import { accessibility, type RenderContext } from "./context.js";
import { escapeXml } from "./escape.js";
import { colorValue, paintAttributes, strokeAttributes, transformAttribute } from "./style.js";

function n(value: number): number { return Math.round(value * 1000) / 1000; }

function textStyle(style: NormalizedChartTextStyle): string {
  return `font-family="${escapeXml(style.fontFamily)}" font-size="${n(style.fontSize)}" font-weight="${style.bold ? "700" : "400"}" font-style="${style.italic ? "italic" : "normal"}" fill="${style.color}"`;
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

function tickLength(mark: "none" | "inside" | "outside" | "cross"): { before: number; after: number } {
  if (mark === "none") return { before: 0, after: 0 };
  if (mark === "inside") return { before: 3, after: 0 };
  if (mark === "cross") return { before: 3, after: 3 };
  return { before: 0, after: 3 };
}

function axesSvg(element: Exclude<LayoutChartElement, { chartType: "pie" }>, context: RenderContext): string {
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
        ? `<text x="${n(plot.x - length.after - 4)}" y="${n(position.y + element.axes.category.labelStyle.fontSize * 0.35)}" ${textStyle(element.axes.category.labelStyle)} text-anchor="end" data-chart-role="category-label">${escapeXml(element.categories[index] ?? "")}</text>`
        : `<text x="${n(position.x)}" y="${n(plot.y + plot.height + length.after + element.axes.category.labelStyle.fontSize + 3)}" ${textStyle(element.axes.category.labelStyle)} text-anchor="middle" data-chart-role="category-label">${escapeXml(element.categories[index] ?? "")}</text>`);
    }
  }
  if (element.axes.value.visible) {
    pieces.push(horizontal
      ? `<line x1="${n(plot.x)}" y1="${n(plot.y + plot.height)}" x2="${n(plot.x + plot.width)}" y2="${n(plot.y + plot.height)}" ${strokeAttributes(element.axes.value.line, context.theme)} data-chart-role="value-axis"/>`
      : `<line x1="${n(plot.x)}" y1="${n(plot.y)}" x2="${n(plot.x)}" y2="${n(plot.y + plot.height)}" ${strokeAttributes(element.axes.value.line, context.theme)} data-chart-role="value-axis"/>`);
    const length = tickLength(element.axes.value.majorTick);
    for (const value of scale.ticks) {
      const ratio = (value - scale.min) / (scale.max - scale.min);
      const point = horizontal ? { x: plot.x + ratio * plot.width, y: plot.y + plot.height } : { x: plot.x, y: plot.y + plot.height - ratio * plot.height };
      pieces.push(horizontal
        ? `<line x1="${n(point.x)}" y1="${n(point.y - length.before)}" x2="${n(point.x)}" y2="${n(point.y + length.after)}" ${strokeAttributes(element.axes.value.line, context.theme)} data-chart-role="value-tick"/>`
        : `<line x1="${n(point.x - length.after)}" y1="${n(point.y)}" x2="${n(point.x + length.before)}" y2="${n(point.y)}" ${strokeAttributes(element.axes.value.line, context.theme)} data-chart-role="value-tick"/>`);
      if (element.axes.value.labels) pieces.push(horizontal
        ? `<text x="${n(point.x)}" y="${n(point.y + length.after + element.axes.value.labelStyle.fontSize + 3)}" ${textStyle(element.axes.value.labelStyle)} text-anchor="middle" data-chart-role="value-label">${n(value)}</text>`
        : `<text x="${n(point.x - length.after - 4)}" y="${n(point.y + element.axes.value.labelStyle.fontSize * 0.35)}" ${textStyle(element.axes.value.labelStyle)} text-anchor="end" data-chart-role="value-label">${n(value)}</text>`);
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
    if (element.chartType === "line" && item.marker !== false) {
      const line = element.series[item.seriesIndex ?? 0]?.line;
      glyph = `${line === false || line === undefined ? "" : `<line x1="${n(x)}" y1="${n(y + 5)}" x2="${n(x + 16)}" y2="${n(y + 5)}" ${strokeAttributes(line, context.theme)} data-chart-role="legend-line"/>`}${markerShape(item.marker, { x: x + 8, y: y + 5 }, context, "legend-marker")}`;
    } else glyph = `<rect x="${n(x + 3)}" y="${n(y + 1)}" width="8" height="8" fill="${colorValue(item.color, context.theme)}" data-chart-role="legend-swatch"/>`;
    return `${glyph}<text x="${n(x + 20)}" y="${n(y + 9)}" ${textStyle(element.legend.textStyle)} text-anchor="start" data-chart-role="legend-label">${escapeXml(item.label)}</text>`;
  }).join("");
}

export function chartElement(element: LayoutChartElement, context: RenderContext): string {
  const layout = element.chartLayout;
  const clipId = `pptkit-chart-${element.id.replace(/[^A-Za-z0-9_.-]/g, "-")}`;
  context.defs.push(`<clipPath id="${clipId}"><rect x="${n(layout.plotBox.x)}" y="${n(layout.plotBox.y)}" width="${n(layout.plotBox.width)}" height="${n(layout.plotBox.height)}"/></clipPath>`);
  const background = paintRect({ x: 0, y: 0, width: element.box.width, height: element.box.height }, element.style.chartArea, context, "chart-area");
  const plotBackground = paintRect(layout.plotBox, element.style.plotArea, context, "plot-area");
  const grid: string[] = [];
  if (element.chartType !== "pie" && element.axes.category.majorGridlines !== false) {
    const horizontal = element.chartType === "bar" && element.orientation === "horizontal";
    for (const position of layout.categoryTickPositions) grid.push(horizontal
      ? `<line x1="${n(layout.plotBox.x)}" y1="${n(position.y)}" x2="${n(layout.plotBox.x + layout.plotBox.width)}" y2="${n(position.y)}" ${strokeAttributes(element.axes.category.majorGridlines, context.theme)} data-chart-role="category-gridline"/>`
      : `<line x1="${n(position.x)}" y1="${n(layout.plotBox.y)}" x2="${n(position.x)}" y2="${n(layout.plotBox.y + layout.plotBox.height)}" ${strokeAttributes(element.axes.category.majorGridlines, context.theme)} data-chart-role="category-gridline"/>`);
  }
  if (element.chartType !== "pie" && element.axes.value.majorGridlines !== false) {
    const horizontal = element.chartType === "bar" && element.orientation === "horizontal";
    for (const value of layout.valueScale.ticks) {
      const ratio = (value - layout.valueScale.min) / (layout.valueScale.max - layout.valueScale.min);
      grid.push(horizontal
        ? `<line x1="${n(layout.plotBox.x + ratio * layout.plotBox.width)}" y1="${n(layout.plotBox.y)}" x2="${n(layout.plotBox.x + ratio * layout.plotBox.width)}" y2="${n(layout.plotBox.y + layout.plotBox.height)}" ${strokeAttributes(element.axes.value.majorGridlines, context.theme)} data-chart-role="gridline"/>`
        : `<line x1="${n(layout.plotBox.x)}" y1="${n(layout.plotBox.y + layout.plotBox.height - ratio * layout.plotBox.height)}" x2="${n(layout.plotBox.x + layout.plotBox.width)}" y2="${n(layout.plotBox.y + layout.plotBox.height - ratio * layout.plotBox.height)}" ${strokeAttributes(element.axes.value.majorGridlines, context.theme)} data-chart-role="gridline"/>`);
    }
  }
  const data: string[] = [];
  if (element.chartType === "bar") for (const bar of layout.bars) data.push(`<rect x="${n(bar.box.x)}" y="${n(bar.box.y)}" width="${n(bar.box.width)}" height="${n(bar.box.height)}" fill="${colorValue(element.series[bar.seriesIndex]?.color ?? "#000000", context.theme)}" data-chart-role="bar"/>`);
  else if (element.chartType === "line") for (const line of layout.lines) {
    const series = element.series[line.seriesIndex];
    if (series === undefined || series.line === false) continue;
    data.push(`<polyline points="${line.points.map((point) => `${n(point.x)},${n(point.y)}`).join(" ")}" fill="none" ${strokeAttributes(series.line, context.theme)} data-chart-role="line"/>`);
    if (series.marker !== false) for (const point of line.points) data.push(markerShape(series.marker, point, context, "series-marker"));
  } else if (layout.pie !== undefined) for (const slice of layout.pie.slices) if (slice.endAngle - slice.startAngle > 1e-9) data.push(`<path d="${piePath(layout.pie.center, layout.pie.radius, slice.startAngle, slice.endAngle)}" fill="${colorValue(slice.color, context.theme)}" data-chart-role="pie-slice"/>`);
  const axes = element.chartType === "pie" ? "" : axesSvg(element, context);
  const title = element.title === undefined || layout.titleBox === undefined ? "" : `<text x="${n(layout.titleBox.x + layout.titleBox.width / 2)}" y="${n(layout.titleBox.y + element.titleStyle.fontSize + 2)}" ${textStyle(element.titleStyle)} text-anchor="middle" data-chart-role="title">${escapeXml(element.title)}</text>`;
  const local = `${background}${plotBackground}<g data-chart-layer="grid">${grid.join("")}</g><g clip-path="url(#${clipId})" data-chart-layer="series">${data.join("")}</g><g data-chart-layer="axes">${axes}</g>${title}<g data-chart-layer="legend">${legendSvg(element, context)}</g>`;
  return `<g${transformAttribute(element.box, element.transform)} opacity="${element.opacity}"${accessibility(element)}><g transform="translate(${n(element.box.x)} ${n(element.box.y)})">${local}</g></g>`;
}
