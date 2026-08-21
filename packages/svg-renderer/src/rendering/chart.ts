import type { NormalizedChartElement } from "@pptkit/core";
import type { LayoutElement } from "@pptkit/layout";
import { accessibility, type RenderContext } from "./context.js";
import { escapeXml } from "./escape.js";
import { colorValue, transformAttribute } from "./style.js";

function formatNumber(value: number): number {
  return Math.round(value * 1000) / 1000;
}

export function chartElement(
  element: Extract<LayoutElement, { type: "chart" }>,
  context: RenderContext,
): string {
  const { box } = element;
  const { categories, series, title, showLegend, xAxis, yAxis, chartType } = element;

  const padding = 8;
  const titleHeight = title ? 20 : 0;
  const legendHeight = showLegend ? 16 : 0;
  const leftMargin = yAxis.show ? 36 : padding;
  const bottomMargin = (xAxis.show && xAxis.labels ? 20 : 0) + legendHeight + padding;
  const topMargin = padding + titleHeight;
  const rightMargin = padding;

  const plotX = leftMargin;
  const plotY = topMargin;
  const plotW = Math.max(1, box.width - leftMargin - rightMargin);
  const plotH = Math.max(1, box.height - topMargin - bottomMargin);

  let minValue = Infinity;
  let maxValue = -Infinity;
  for (const s of series) {
    for (const v of s.values) {
      if (v < minValue) minValue = v;
      if (v > maxValue) maxValue = v;
    }
  }
  if (!isFinite(minValue)) minValue = 0;
  if (!isFinite(maxValue)) maxValue = 0;
  const baseline = Math.min(0, minValue);
  const range = Math.max(maxValue - minValue, 1e-6);

  function yForValue(v: number): number {
    return plotY + plotH - ((v - baseline) / range) * plotH;
  }

  const zeroY = yForValue(0);
  const parts: string[] = [];
  const legendParts: string[] = [];

  if (title) {
    parts.push(
      `<text x="${formatNumber(box.width / 2)}" y="${formatNumber(padding + 14)}" font-family="Arial" font-size="12" fill="#333333" text-anchor="middle">${escapeXml(title)}</text>`,
    );
  }

  if (chartType === "bar") {
    const slotWidth = plotW / Math.max(categories.length, 1);
    const groupGap = slotWidth * 0.2;
    const barWidth = (slotWidth - groupGap) / Math.max(series.length, 1);

    for (let c = 0; c < categories.length; c++) {
      for (const [i, s] of series.entries()) {
        const value = s.values[c] ?? 0;
        const x = plotX + c * slotWidth + groupGap / 2 + i * barWidth;
        const y = Math.min(yForValue(value), zeroY);
        const height = Math.abs(yForValue(value) - zeroY);
        const fill = colorValue(s.color, context.theme);
        parts.push(
          `<rect x="${formatNumber(x)}" y="${formatNumber(y)}" width="${formatNumber(barWidth)}" height="${formatNumber(height)}" fill="${fill}"/>`,
        );
      }
    }
  } else if (chartType === "line") {
    const stepX = plotW / Math.max(categories.length, 1);

    for (const s of series) {
      const points: string[] = [];
      for (let i = 0; i < categories.length; i++) {
        const x = plotX + (i + 0.5) * stepX;
        const y = yForValue(s.values[i] ?? 0);
        points.push(`${formatNumber(x)},${formatNumber(y)}`);
      }

      if (points.length > 1) {
        const stroke = colorValue(s.color, context.theme);
        parts.push(
          `<polyline points="${points.join(" ")}" fill="none" stroke="${stroke}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>`,
        );
      }

      for (let i = 0; i < categories.length; i++) {
        const x = plotX + (i + 0.5) * stepX;
        const y = yForValue(s.values[i] ?? 0);
        const fill = colorValue(s.color, context.theme);
        parts.push(
          `<rect x="${formatNumber(x - 3)}" y="${formatNumber(y - 3)}" width="6" height="6" fill="${fill}"/>`,
        );
      }
    }
  } else if (chartType === "pie") {
    const firstSeries = series[0];
    if (firstSeries) {
      const total = firstSeries.values.reduce((sum, v) => sum + v, 0);
      const cx = plotX + plotW / 2;
      const cy = plotY + plotH / 2;
      const radius = (Math.min(plotW, plotH) / 2) * 0.85;

      let startAngle = -Math.PI / 2;

      for (let i = 0; i < categories.length; i++) {
        const value = firstSeries.values[i] ?? 0;
        const sliceAngle = total > 0 ? (value / total) * 2 * Math.PI : 0;
        const endAngle = startAngle + sliceAngle;

        if (sliceAngle > 1e-9) {
          const x1 = formatNumber(cx + radius * Math.cos(startAngle));
          const y1 = formatNumber(cy + radius * Math.sin(startAngle));
          const x2 = formatNumber(cx + radius * Math.cos(endAngle));
          const y2 = formatNumber(cy + radius * Math.sin(endAngle));
          const largeArc = sliceAngle > Math.PI ? 1 : 0;

          const colorSource = series[i % series.length];
          if (colorSource !== undefined) {
            const fill = colorValue(colorSource.color, context.theme);
            parts.push(
              `<path d="M${formatNumber(cx)} ${formatNumber(cy)} L${x1} ${y1} A${formatNumber(radius)} ${formatNumber(radius)} 0 ${largeArc} 1 ${x2} ${y2} Z" fill="${fill}"/>`,
            );
          }
        }

        startAngle = endAngle;
      }
    }
  }

  if (xAxis.show && chartType !== "pie") {
    parts.push(
      `<line x1="${formatNumber(plotX)}" y1="${formatNumber(plotY + plotH)}" x2="${formatNumber(plotX + plotW)}" y2="${formatNumber(plotY + plotH)}" stroke="#888888" stroke-width="1"/>`,
    );
  }
  if (yAxis.show && chartType !== "pie") {
    parts.push(
      `<line x1="${formatNumber(plotX)}" y1="${formatNumber(plotY)}" x2="${formatNumber(plotX)}" y2="${formatNumber(plotY + plotH)}" stroke="#888888" stroke-width="1"/>`,
    );
  }

  if (xAxis.show && xAxis.labels && chartType !== "pie") {
    const stepX = plotW / Math.max(categories.length, 1);
    for (const [i, cat] of categories.entries()) {
      const x = plotX + (i + 0.5) * stepX;
      const y = plotY + plotH + 14;
      parts.push(
        `<text x="${formatNumber(x)}" y="${formatNumber(y)}" font-family="Arial" font-size="9" fill="#666666" text-anchor="middle">${escapeXml(cat)}</text>`,
      );
    }
  }

  if (showLegend) {
    const legendY = box.height - legendHeight + 4;
    const itemWidth = 80;
    const totalWidth = series.length * itemWidth;
    let legendX = plotX + (plotW - totalWidth) / 2;

    for (const s of series) {
      const fill = colorValue(s.color, context.theme);
      legendParts.push(
        `<rect x="${formatNumber(legendX)}" y="${formatNumber(legendY)}" width="8" height="8" fill="${fill}"/>`,
      );
      legendParts.push(
        `<text x="${formatNumber(legendX + 12)}" y="${formatNumber(legendY + 7)}" font-family="Arial" font-size="8" fill="#666666" text-anchor="start">${escapeXml(s.name)}</text>`,
      );
      legendX += itemWidth;
    }
  }

  const transform = transformAttribute(box, element.transform);
  const legend = legendParts.length > 0 ? `<g${transform} opacity="${element.opacity}">${legendParts.join("")}</g>` : "";

  return `<g${transform} opacity="${element.opacity}"${accessibility(element)}>${parts.join("")}</g>${legend}`;
}
