import path from "node:path";

import {
  createPresentation,
  type Box,
  type PresentationDocument,
  type PresentationSlide,
  type TextParagraphInput,
} from "@pptkit/core";

import type { ChartPlan, DeckSpec, ImagePlan, SlidePlan, StructuralIssue } from "../contracts.js";
import { getTheme, SLIDE, type ThemeTokens } from "./themes.js";

const CONTENT_TOP = 116;

function solid(color: string, opacity?: number) {
  return { type: "solid" as const, color, ...(opacity === undefined ? {} : { opacity }) };
}

function paragraph(text: string, tokens: ThemeTokens, options: { size?: number; color?: string; bold?: boolean; bullet?: boolean; align?: "left" | "center" | "right"; font?: string } = {}): TextParagraphInput {
  return {
    style: {
      align: options.align ?? "left",
      spaceAfter: options.bullet ? 8 : 0,
      bullet: options.bullet ? { type: "bullet" } : { type: "none" },
      ...(options.bullet ? { indent: 18, hanging: 5 } : {}),
    },
    runs: [{
      text,
      style: {
        fontFamily: options.font ?? tokens.bodyFont,
        fontSize: options.size ?? tokens.bodySize,
        color: options.color ?? tokens.text,
        bold: options.bold ?? false,
      },
    }],
  };
}

function addStyledText(
  slide: PresentationSlide,
  tokens: ThemeTokens,
  content: string | string[],
  box: Box,
  options: { size?: number; color?: string; bold?: boolean; bullet?: boolean; align?: "left" | "center" | "right"; font?: string; verticalAlign?: "top" | "middle" | "bottom"; name?: string } = {},
) {
  const paragraphs = (Array.isArray(content) ? content : [content]).map((text) => paragraph(text, tokens, options));
  return slide.addElement({
    type: "text",
    name: options.name,
    content: paragraphs,
    box,
    frame: { margin: 0, verticalAlign: options.verticalAlign ?? "top", autoFit: { mode: "shrink", fontScale: 0.78 } },
  });
}

function addCard(slide: PresentationSlide, tokens: ThemeTokens, box: Box, color = tokens.surface) {
  return slide.addElement({
    type: "shape",
    shape: tokens.radius === 0 ? "rect" : "roundRect",
    box,
    style: { fill: solid(color), stroke: { paint: solid(tokens.text, 0.08), width: 1 } },
    accessibility: { decorative: true },
  });
}

function addTitle(slide: PresentationSlide, tokens: ThemeTokens, title: string, index: number) {
  addStyledText(slide, tokens, title, { x: tokens.margin, y: 42, width: 720, height: 50 }, {
    size: tokens.titleSize,
    bold: true,
    font: tokens.headingFont,
    name: "Slide title",
  });
  slide.addElement({
    type: "shape",
    shape: "rect",
    box: { x: tokens.margin, y: 100, width: tokens.id === "swiss-grid" ? 72 : 36, height: 4 },
    style: { fill: solid(tokens.accent) },
    accessibility: { decorative: true },
  });
  addStyledText(slide, tokens, String(index).padStart(2, "0"), { x: 866, y: 48, width: 40, height: 22 }, {
    size: tokens.captionSize,
    color: tokens.muted,
    align: "right",
    name: "Slide number",
  });
}

function addSources(slide: PresentationSlide, tokens: ThemeTokens, plan: SlidePlan) {
  const refs = plan.sourceRefs?.map((source) => source.label ? `${source.id}: ${source.label}` : source.id) ?? [];
  if (refs.length === 0) return;
  addStyledText(slide, tokens, `Sources: ${refs.join(" · ")}`, { x: tokens.margin, y: 508, width: 800, height: 16 }, {
    size: tokens.captionSize,
    color: tokens.muted,
    name: "Source references",
  });
}

function mimeTypeFor(file: string) {
  const ext = path.extname(file).toLowerCase();
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".gif") return "image/gif";
  if (ext === ".svg") return "image/svg+xml";
  return "image/png";
}

function addImage(document: PresentationDocument, slide: PresentationSlide, image: ImagePlan, box: Box) {
  const absolutePath = path.resolve(image.path);
  const asset = document.registerAsset({
    kind: "image",
    source: { type: "path", value: absolutePath },
    mimeType: mimeTypeFor(absolutePath),
    width: image.width,
    height: image.height,
    accessibility: { description: image.alt },
    dedupeKey: absolutePath,
  });
  slide.addElement({
    type: "image",
    assetId: asset.id,
    box,
    fit: image.fit ?? "cover",
    accessibility: { description: image.alt },
  });
}

function addBarChart(slide: PresentationSlide, tokens: ThemeTokens, chart: ChartPlan, box: Box) {
  const categories = chart.categories.slice(0, 8);
  const series = chart.series.slice(0, 2);
  const values = series.flatMap((item) => item.values.slice(0, categories.length));
  const max = Math.max(1, ...values.map((value) => Math.abs(value)));
  const groupWidth = box.width / Math.max(1, categories.length);
  const barWidth = Math.min(28, (groupWidth - 10) / Math.max(1, series.length));
  const palette = [tokens.accent, tokens.accent2];

  slide.addElement({ type: "connector", start: { x: box.x, y: box.y + box.height - 28 }, end: { x: box.x + box.width, y: box.y + box.height - 28 }, style: { paint: solid(tokens.muted, 0.45), width: 1 } });
  categories.forEach((category, categoryIndex) => {
    series.forEach((item, seriesIndex) => {
      const value = item.values[categoryIndex] ?? 0;
      const height = Math.max(1, (Math.abs(value) / max) * (box.height - 70));
      slide.addElement({
        type: "shape",
        shape: "rect",
        box: {
          x: box.x + categoryIndex * groupWidth + (groupWidth - barWidth * series.length) / 2 + seriesIndex * barWidth,
          y: box.y + box.height - 28 - height,
          width: Math.max(4, barWidth - 3),
          height,
        },
        style: { fill: solid(palette[seriesIndex] ?? tokens.accent) },
        accessibility: { decorative: true },
      });
    });
    addStyledText(slide, tokens, category, { x: box.x + categoryIndex * groupWidth, y: box.y + box.height - 22, width: groupWidth, height: 22 }, { size: tokens.captionSize, color: tokens.muted, align: "center" });
  });
}

function addLineChart(slide: PresentationSlide, tokens: ThemeTokens, chart: ChartPlan, box: Box) {
  const categories = chart.categories.slice(0, 8);
  const series = chart.series.slice(0, 2);
  const values = series.flatMap((item) => item.values.slice(0, categories.length));
  const max = Math.max(1, ...values);
  const min = Math.min(0, ...values);
  const range = Math.max(1, max - min);
  const palette = [tokens.accent, tokens.accent2];
  const plotBottom = box.y + box.height - 28;

  slide.addElement({ type: "connector", start: { x: box.x, y: plotBottom }, end: { x: box.x + box.width, y: plotBottom }, style: { paint: solid(tokens.muted, 0.45), width: 1 } });
  categories.forEach((category, index) => {
    const x = box.x + (index / Math.max(1, categories.length - 1)) * box.width;
    addStyledText(slide, tokens, category, { x: x - 35, y: plotBottom + 5, width: 70, height: 18 }, { size: tokens.captionSize, color: tokens.muted, align: "center" });
  });
  series.forEach((item, seriesIndex) => {
    const points = categories.map((_, index) => ({
      x: box.x + (index / Math.max(1, categories.length - 1)) * box.width,
      y: plotBottom - (((item.values[index] ?? 0) - min) / range) * (box.height - 65),
    }));
    for (let index = 1; index < points.length; index += 1) {
      slide.addElement({ type: "connector", start: points[index - 1]!, end: points[index]!, style: { paint: solid(palette[seriesIndex] ?? tokens.accent), width: 2.5 } });
    }
    for (const point of points) {
      slide.addElement({ type: "shape", shape: "ellipse", box: { x: point.x - 4, y: point.y - 4, width: 8, height: 8 }, style: { fill: solid(palette[seriesIndex] ?? tokens.accent) }, accessibility: { decorative: true } });
    }
  });
}

function addChart(slide: PresentationSlide, tokens: ThemeTokens, chart: ChartPlan, box: Box) {
  if (chart.type === "line") addLineChart(slide, tokens, chart, box);
  else addBarChart(slide, tokens, chart, box);
}

function renderSlide(document: PresentationDocument, plan: SlidePlan, tokens: ThemeTokens, index: number) {
  const slide = document.addSlide({ id: plan.id, background: solid(tokens.background), notes: plan.notes });

  if (plan.role === "cover") {
    slide.addElement({ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 18, height: SLIDE.height }, style: { fill: solid(tokens.accent) }, accessibility: { decorative: true } });
    addStyledText(slide, tokens, plan.title, { x: tokens.margin + 30, y: 145, width: 730, height: 130 }, { size: 50, bold: true, font: tokens.headingFont, verticalAlign: "middle", name: "Cover title" });
    if (plan.subtitle) addStyledText(slide, tokens, plan.subtitle, { x: tokens.margin + 32, y: 295, width: 650, height: 60 }, { size: 21, color: tokens.muted, name: "Cover subtitle" });
    addStyledText(slide, tokens, document.metadata.author ?? "PPTKit", { x: tokens.margin + 32, y: 458, width: 400, height: 24 }, { size: tokens.captionSize, color: tokens.muted });
  } else {
    addTitle(slide, tokens, plan.title, index);
    if (plan.role === "agenda") {
      const items = (plan.items ?? []).slice(0, 6);
      const height = Math.min(62, 320 / Math.max(1, items.length));
      items.forEach((item, itemIndex) => {
        addStyledText(slide, tokens, String(itemIndex + 1).padStart(2, "0"), { x: tokens.margin, y: 138 + itemIndex * height, width: 48, height: 34 }, { size: 15, color: tokens.accent, bold: true });
        addStyledText(slide, tokens, item, { x: tokens.margin + 62, y: 134 + itemIndex * height, width: 710, height: 42 }, { size: 22, bold: itemIndex === 0 });
      });
    } else if (plan.role === "section") {
      addStyledText(slide, tokens, plan.message ?? plan.subtitle ?? "", { x: tokens.margin, y: 188, width: 720, height: 150 }, { size: 34, color: tokens.accent, bold: true, font: tokens.headingFont, verticalAlign: "middle" });
    } else if (plan.role === "statement") {
      addStyledText(slide, tokens, plan.message ?? "", { x: tokens.margin, y: 160, width: 800, height: 220 }, { size: 34, bold: true, font: tokens.headingFont, verticalAlign: "middle" });
      if (plan.items?.length) addStyledText(slide, tokens, plan.items.slice(0, 3), { x: tokens.margin, y: 402, width: 760, height: 72 }, { size: 16, color: tokens.muted, bullet: true });
    } else if (plan.role === "image") {
      const left = { x: tokens.margin, y: CONTENT_TOP + 18, width: 340, height: 330 };
      const right = { x: 440, y: CONTENT_TOP, width: 464, height: 338 };
      if (plan.message) addStyledText(slide, tokens, plan.message, { ...left, height: 92 }, { size: 25, bold: true, font: tokens.headingFont });
      if (plan.items?.length) addStyledText(slide, tokens, plan.items.slice(0, 5), { x: left.x, y: left.y + 120, width: left.width, height: 190 }, { bullet: true });
      if (plan.image) addImage(document, slide, plan.image, right);
      else addCard(slide, tokens, right, tokens.surface);
    } else if (plan.role === "kpi") {
      const kpis = (plan.kpis ?? []).slice(0, 4);
      const width = (SLIDE.width - tokens.margin * 2 - tokens.gap * Math.max(0, kpis.length - 1)) / Math.max(1, kpis.length);
      kpis.forEach((kpi, kpiIndex) => {
        const x = tokens.margin + kpiIndex * (width + tokens.gap);
        addCard(slide, tokens, { x, y: 162, width, height: 238 });
        addStyledText(slide, tokens, kpi.value, { x: x + 18, y: 188, width: width - 36, height: 72 }, { size: 38, color: tokens.accent, bold: true, font: tokens.headingFont });
        addStyledText(slide, tokens, kpi.label, { x: x + 18, y: 274, width: width - 36, height: 52 }, { size: 18, bold: true });
        if (kpi.detail) addStyledText(slide, tokens, kpi.detail, { x: x + 18, y: 340, width: width - 36, height: 42 }, { size: 13, color: tokens.muted });
      });
    } else if (plan.role === "comparison") {
      const columns = [plan.comparison?.left, plan.comparison?.right];
      columns.forEach((column, columnIndex) => {
        const x = tokens.margin + columnIndex * 426;
        addCard(slide, tokens, { x, y: 142, width: 402, height: 318 });
        addStyledText(slide, tokens, column?.heading ?? "", { x: x + 24, y: 168, width: 350, height: 42 }, { size: 23, bold: true, color: columnIndex === 0 ? tokens.accent : tokens.accent2 });
        addStyledText(slide, tokens, (column?.items ?? []).slice(0, 5), { x: x + 24, y: 232, width: 350, height: 196 }, { bullet: true });
      });
    } else if (plan.role === "process") {
      const steps = (plan.steps ?? []).slice(0, 6);
      const width = (SLIDE.width - tokens.margin * 2 - tokens.gap * Math.max(0, steps.length - 1)) / Math.max(1, steps.length);
      steps.forEach((step, stepIndex) => {
        const x = tokens.margin + stepIndex * (width + tokens.gap);
        addStyledText(slide, tokens, String(stepIndex + 1).padStart(2, "0"), { x, y: 166, width, height: 40 }, { size: 22, color: tokens.accent, bold: true });
        addCard(slide, tokens, { x, y: 222, width, height: 176 });
        addStyledText(slide, tokens, step, { x: x + 14, y: 242, width: width - 28, height: 132 }, { size: 16, bold: true, verticalAlign: "middle" });
        if (stepIndex < steps.length - 1) slide.addElement({ type: "connector", start: { x: x + width, y: 310 }, end: { x: x + width + tokens.gap, y: 310 }, style: { paint: solid(tokens.accent), width: 1.5, endArrow: "triangle" } });
      });
    } else if (plan.role === "table") {
      if (plan.chart) {
        addChart(slide, tokens, plan.chart, { x: tokens.margin, y: 148, width: 530, height: 300 });
        const legend = plan.chart.series.slice(0, 2).map((series, seriesIndex) => `${seriesIndex === 0 ? "●" : "■"} ${series.name}`);
        addStyledText(slide, tokens, legend, { x: 650, y: 170, width: 230, height: 100 }, { size: 15, color: tokens.muted });
      } else if (plan.table) {
        const headers = plan.table.headers.slice(0, 6);
        const rows = plan.table.rows.slice(0, 8);
        const width = 840 / Math.max(1, headers.length);
        slide.addElement({
          type: "table",
          box: { x: tokens.margin, y: 142, width: 840, height: Math.min(320, 42 + rows.length * 34) },
          columns: headers.map(() => width),
          rows: [
            { height: 38, cells: headers.map((value) => ({ content: [paragraph(value, tokens, { size: 14, bold: true, color: "FFFFFF" })], style: { fill: solid(tokens.accent), margin: 8 } })) },
            ...rows.map((row, rowIndex) => ({ height: 34, cells: headers.map((_, index) => ({ content: [paragraph(String(row[index] ?? ""), tokens, { size: 13 })], style: { fill: solid(rowIndex % 2 === 0 ? tokens.surface : tokens.background), margin: 7 } })) })),
          ],
        });
      }
    } else if (plan.role === "closing") {
      addStyledText(slide, tokens, plan.message ?? "", { x: tokens.margin, y: 166, width: 760, height: 150 }, { size: 34, bold: true, font: tokens.headingFont, verticalAlign: "middle" });
      if (plan.items?.length) addStyledText(slide, tokens, plan.items.slice(0, 4), { x: tokens.margin, y: 356, width: 680, height: 110 }, { bullet: true });
    }
  }
  addSources(slide, tokens, plan);
}

export function validateDeckSpec(spec: DeckSpec): StructuralIssue[] {
  const issues: StructuralIssue[] = [];
  if (spec.slides.length < spec.brief.slideCountRange[0] || spec.slides.length > spec.brief.slideCountRange[1]) {
    issues.push({ severity: "warning", code: "slide-count-range", message: `Slide count ${spec.slides.length} is outside ${spec.brief.slideCountRange.join("–")}.` });
  }
  const ids = new Set<string>();
  for (const slide of spec.slides) {
    if (ids.has(slide.id)) issues.push({ severity: "error", code: "duplicate-slide-id", message: `Duplicate slide id: ${slide.id}`, slideId: slide.id });
    ids.add(slide.id);
    if (!slide.title.trim()) issues.push({ severity: "error", code: "missing-title", message: "Every slide requires a title.", slideId: slide.id });
    if (slide.role === "kpi" && !(slide.kpis?.length)) issues.push({ severity: "error", code: "missing-kpis", message: "KPI slide requires kpis.", slideId: slide.id });
    if (slide.role === "comparison" && !slide.comparison) issues.push({ severity: "error", code: "missing-comparison", message: "Comparison slide requires comparison content.", slideId: slide.id });
    if (slide.role === "process" && !(slide.steps?.length)) issues.push({ severity: "error", code: "missing-steps", message: "Process slide requires steps.", slideId: slide.id });
    if (slide.role === "table" && !slide.table && !slide.chart) issues.push({ severity: "error", code: "missing-data", message: "Table slide requires table or chart data.", slideId: slide.id });
    if (slide.chart && (slide.chart.series.length > 2 || slide.chart.categories.length > 8)) issues.push({ severity: "warning", code: "chart-density", message: "Charts render at most two series and eight categories.", slideId: slide.id });
  }
  return issues;
}

export function authorPresentation(spec: DeckSpec): PresentationDocument {
  const tokens = getTheme(spec.brief.themeId);
  const document = createPresentation({
    metadata: { title: spec.brief.title, author: spec.brief.author ?? "PPTKit", language: spec.brief.language, subject: spec.brief.purpose },
    size: SLIDE,
    theme: {
      name: tokens.name,
      colors: { background1: tokens.background, background2: tokens.surface, text1: tokens.text, text2: tokens.muted, accent1: tokens.accent, accent2: tokens.accent2 },
      fonts: { heading: tokens.headingFont, body: tokens.bodyFont },
    },
  });
  spec.slides.forEach((slide, index) => renderSlide(document, slide, tokens, index + 1));
  return document;
}
