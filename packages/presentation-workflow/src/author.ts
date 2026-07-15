import {
  createPresentation,
  type Box,
  type PresentationDocument,
  type PresentationSlide,
  type TextParagraphInput,
} from "@pptkit/core";

import type {
  AssetResolver,
  ChartPlan,
  DeckSessionV1,
  DeckSpec,
  SlidePlan,
  StructuralIssue,
} from "./contracts.js";
import { getTheme, SLIDE, type ThemeTokens } from "./themes.js";

const CONTENT_TOP = 116;

function solid(color: string, opacity?: number) {
  return { type: "solid" as const, color, ...(opacity === undefined ? {} : { opacity }) };
}

function paragraph(
  text: string,
  tokens: ThemeTokens,
  options: { size?: number; color?: string; bold?: boolean; bullet?: boolean; align?: "left" | "center" | "right"; font?: string } = {},
): TextParagraphInput {
  return {
    style: {
      align: options.align ?? "left",
      spaceAfter: options.bullet ? 8 : 0,
      bullet: options.bullet ? { type: "bullet" } : { type: "none" },
      ...(options.bullet ? { indent: 27, hanging: 27 } : {}),
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

function addText(
  slide: PresentationSlide,
  tokens: ThemeTokens,
  content: string | string[],
  box: Box,
  options: { size?: number; color?: string; bold?: boolean; bullet?: boolean; align?: "left" | "center" | "right"; font?: string; verticalAlign?: "top" | "middle" | "bottom"; name?: string } = {},
) {
  slide.addElement({
    type: "text",
    ...(options.name ? { name: options.name } : {}),
    content: (Array.isArray(content) ? content : [content]).map((text) => paragraph(text, tokens, options)),
    box,
    frame: { margin: 0, verticalAlign: options.verticalAlign ?? "top", autoFit: { mode: "shrink", fontScale: 0.78 } },
  });
}

function addCard(slide: PresentationSlide, tokens: ThemeTokens, box: Box, color = tokens.surface) {
  slide.addElement({
    type: "shape",
    shape: tokens.radius === 0 ? "rect" : "roundRect",
    box,
    style: { fill: solid(color), stroke: { paint: solid(tokens.text, 0.08), width: 1 } },
    accessibility: { decorative: true },
  });
}

function addTitle(slide: PresentationSlide, tokens: ThemeTokens, title: string, index: number) {
  addText(slide, tokens, title, { x: tokens.margin, y: 42, width: 720, height: 50 }, {
    size: tokens.titleSize, bold: true, font: tokens.headingFont, name: "Slide title",
  });
  slide.addElement({
    type: "shape", shape: "rect", box: { x: tokens.margin, y: 100, width: tokens.id === "swiss-grid" ? 72 : 36, height: 4 },
    style: { fill: solid(tokens.accent) }, accessibility: { decorative: true },
  });
  addText(slide, tokens, String(index).padStart(2, "0"), { x: 866, y: 48, width: 40, height: 22 }, {
    size: tokens.captionSize, color: tokens.muted, align: "right", name: "Slide number",
  });
}

function addSources(slide: PresentationSlide, tokens: ThemeTokens, plan: SlidePlan) {
  const refs = plan.sourceRefs?.map((source) => source.label ? `${source.id}: ${source.label}` : source.id) ?? [];
  if (refs.length > 0) {
    addText(slide, tokens, `Sources: ${refs.join(" · ")}`, { x: tokens.margin, y: 508, width: 800, height: 16 }, {
      size: tokens.captionSize, color: tokens.muted, name: "Source references",
    });
  }
}

function addImage(document: PresentationDocument, slide: PresentationSlide, plan: NonNullable<SlidePlan["image"]>, box: Box, resolveAsset: AssetResolver) {
  const resolved = resolveAsset(plan.assetId);
  if (!resolved) return;
  const asset = document.registerAsset({
    kind: "image",
    source: resolved.source,
    mimeType: resolved.mimeType,
    ...(plan.width === undefined ? {} : { width: plan.width }),
    ...(plan.height === undefined ? {} : { height: plan.height }),
    accessibility: { description: plan.alt },
    dedupeKey: resolved.dedupeKey ?? plan.assetId,
  });
  slide.addElement({ type: "image", assetId: asset.id, box, fit: plan.fit ?? "cover", accessibility: { description: plan.alt } });
}

function addChart(slide: PresentationSlide, tokens: ThemeTokens, chart: ChartPlan, box: Box) {
  const categories = chart.categories.slice(0, 8);
  const series = chart.series.slice(0, 2);
  const values = series.flatMap((item) => item.values.slice(0, categories.length));
  const max = Math.max(1, ...values);
  const min = chart.type === "line" ? Math.min(0, ...values) : 0;
  const range = Math.max(1, max - min);
  const palette = [tokens.accent, tokens.accent2];
  const bottom = box.y + box.height - 28;
  slide.addElement({ type: "connector", start: { x: box.x, y: bottom }, end: { x: box.x + box.width, y: bottom }, style: { paint: solid(tokens.muted, 0.45), width: 1 } });

  if (chart.type === "bar") {
    const groupWidth = box.width / Math.max(1, categories.length);
    const barWidth = Math.min(28, (groupWidth - 10) / Math.max(1, series.length));
    categories.forEach((category, categoryIndex) => {
      series.forEach((item, seriesIndex) => {
        const height = Math.max(1, (Math.abs(item.values[categoryIndex] ?? 0) / Math.max(1, max)) * (box.height - 70));
        slide.addElement({
          type: "shape", shape: "rect",
          box: { x: box.x + categoryIndex * groupWidth + (groupWidth - barWidth * series.length) / 2 + seriesIndex * barWidth, y: bottom - height, width: Math.max(4, barWidth - 3), height },
          style: { fill: solid(palette[seriesIndex] ?? tokens.accent) }, accessibility: { decorative: true },
        });
      });
      addText(slide, tokens, category, { x: box.x + categoryIndex * groupWidth, y: bottom + 5, width: groupWidth, height: 18 }, { size: tokens.captionSize, color: tokens.muted, align: "center" });
    });
    return;
  }

  categories.forEach((category, index) => {
    const x = box.x + (index / Math.max(1, categories.length - 1)) * box.width;
    addText(slide, tokens, category, { x: x - 35, y: bottom + 5, width: 70, height: 18 }, { size: tokens.captionSize, color: tokens.muted, align: "center" });
  });
  series.forEach((item, seriesIndex) => {
    const points = categories.map((_, index) => ({
      x: box.x + (index / Math.max(1, categories.length - 1)) * box.width,
      y: bottom - (((item.values[index] ?? 0) - min) / range) * (box.height - 65),
    }));
    for (let index = 1; index < points.length; index += 1) {
      slide.addElement({ type: "connector", start: points[index - 1]!, end: points[index]!, style: { paint: solid(palette[seriesIndex] ?? tokens.accent), width: 2.5 } });
    }
    for (const point of points) {
      slide.addElement({ type: "shape", shape: "ellipse", box: { x: point.x - 4, y: point.y - 4, width: 8, height: 8 }, style: { fill: solid(palette[seriesIndex] ?? tokens.accent) }, accessibility: { decorative: true } });
    }
  });
}

function renderSlide(document: PresentationDocument, plan: SlidePlan, tokens: ThemeTokens, index: number, resolveAsset: AssetResolver) {
  const slide = document.addSlide({ id: plan.id, background: solid(tokens.background), ...(plan.notes ? { notes: plan.notes } : {}) });
  if (plan.role === "cover") {
    slide.addElement({ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 18, height: SLIDE.height }, style: { fill: solid(tokens.accent) }, accessibility: { decorative: true } });
    addText(slide, tokens, plan.title, { x: tokens.margin + 30, y: 145, width: 730, height: 130 }, { size: 50, bold: true, font: tokens.headingFont, verticalAlign: "middle", name: "Cover title" });
    if (plan.subtitle) addText(slide, tokens, plan.subtitle, { x: tokens.margin + 32, y: 295, width: 650, height: 60 }, { size: 21, color: tokens.muted });
    addText(slide, tokens, document.metadata.author ?? "PPTKit", { x: tokens.margin + 32, y: 458, width: 400, height: 24 }, { size: tokens.captionSize, color: tokens.muted });
    return;
  }

  addTitle(slide, tokens, plan.title, index);
  if (plan.role === "agenda") {
    const items = (plan.items ?? []).slice(0, 6);
    const height = Math.min(62, 320 / Math.max(1, items.length));
    items.forEach((item, itemIndex) => {
      addText(slide, tokens, String(itemIndex + 1).padStart(2, "0"), { x: tokens.margin, y: 138 + itemIndex * height, width: 48, height: 34 }, { size: 15, color: tokens.accent, bold: true });
      addText(slide, tokens, item, { x: tokens.margin + 62, y: 134 + itemIndex * height, width: 710, height: 42 }, { size: 22, bold: itemIndex === 0 });
    });
  } else if (plan.role === "section") {
    addText(slide, tokens, plan.message ?? plan.subtitle ?? "", { x: tokens.margin, y: 188, width: 720, height: 150 }, { size: 34, color: tokens.accent, bold: true, font: tokens.headingFont, verticalAlign: "middle" });
  } else if (plan.role === "statement") {
    addText(slide, tokens, plan.message ?? "", { x: tokens.margin, y: 160, width: 800, height: 220 }, { size: 34, bold: true, font: tokens.headingFont, verticalAlign: "middle" });
    if (plan.items?.length) addText(slide, tokens, plan.items.slice(0, 3), { x: tokens.margin, y: 402, width: 760, height: 72 }, { size: 16, color: tokens.muted, bullet: true });
  } else if (plan.role === "image") {
    const left = { x: tokens.margin, y: CONTENT_TOP + 18, width: 340, height: 330 };
    const right = { x: 440, y: CONTENT_TOP, width: 464, height: 338 };
    if (plan.message) addText(slide, tokens, plan.message, { ...left, height: 92 }, { size: 25, bold: true, font: tokens.headingFont });
    if (plan.items?.length) addText(slide, tokens, plan.items.slice(0, 5), { x: left.x, y: left.y + 120, width: left.width, height: 190 }, { bullet: true });
    if (plan.image) addImage(document, slide, plan.image, right, resolveAsset);
    else addCard(slide, tokens, right);
  } else if (plan.role === "kpi") {
    const kpis = (plan.kpis ?? []).slice(0, 4);
    const width = (SLIDE.width - tokens.margin * 2 - tokens.gap * Math.max(0, kpis.length - 1)) / Math.max(1, kpis.length);
    kpis.forEach((kpi, kpiIndex) => {
      const x = tokens.margin + kpiIndex * (width + tokens.gap);
      addCard(slide, tokens, { x, y: 162, width, height: 238 });
      addText(slide, tokens, kpi.value, { x: x + 18, y: 188, width: width - 36, height: 72 }, { size: 38, color: tokens.accent, bold: true, font: tokens.headingFont });
      addText(slide, tokens, kpi.label, { x: x + 18, y: 274, width: width - 36, height: 52 }, { size: 18, bold: true });
      if (kpi.detail) addText(slide, tokens, kpi.detail, { x: x + 18, y: 340, width: width - 36, height: 42 }, { size: 13, color: tokens.muted });
    });
  } else if (plan.role === "comparison") {
    [plan.comparison?.left, plan.comparison?.right].forEach((column, columnIndex) => {
      const x = tokens.margin + columnIndex * 426;
      addCard(slide, tokens, { x, y: 142, width: 402, height: 318 });
      addText(slide, tokens, column?.heading ?? "", { x: x + 24, y: 168, width: 350, height: 42 }, { size: 23, bold: true, color: columnIndex === 0 ? tokens.accent : tokens.accent2 });
      addText(slide, tokens, (column?.items ?? []).slice(0, 5), { x: x + 24, y: 232, width: 350, height: 196 }, { bullet: true });
    });
  } else if (plan.role === "process") {
    const steps = (plan.steps ?? []).slice(0, 6);
    const width = (SLIDE.width - tokens.margin * 2 - tokens.gap * Math.max(0, steps.length - 1)) / Math.max(1, steps.length);
    steps.forEach((step, stepIndex) => {
      const x = tokens.margin + stepIndex * (width + tokens.gap);
      addText(slide, tokens, String(stepIndex + 1).padStart(2, "0"), { x, y: 166, width, height: 40 }, { size: 22, color: tokens.accent, bold: true });
      addCard(slide, tokens, { x, y: 222, width, height: 176 });
      addText(slide, tokens, step, { x: x + 14, y: 242, width: width - 28, height: 132 }, { size: 16, bold: true, verticalAlign: "middle" });
      if (stepIndex < steps.length - 1) slide.addElement({ type: "connector", start: { x: x + width, y: 310 }, end: { x: x + width + tokens.gap, y: 310 }, style: { paint: solid(tokens.accent), width: 1.5, endArrow: "triangle" } });
    });
  } else if (plan.role === "table") {
    if (plan.chart) {
      addChart(slide, tokens, plan.chart, { x: tokens.margin, y: 148, width: 530, height: 300 });
      addText(slide, tokens, plan.chart.series.slice(0, 2).map((series, seriesIndex) => `${seriesIndex === 0 ? "●" : "■"} ${series.name}`), { x: 650, y: 170, width: 230, height: 100 }, { size: 15, color: tokens.muted });
    } else if (plan.table) {
      const headers = plan.table.headers.slice(0, 6);
      const rows = plan.table.rows.slice(0, 8);
      const width = 840 / Math.max(1, headers.length);
      slide.addElement({
        type: "table", box: { x: tokens.margin, y: 142, width: 840, height: Math.min(320, 42 + rows.length * 34) }, columns: headers.map(() => width),
        rows: [
          { height: 38, cells: headers.map((value) => ({ content: [paragraph(value, tokens, { size: 14, bold: true, color: "FFFFFF" })], style: { fill: solid(tokens.accent), margin: 8 } })) },
          ...rows.map((row, rowIndex) => ({ height: 34, cells: headers.map((_, cellIndex) => ({ content: [paragraph(String(row[cellIndex] ?? ""), tokens, { size: 13 })], style: { fill: solid(rowIndex % 2 === 0 ? tokens.surface : tokens.background), margin: 7 } })) })),
        ],
      });
    }
  } else if (plan.role === "closing") {
    addText(slide, tokens, plan.message ?? "", { x: tokens.margin, y: 166, width: 760, height: 150 }, { size: 34, bold: true, font: tokens.headingFont, verticalAlign: "middle" });
    if (plan.items?.length) addText(slide, tokens, plan.items.slice(0, 4), { x: tokens.margin, y: 356, width: 680, height: 110 }, { bullet: true });
  }
  addSources(slide, tokens, plan);
}

export function validateDeckSpec(spec: DeckSpec, availableAssetIds: ReadonlySet<string> = new Set()): StructuralIssue[] {
  const issues: StructuralIssue[] = [];
  if (spec.slides.length < spec.brief.slideCountRange[0] || spec.slides.length > spec.brief.slideCountRange[1]) {
    issues.push({ severity: "warning", code: "slide-count-range", message: `Slide count ${spec.slides.length} is outside ${spec.brief.slideCountRange.join("–")}.` });
  }
  const ids = new Set<string>();
  for (const slide of spec.slides) {
    if (ids.has(slide.id)) issues.push({ severity: "error", code: "duplicate-slide-id", message: `Duplicate slide id: ${slide.id}`, slideId: slide.id });
    ids.add(slide.id);
    if (!slide.title.trim()) issues.push({ severity: "error", code: "missing-title", message: "Every slide requires a title.", slideId: slide.id });
    if (slide.role === "kpi" && !slide.kpis?.length) issues.push({ severity: "error", code: "missing-kpis", message: "KPI slide requires kpis.", slideId: slide.id });
    if (slide.role === "comparison" && !slide.comparison) issues.push({ severity: "error", code: "missing-comparison", message: "Comparison slide requires comparison content.", slideId: slide.id });
    if (slide.role === "process" && !slide.steps?.length) issues.push({ severity: "error", code: "missing-steps", message: "Process slide requires steps.", slideId: slide.id });
    if (slide.role === "table" && !slide.table && !slide.chart) issues.push({ severity: "error", code: "missing-data", message: "Table slide requires table or chart data.", slideId: slide.id });
    if (slide.chart && (slide.chart.series.length > 2 || slide.chart.categories.length > 8)) issues.push({ severity: "warning", code: "chart-density", message: "Charts render at most two series and eight categories.", slideId: slide.id });
    if (slide.image && !availableAssetIds.has(slide.image.assetId)) issues.push({ severity: "error", code: "missing-asset", message: `Image asset ${slide.image.assetId} is not available.`, slideId: slide.id });
  }
  return issues;
}

export function authorPresentation(spec: DeckSpec, resolveAsset: AssetResolver = () => undefined): PresentationDocument {
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
  spec.slides.forEach((slide, index) => renderSlide(document, slide, tokens, index + 1, resolveAsset));
  return document;
}

export function createSessionAssetResolver(session: DeckSessionV1): AssetResolver {
  const assets = new Map(session.assets.map((asset) => [asset.id, asset]));
  return (assetId) => {
    const asset = assets.get(assetId);
    if (!asset?.dataUrl) return undefined;
    return { source: { type: "url", value: asset.dataUrl }, mimeType: asset.mimeType, dedupeKey: asset.id };
  };
}
