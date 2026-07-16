import { createPresentation, type PresentationDocument } from "@pptkit/core";

import { renderSlide } from "./authoring/recipes.js";
import type {
  AssetResolver,
  DeckSpec,
  SlidePlan,
  SourceRef,
  StructuralIssue,
} from "./contracts.js";
import { getTheme, SLIDE } from "./themes.js";

function visibleStrings(plan: SlidePlan): string[] {
  return [
    plan.title,
    plan.subtitle,
    plan.message,
    ...(plan.items ?? []),
    ...(plan.steps ?? []),
    ...(plan.kpis ?? []).flatMap((kpi) => [kpi.value, kpi.label, kpi.detail]),
    plan.comparison?.left.heading,
    ...(plan.comparison?.left.items ?? []),
    plan.comparison?.right.heading,
    ...(plan.comparison?.right.items ?? []),
    ...(plan.table?.headers ?? []),
    ...(plan.table?.rows.flat() ?? []),
  ].filter((value): value is string => typeof value === "string" && value.length > 0);
}

function weightedLength(value: string): number {
  return Array.from(value).reduce((length, character) => length + (/\p{Script=Han}|\p{Script=Hiragana}|\p{Script=Katakana}|\p{Script=Hangul}/u.test(character) ? 2 : 1), 0);
}

function containsInternalMetadata(value: string): boolean {
  if (/\bSources?:\s|\bsrc-\d{2}-/i.test(value)) return true;
  const proseWithoutUrls = value.replace(/\bhttps?:\/\/\S+/gi, "");
  return /(?:^|\s|[\\/])[^\s\\/]+\.(?:md|txt|pdf|docx|pptx|xlsx)\b/i.test(proseWithoutUrls);
}

function sourceNotes(plan: SlidePlan): string | undefined {
  // Early DeckSessionV1 producers emitted source IDs as strings. Keep those
  // sessions usable without restoring the old visible source-footer behavior.
  const refs = (plan.sourceRefs ?? []) as Array<SourceRef | string>;
  if (!plan.notes?.trim() && refs.length === 0) return undefined;
  const sections: string[] = [];
  if (plan.notes?.trim()) sections.push(plan.notes.trim());
  if (refs.length > 0) {
    sections.push([
      "Source references (provenance only — not slide content):",
      ...refs.map((source) => typeof source === "string"
        ? `- ${source}`
        : `- ${source.id}${source.label ? ` — ${source.label}` : ""}`),
    ].join("\n"));
  }
  return sections.join("\n\n");
}

function addDensityIssues(plan: SlidePlan, issues: StructuralIssue[]) {
  const titleWeight = weightedLength(plan.title);
  if (titleWeight > 82) {
    issues.push({
      severity: "warning",
      code: "title-density",
      message: "The title is too long for the presentation-scale title treatment; shorten it or split the slide.",
      slideId: plan.id,
    });
  }
  if (plan.message && weightedLength(plan.message) > 250) {
    issues.push({
      severity: "warning",
      code: "message-density",
      message: "The primary message exceeds the safe reading measure; edit it down or continue it on another slide.",
      slideId: plan.id,
    });
  }
  const denseItem = [...(plan.items ?? []), ...(plan.steps ?? [])].find((item) => weightedLength(item) > 120);
  if (denseItem) {
    issues.push({
      severity: "warning",
      code: "item-density",
      message: "At least one list or process item is too long for the role layout; shorten it or split the slide.",
      slideId: plan.id,
    });
  }
  const denseCell = plan.table?.rows.flat().find((cell) => weightedLength(cell) > 110);
  if (denseCell) {
    issues.push({
      severity: "warning",
      code: "table-cell-density",
      message: "At least one table cell is too long for a readable 15 pt minimum; reduce the copy or use a narrative slide.",
      slideId: plan.id,
    });
  }
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
    if ((slide.items?.length ?? 0) > 6 || (slide.steps?.length ?? 0) > 6 || (slide.kpis?.length ?? 0) > 4 || (slide.table?.rows.length ?? 0) > 8) {
      issues.push({ severity: "warning", code: "role-density", message: "The selected role contains more items than its readable layout supports; split the slide instead of truncating it.", slideId: slide.id });
    }
    if (slide.image && !availableAssetIds.has(slide.image.assetId)) issues.push({ severity: "error", code: "missing-asset", message: `Image asset ${slide.image.assetId} is not available.`, slideId: slide.id });
    if (visibleStrings(slide).some(containsInternalMetadata)) {
      issues.push({
        severity: "warning",
        code: "internal-metadata-visible",
        message: "Visible slide copy appears to contain an internal source ID, file path, or workflow label; keep provenance in sourceRefs/notes and use a human-readable citation only when requested.",
        slideId: slide.id,
      });
    }
    addDensityIssues(slide, issues);
  }
  return issues;
}

export function authorPresentation(spec: DeckSpec, resolveAsset: AssetResolver = () => undefined): PresentationDocument {
  const tokens = getTheme(spec.brief.themeId);
  const document = createPresentation({
    metadata: {
      title: spec.brief.title,
      author: spec.brief.author ?? "PPTKit",
      language: spec.brief.language,
      subject: spec.brief.purpose,
    },
    size: SLIDE,
    theme: {
      name: tokens.name,
      colors: {
        background1: tokens.background,
        background2: tokens.surface,
        text1: tokens.text,
        text2: tokens.muted,
        accent1: tokens.accent,
        accent2: tokens.accent2,
      },
      fonts: { heading: tokens.headingFont, body: tokens.bodyFont },
    },
  });
  spec.slides.forEach((slide, index) => renderSlide(document, slide, tokens, index + 1, resolveAsset, sourceNotes(slide)));
  return document;
}
