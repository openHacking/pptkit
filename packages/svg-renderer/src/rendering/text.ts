import type {
  Box,
  ElementAction,
  NormalizedPresentationTheme,
  NormalizedTextFrameStyle,
  NormalizedTextParagraph,
} from "@pptkit/core";
import { formatListMarker, resolveTextFontSegments } from "@pptkit/core";
import { cssFontFamily, escapeXml, safeActionUrl } from "./escape.js";
import { needsBidiParagraphLayout } from "./bidi-ranges.js";
import { colorValue } from "./style.js";
import type { SvgRenderWarning } from "../types/public.js";

interface TextContext {
  theme: NormalizedPresentationTheme;
  warnings: SvgRenderWarning[];
  slideId: string;
  elementId: string;
  skipLinkTabStop?: boolean;
  slideActionTargets?: ReadonlyMap<string, string>;
  recordFont?: (family: string, bold: boolean, italic: boolean, text: string, fontSize: number, language?: string) => void;
}

type TextRun = NormalizedTextParagraph["runs"][number];

function fontRuns(run: TextRun, theme: NormalizedPresentationTheme): TextRun[] {
  return resolveTextFontSegments(run.text, run.style, theme).map(segment => ({ ...run, text: segment.text, style: { ...run.style, fontFamily: segment.family } }));
}

interface NativeTextSegment {
  run: TextRun;
  text: string;
}

interface NativeTextGrapheme {
  text: string;
  segments: NativeTextSegment[];
}

const graphemeSegmenter = typeof Intl !== "undefined" && typeof Intl.Segmenter === "function" ? new Intl.Segmenter("und", { granularity: "grapheme" }) : undefined;

/** Segment joined text first so a selected-font boundary cannot split a cluster. */
function* paragraphGraphemes(runs: TextRun[]): Generator<NativeTextGrapheme> {
  const text = runs.map(run => run.text).join("");
  let runIndex = 0;
  let runStart = 0;
  for (const item of graphemeSegmenter!.segment(text)) {
    const end = item.index + item.segment.length;
    const segments: NativeTextSegment[] = [];
    let offset = item.index;
    while (offset < end) {
      const run = runs[runIndex]!;
      const runEnd = runStart + run.text.length;
      if (offset >= runEnd) { runIndex++; runStart = runEnd; continue; }
      const through = Math.min(end, runEnd);
      segments.push({ run, text: run.text.slice(offset - runStart, through - runStart) });
      offset = through;
    }
    yield { text: item.segment, segments };
  }
}

interface NativeTextLine {
  first: boolean;
  hardBreak: string;
  softBreakSpace: string;
  fontSize: number;
  lineHeight: number;
  paragraph: NormalizedTextParagraph;
  segments: NativeTextSegment[];
  y: number;
}

// Keep wrapping deterministic without depending on browser canvas metrics or
// hidden host fonts. Latin glyph advances vary too much for one average width:
// treating "i" like "m" makes compact bold headings wrap substantially earlier
// than they do in PowerPoint.
const NORMAL_CHARACTER_WIDTH = 0.52;
const NARROW_CHARACTER_WIDTH = 0.3;
const WIDE_CHARACTER_WIDTH = 0.72;
const UPPERCASE_CHARACTER_WIDTH = 0.6;
const DIGIT_CHARACTER_WIDTH = 0.55;
const SPACE_WIDTH = 0.28;
const PUNCTUATION_WIDTH = 0.32;
const POWERPOINT_WIDTH_HEADROOM = 1.06;
const BOLD_WIDTH_MULTIPLIER = 1.03;
// An alphabetic baseline at 0.8em removes the lower HTML line-box offset while
// retaining the authored top inset used by the DrawingML exporter.
const TEXT_BASELINE_RATIO = 0.8;

function isWideCharacter(character: string): boolean {
  const codePoint = character.codePointAt(0) ?? 0;
  return (
    (codePoint >= 0x1100 && codePoint <= 0x11ff) ||
    (codePoint >= 0x2e80 && codePoint <= 0x9fff) ||
    (codePoint >= 0xac00 && codePoint <= 0xd7af) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0x1f300 && codePoint <= 0x1faff)
  );
}

function scaledFontSize(run: TextRun, scale: number): number {
  return run.style.fontSize * scale;
}

function formatNumber(value: number): number {
  return Math.round(value * 1000) / 1000;
}

function characterWidth(character: string, run: TextRun, scale: number): number {
  if (/[\p{Mn}\p{Me}\u200c\u200d]/u.test(character)) return 0;
  const fontSize = scaledFontSize(run, scale);
  const base = /\s/u.test(character)
    ? SPACE_WIDTH
    : isWideCharacter(character)
      ? 1
      : /[,.;:!?()[\]{}'"`-]/u.test(character)
        ? PUNCTUATION_WIDTH
        : /[fijltIr|]/u.test(character)
          ? NARROW_CHARACTER_WIDTH
          : /[mMwWOQ%@]/u.test(character)
            ? WIDE_CHARACTER_WIDTH
            : /\d/u.test(character)
              ? DIGIT_CHARACTER_WIDTH
              : /[A-Z]/u.test(character)
                ? UPPERCASE_CHARACTER_WIDTH
                : NORMAL_CHARACTER_WIDTH;
  return fontSize * base * POWERPOINT_WIDTH_HEADROOM * (run.style.bold && !isWideCharacter(character) ? BOLD_WIDTH_MULTIPLIER : 1) * (run.style.italic ? 1.02 : 1);
}

function textWidth(text: string, run: TextRun, scale: number): number {
  let width = 0;
  for (const character of text) width += characterWidth(character, run, scale);
  return width;
}

function appendSegment(segments: NativeTextSegment[], run: TextRun, text: string): void {
  if (text === "") return;
  const previous = segments.at(-1);
  if (previous?.run === run) previous.text += text;
  else segments.push({ run, text });
}

function paragraphLines(paragraph: NormalizedTextParagraph, availableWidth: number, scale: number, wrap: boolean): Omit<NativeTextLine, "y">[] {
  const firstRun = bulletRun(paragraph);
  if (firstRun === undefined) return [];
  const bullet = bulletText(paragraph);
  const firstWidth = Math.max(1, availableWidth + Math.min(paragraph.style.hanging, paragraph.style.indent));
  const laterWidth = Math.max(1, availableWidth);
  const lines: Omit<NativeTextLine, "y">[] = [];
  let first = true;
  let segments: NativeTextSegment[] = [];
  let width = 0;
  let hasContent = false;
  let pendingSpace: NativeTextGrapheme[] = [];
  let endedWithBreak = false;

  const unitWidth = (unit: NativeTextGrapheme) => unit.segments.reduce((sum, segment) => sum + textWidth(segment.text, segment.run, scale), 0);
  const appendUnit = (unit: NativeTextGrapheme) => {
    for (const segment of unit.segments) appendSegment(segments, segment.run, segment.text);
    width += unitWidth(unit);
  };
  const flushSpace = () => { for (const unit of pendingSpace) appendUnit(unit); pendingSpace = []; };

  const startLine = (): void => {
    segments = [];
    width = 0;
    hasContent = false;
    if (first && bullet !== "") {
      appendSegment(segments, firstRun, bullet);
      width = textWidth(bullet, firstRun, scale);
    }
  };
  const finishLine = (softBreakSpace = "", hardBreak = ""): void => {
    const lineFontSize = segments.reduce((largest, segment) => Math.max(largest, scaledFontSize(segment.run, scale)), scaledFontSize(firstRun, scale));
    lines.push({ first, hardBreak, softBreakSpace, fontSize: lineFontSize, lineHeight: lineFontSize * paragraph.style.lineSpacing, paragraph, segments });
    first = false;
    pendingSpace = [];
    startLine();
  };
  const appendWord = (word: NativeTextGrapheme[]): void => {
    const maximum = first ? firstWidth : laterWidth;
    const candidateWidth = width + pendingSpace.reduce((sum, unit) => sum + unitWidth(unit), 0) + word.reduce((sum, unit) => sum + unitWidth(unit), 0);
    // Keep discarded soft-break whitespace available for semantic text output.
    if (wrap && hasContent && candidateWidth > maximum) finishLine(pendingSpace.map(unit => unit.text).join(""));
    flushSpace();
    for (const unit of word) {
      const activeMaximum = first ? firstWidth : laterWidth;
      if (wrap && hasContent && width + unitWidth(unit) > activeMaximum) finishLine();
      appendUnit(unit);
      hasContent = true;
    }
  };

  startLine();
  let word: NativeTextGrapheme[] = [];
  const flushWord = () => { if (word.length) appendWord(word); word = []; };
  for (const unit of paragraphGraphemes(paragraph.runs)) {
    if (unit.text === "\r\n" || unit.text === "\r" || unit.text === "\n") {
      flushWord(); flushSpace(); finishLine("", unit.text); endedWithBreak = true;
    } else if (/^\s+$/u.test(unit.text)) {
      flushWord(); pendingSpace.push(unit); endedWithBreak = false;
    } else {
      word.push(unit); endedWithBreak = false;
    }
  }
  flushWord(); flushSpace();
  if (segments.length > 0 || lines.length === 0 || endedWithBreak) finishLine();
  return lines;
}

export function canRenderTextNatively(paragraphs: NormalizedTextParagraph[], frame: NormalizedTextFrameStyle): boolean {
  // Only the bounded no-wrap tab profile can place mixed runs without a shared
  // line-fragment contract; run actions remain on the XHTML path.
  // Independently positioned SVG lines resolve bidi without paragraph context.
  // Let the browser lay out the full paragraph when wrapping RTL scripts or
  // explicit directional controls; the standalone SVG path stays unchanged.
  const contextualBidi = frame.wrap && paragraphs.some(paragraph => needsBidiParagraphLayout(paragraph.runs.map(run => run.text).join("")));
  return !contextualBidi && graphemeSegmenter !== undefined && paragraphs.every((paragraph) => paragraph.style.direction !== "rtl" && !paragraph.runs.some(run => run.style.rightToLeft === true) && ((paragraph.runs.length === 1 && paragraph.runs[0]?.action === undefined) || nativeTabProfile(paragraph, frame)));
}

/** A conservative opt-in: the native path does not implement run links or bidi layout. */
export function canRenderMixedNoWrapNatively(paragraphs: NormalizedTextParagraph[], frame: NormalizedTextFrameStyle): boolean {
  return !frame.wrap && graphemeSegmenter !== undefined && paragraphs.some(paragraph => paragraph.runs.length > 1) && paragraphs.every(paragraph =>
    paragraph.style.direction !== "rtl" && !paragraph.runs.some(run => run.action !== undefined || run.style.rightToLeft === true || run.text.includes("\t")) &&
    !needsBidiParagraphLayout(paragraph.runs.map(run => run.text).join("")));
}

/** Native opt-in text uses approximate line layout, without run actions or bidi. */
export function canRenderNativeParagraphs(paragraphs: NormalizedTextParagraph[]): boolean {
  return graphemeSegmenter !== undefined && paragraphs.every(paragraph =>
    paragraph.style.direction !== "rtl" && !paragraph.runs.some(run => run.action !== undefined || run.style.rightToLeft === true || run.text.includes("\t")) &&
    !needsBidiParagraphLayout(paragraph.runs.map(run => run.text).join("")));
}

/** Keep other shape-text layouts on their existing XHTML path. */
export function canRenderShapeTabTextNatively(paragraphs: NormalizedTextParagraph[], frame: NormalizedTextFrameStyle): boolean {
  if (graphemeSegmenter === undefined || frame.wrap) return false;
  let hasTab = false;
  const supported = paragraphs.every(paragraph => {
    if (paragraph.style.direction === "rtl" || paragraph.runs.some(run => run.style.rightToLeft === true)) return false;
    if (paragraph.runs.some(run => run.text.includes("\t"))) {
      hasTab = true;
      return nativeTabProfile(paragraph, frame);
    }
    return paragraph.runs.length === 1 && paragraph.runs[0]?.action === undefined;
  });
  return supported && hasTab;
}

function textAnchor(align: NormalizedTextParagraph["style"]["align"]): "start" | "middle" | "end" {
  if (align === "center") return "middle";
  if (align === "right") return "end";
  return "start";
}

function nativeRunAttributes(run: TextRun, scale: number, context: TextContext, text: string): string {
  const theme = context.theme;
  context.recordFont?.(fontFamily(run.style.fontFamily, theme), run.style.bold, run.style.italic, text, formatNumber(scaledFontSize(run, scale)), run.style.language);
  const decorations = [run.style.underline ? "underline" : "", run.style.strike ? "line-through" : ""].filter(Boolean).join(" ");
  return [
    `font-family="${cssFontFamily(fontFamily(run.style.fontFamily, theme))}"`,
    `font-size="${formatNumber(scaledFontSize(run, scale))}"`,
    `font-weight="${run.style.bold ? 700 : 400}"`,
    `font-style="${run.style.italic ? "italic" : "normal"}"`,
    `text-decoration="${decorations || "none"}"`,
    `fill="${colorValue(run.style.color, theme)}"`,
    ...(run.style.colorOpacity === undefined ? [] : [`fill-opacity="${run.style.colorOpacity}"`]),
    `lang="${escapeXml(run.style.language)}"`,
  ].join(" ");
}

/** Explicit no-wrap tab chunks can use SVG's native text anchors without guessing font widths. */
function nativeTabProfile(paragraph: NormalizedTextParagraph, frame: NormalizedTextFrameStyle): boolean {
  const { style, runs } = paragraph;
  if (frame.wrap || runs.length === 0 || runs.some(run => run.action !== undefined) || style.align !== "left" || style.indent !== 0 || style.hanging !== 0 || style.bullet.type !== "none" || style.tabStops?.length !== 1) return false;
  const text = runs.map(run => run.text).join("");
  if (!text.includes("\t") || !/^[\x20-\x7e\t\r\n\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]*$/u.test(text)) return false;
  const decimal = style.tabStops[0]!.align === "decimal";
  const words = /^[A-Za-z0-9\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]+(?: [A-Za-z0-9\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]+)*$/u;
  return text.split(/\r\n|\r|\n/u).every(line => {
    if (!line.includes("\t")) return true;
    const parts = line.split("\t");
    return parts.length === 2 && words.test(parts[0]!) && (decimal ? /^[0-9]+\.[0-9]+$/u.test(parts[1]!) : words.test(parts[1]!));
  });
}

function splitNativeSegments(segments: NativeTextSegment[], separator: string): [NativeTextSegment[], NativeTextSegment[]] | undefined {
  const before: NativeTextSegment[] = [], after: NativeTextSegment[] = [];
  let found = false;
  for (const segment of segments) {
    const index = found ? -1 : segment.text.indexOf(separator);
    if (index < 0) (found ? after : before).push(segment);
    else {
      if (index > 0) before.push({ ...segment, text: segment.text.slice(0, index) });
      after.push({ ...segment, text: segment.text.slice(index) });
      found = true;
    }
  }
  return found ? [before, after] : undefined;
}

export function textSvg(
  paragraphs: NormalizedTextParagraph[],
  frame: NormalizedTextFrameStyle,
  box: Box,
  context: TextContext,
): string {
  warnAlphabeticNumbering(paragraphs, context);
  warnTabLayout(paragraphs, context);
  warnRunDirection(paragraphs, context);
  if (frame.autoFit.mode === "resize") {
    context.warnings.push({
      code: "text-autofit-resize-unsupported",
      message: "Text autoFit resize cannot change SVG element geometry and was rendered inside the authored box.",
      slideId: context.slideId,
      elementId: context.elementId,
    });
  }
  const scale = frame.autoFit.mode === "shrink" ? frame.autoFit.fontScale : 1;
  const contentWidth = Math.max(1, box.width - frame.margin.left - frame.margin.right);
  const lines: NativeTextLine[] = [];
  const nativeTabParagraphs = new WeakSet<NormalizedTextParagraph>();
  let cursor = 0;
  for (const paragraph of paragraphs) {
    cursor += paragraph.style.spaceBefore;
    const availableWidth = Math.max(1, contentWidth - paragraph.style.indent);
    const effective = { ...paragraph, runs: paragraph.runs.flatMap(run => fontRuns(run, context.theme)) };
    if (nativeTabProfile(effective, frame)) nativeTabParagraphs.add(effective);
    for (const line of paragraphLines(effective, availableWidth, scale, frame.wrap)) {
      lines.push({ ...line, y: cursor });
      cursor += line.lineHeight;
    }
    cursor += paragraph.style.spaceAfter;
  }
  const availableHeight = Math.max(0, box.height - frame.margin.top - frame.margin.bottom);
  const verticalOffset = frame.verticalAlign === "middle"
    ? Math.max(0, (availableHeight - cursor) / 2)
    : frame.verticalAlign === "bottom"
      ? Math.max(0, availableHeight - cursor)
      : 0;

  return lines.map((line, lineIndex) => {
    const style = line.paragraph.style;
    const left = box.x + frame.margin.left + Math.max(0, style.indent - (line.first ? style.hanging : 0));
    const right = box.x + box.width - frame.margin.right;
    const anchor = textAnchor(style.align);
    const direction = style.direction === "ltr" ? ' direction="ltr"' : "";
    const x = anchor === "middle" ? (left + right) / 2 : anchor === "end" ? right : left;
    const y = box.y + frame.margin.top + verticalOffset + line.y + line.fontSize * TEXT_BASELINE_RATIO;
    if (nativeTabParagraphs.has(line.paragraph)) {
      const chunks = splitNativeSegments(line.segments, "\t");
      if (chunks) {
        const [prefix, suffixWithTab] = chunks;
        const suffix = suffixWithTab.map((segment, index) => index === 0 ? { ...segment, text: segment.text.slice(1) } : segment).filter(segment => segment.text !== "");
        // Observed mixed-size and East Asian tab lines place the baseline near the largest em box.
        // This remains an approximation for font-dependent ascent and descent.
        const mixedSizes = new Set(line.segments.filter(segment => segment.text !== "\t").map(segment => segment.run.style.fontSize)).size > 1;
        const eastAsian = line.segments.some(segment => /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u.test(segment.text));
        const tabY = mixedSizes || eastAsian ? y + line.fontSize * (1 - TEXT_BASELINE_RATIO) : y;
        const stop = line.paragraph.style.tabStops![0]!;
        const stopX = box.x + frame.margin.left + stop.position;
        const textChunk = (segments: NativeTextSegment[], chunkX: number, chunkAnchor: "start" | "middle" | "end", part: string, last = false) => `<text x="${chunkX}" y="${tabY}" text-anchor="${chunkAnchor}"${direction} xml:space="preserve" style="white-space:pre" data-pptkit-text-line="${lineIndex}" data-pptkit-tab-part="${part}">${segments.map(segment => `<tspan ${nativeRunAttributes(segment.run, scale, context, segment.text)}>${escapeXml(segment.text)}</tspan>`).join("")}${part === "prefix" ? '<tspan aria-hidden="true" font-size="0">&#9;</tspan>' : ""}${last && line.hardBreak !== "" ? `<tspan style="display:none">${escapeXml(line.hardBreak)}</tspan>` : ""}</text>`;
        const start = textChunk(prefix, left, "start", "prefix");
        if (stop.align === "decimal") {
          const decimal = splitNativeSegments(suffix, ".")!;
          return start + textChunk(decimal[0], stopX, "end", "integer") + textChunk(decimal[1], stopX, "start", "decimal", true);
        }
        return start + textChunk(suffix, stopX, stop.align === "center" ? "middle" : stop.align === "right" ? "end" : "start", "suffix", true);
      }
    }
    const content = line.segments.map((segment) => `<tspan ${nativeRunAttributes(segment.run, scale, context, segment.text)}>${escapeXml(segment.text)}</tspan>`).join("") +
      (line.softBreakSpace === "" ? "" : `<tspan font-size="0">${escapeXml(line.softBreakSpace)}</tspan>`) +
      (line.hardBreak === "" ? "" : `<tspan style="display:none">${escapeXml(line.hardBreak)}</tspan>`);
    return `<text x="${x}" y="${y}" text-anchor="${anchor}"${direction} xml:space="preserve" style="white-space:pre" data-pptkit-text-line="${lineIndex}">${content}</text>`;
  }).join("");
}

function actionAttributes(action: ElementAction | undefined, context: TextContext): { open: string; close: string } {
  if (action === undefined) return { open: "", close: "" };
  if (action.type === "slide") {
    const target = context.slideActionTargets?.get(action.slideId);
    if (target !== undefined) return { open: `<a href="#${target}" data-pptkit-slide-target="${escapeXml(action.slideId)}"${context.skipLinkTabStop ? ' tabindex="-1"' : ""}${action.tooltip === undefined ? "" : ` title="${escapeXml(action.tooltip)}"`}>`, close: "</a>" };
    context.warnings.push({
      code: "slide-action-not-linked",
      message: context.slideActionTargets ? `Slide action targeting ${action.slideId} has no selected document target.` : `Slide action targeting ${action.slideId} is metadata-only in standalone SVG output.`,
      slideId: context.slideId,
      elementId: context.elementId,
    });
    return { open: `<span data-pptkit-slide-target="${escapeXml(action.slideId)}">`, close: "</span>" };
  }
  const url = safeActionUrl(action.url);
  if (url === undefined) {
    context.warnings.push({
      code: "unsafe-action-url",
      message: "An action URL with an unsupported protocol was omitted.",
      slideId: context.slideId,
      elementId: context.elementId,
    });
    return { open: "", close: "" };
  }
  return {
    open: `<a href="${escapeXml(url)}" rel="noopener noreferrer"${context.skipLinkTabStop ? ' tabindex="-1"' : ""}${action.tooltip === undefined ? "" : ` title="${escapeXml(action.tooltip)}"`}>`,
    close: "</a>",
  };
}

function fontFamily(value: NormalizedTextParagraph["runs"][number]["style"]["fontFamily"], theme: NormalizedPresentationTheme): string {
  return typeof value === "string" ? value : theme.fonts[value.theme];
}

function bulletText(paragraph: NormalizedTextParagraph): string { return formatListMarker(paragraph.style.bullet); }

function warnAlphabeticNumbering(paragraphs: NormalizedTextParagraph[], context: TextContext): void {
  if (paragraphs.some(p => p.style.bullet.type === "number" && (p.style.bullet.style === "alphaLowerPeriod" || p.style.bullet.style === "alphaUpperPeriod") && p.style.bullet.startAt > 27)) {
    context.warnings.push({ code: "alphabetic-numbering-application-dependent", slideId: context.slideId, elementId: context.elementId, message: "Alphabetic numbering after 27 uses the preview's base-26 sequence; native application numbering may differ." });
  }
}

function warnTabLayout(paragraphs: NormalizedTextParagraph[], context: TextContext): void {
  if (paragraphs.some(paragraph => paragraph.runs.some(run => run.text.includes("\t")))) {
    context.warnings.push({ code: "text-tab-layout-unsupported", slideId: context.slideId, elementId: context.elementId, message: "Some no-wrap tabs use declared anchors; wrapping, overlap and other tab cases may differ from the exported presentation." });
  }
}

function warnRunDirection(paragraphs: NormalizedTextParagraph[], context: TextContext): void {
  if (paragraphs.some(paragraph => paragraph.runs.some(run => run.style.rightToLeft === true)) &&
      !context.warnings.some(warning => warning.code === "text-run-rtl-approximation" && warning.slideId === context.slideId && warning.elementId === context.elementId)) {
    context.warnings.push({ code: "text-run-rtl-approximation", slideId: context.slideId, elementId: context.elementId, message: "Preview does not implement DrawingML's selective run-level RTL character override; font selection may differ from native applications." });
  }
}

function bulletRun(paragraph: NormalizedTextParagraph): TextRun | undefined {
  // A zero-character insertion style must not hide or restyle a marker before
  // real text. Wholly empty paragraphs use the final insertion style.
  const run = paragraph.runs.find(run => run.text.length > 0) ?? paragraph.runs.at(-1);
  // PowerPoint's generated marker follows the text font/strike but is not
  // underlined with the text (observed independently in native list controls).
  return run?.style.underline ? { ...run, style: { ...run.style, underline: false } } : run;
}

function htmlRunStyle(run: TextRun, scale: number, context: TextContext, text: string): string {
  const theme = context.theme;
  context.recordFont?.(fontFamily(run.style.fontFamily, theme), run.style.bold, run.style.italic, text, formatNumber(run.style.fontSize * scale), run.style.language);
  const decorations = [run.style.underline ? "underline" : "", run.style.strike ? "line-through" : ""].filter(Boolean).join(" ");
  return [
    `font-family:${cssFontFamily(fontFamily(run.style.fontFamily, theme))}`,
    `font-size:${formatNumber(run.style.fontSize * scale)}px`,
    `font-weight:${run.style.bold ? 700 : 400}`,
    `font-style:${run.style.italic ? "italic" : "normal"}`,
    `text-decoration:${decorations || "none"}`,
    `color:${colorValue(run.style.color, theme)}`,
    ...(run.style.colorOpacity === undefined ? [] : [`opacity:${run.style.colorOpacity}`]),
  ].join(";");
}

export function textHtml(
  paragraphs: NormalizedTextParagraph[],
  frame: NormalizedTextFrameStyle,
  context: TextContext,
): string {
  warnAlphabeticNumbering(paragraphs, context);
  warnTabLayout(paragraphs, context);
  warnRunDirection(paragraphs, context);
  if (frame.autoFit.mode === "resize") {
    context.warnings.push({
      code: "text-autofit-resize-unsupported",
      message: "Text autoFit resize cannot change SVG element geometry and was rendered inside the authored box.",
      slideId: context.slideId,
      elementId: context.elementId,
    });
  }
  const direction = frame.textDirection;
  const eastAsian = direction === "eastAsianVertical";
  const stacked = direction === "stackedVertical" || direction === "stackedVerticalRtl";
  const vertical = eastAsian || stacked || direction === "mongolianVertical";
  if (vertical && !eastAsian && !context.warnings.some(w => w.code === "vertical-text-preview-approximation" && w.slideId === context.slideId && w.elementId === context.elementId)) {
    context.warnings.push({ code: "vertical-text-preview-approximation", message: "Preview approximates native column flow and glyph orientation with CSS; font shaping and wrapping may differ.", slideId: context.slideId, elementId: context.elementId });
  }
  if (eastAsian && !context.warnings.some(w => w.code === "east-asian-vertical-preview-approximation" && w.slideId === context.slideId && w.elementId === context.elementId)) {
    context.warnings.push({ code: "east-asian-vertical-preview-approximation", message: "Preview uses CSS mixed vertical glyph orientation; native fonts, punctuation and wrapping may differ.", slideId: context.slideId, elementId: context.elementId });
  }
  const scale = frame.autoFit.mode === "shrink" ? frame.autoFit.fontScale : 1;
  const justify = frame.verticalAlign === "middle" ? "center" : frame.verticalAlign === "bottom" ? "flex-end" : "flex-start";
  const rootStyle = [
    "width:100%",
    "height:100%",
    "box-sizing:border-box",
    "display:flex",
    ...(vertical ? [`writing-mode:vertical-${eastAsian || direction === "stackedVerticalRtl" ? "rl" : "lr"}`, `text-orientation:${stacked ? "upright" : "mixed"}`] : []),
    "flex-direction:column",
    `justify-content:${justify}`,
    `padding:${frame.margin.top}px ${frame.margin.right}px ${frame.margin.bottom}px ${frame.margin.left}px`,
    "overflow:visible",
    frame.wrap ? "overflow-wrap:anywhere" : "white-space:pre",
  ].join(";");
  const content = paragraphs.map((paragraph) => {
    const style = paragraph.style;
    const paragraphStyle = [
      "margin:0",
      `padding-${vertical ? "inline-start" : "left"}:${style.indent}px`,
      `text-indent:${-Math.min(style.hanging, style.indent)}px`,
      `line-height:${style.lineSpacing}`,
      `padding-${vertical ? "block-start" : "top"}:${style.spaceBefore}px`,
      `padding-${vertical ? "block-end" : "bottom"}:${style.spaceAfter}px`,
      `text-align:${style.align}`,
      ...(style.direction === undefined ? [] : [`direction:${style.direction}`]),
      frame.wrap ? "white-space:pre-wrap" : "white-space:pre",
    ].join(";");
    let lastRun: TextRun | undefined;
    const runs = paragraph.runs.map((sourceRun) => {
      const content = fontRuns(sourceRun, context.theme).map(run => {
        lastRun = run;
        return `<span lang="${escapeXml(run.style.language)}" style="${htmlRunStyle(run, scale, context, run.text)}">${escapeXml(run.text)}</span>`;
      }).join("");
      const action = actionAttributes(sourceRun.action, context);
      return `${action.open}${content}${action.close}`;
    }).join("");
    const bullet = bulletText(paragraph);
    const sourceBulletRun = bulletRun(paragraph);
    const selectedBulletRun = sourceBulletRun === undefined ? undefined : fontRuns(sourceBulletRun, context.theme)[0];
    const bulletPaint = bullet === "" || selectedBulletRun === undefined ? "" : ` style="display:inline-block;min-width:${style.hanging}px;white-space:normal;${htmlRunStyle(selectedBulletRun, scale, context, bullet)}"`;
    const sourceText = paragraph.runs.map(run => run.text).join("");
    // CSS otherwise collapses a wholly empty paragraph or its final empty line.
    const emptyLine = lastRun && (sourceText === "" || /[\r\n]$/u.test(sourceText)) ? `<span aria-hidden="true" style="${htmlRunStyle(lastRun, scale, context, "")}">\u200b</span>` : "";
    return `<div style="${paragraphStyle}">${bullet === "" ? "" : `<span aria-hidden="true"${bulletPaint}>${escapeXml(bullet)}</span>`}${runs}${emptyLine}</div>`;
  }).join("");
  return `<div xmlns="http://www.w3.org/1999/xhtml" style="${rootStyle}">${content}</div>`;
}
