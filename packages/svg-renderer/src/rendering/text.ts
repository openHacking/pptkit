import type {
  Box,
  ElementAction,
  NormalizedPresentationTheme,
  NormalizedTextFrameStyle,
  NormalizedTextParagraph,
} from "@pptkit/core";
import { escapeXml, safeActionUrl } from "./escape.js";
import { colorValue } from "./style.js";
import type { SvgRenderWarning } from "../types/public.js";

interface TextContext {
  theme: NormalizedPresentationTheme;
  warnings: SvgRenderWarning[];
  slideId: string;
  elementId: string;
}

type TextRun = NormalizedTextParagraph["runs"][number];

interface NativeTextSegment {
  run: TextRun;
  text: string;
}

interface NativeTextLine {
  first: boolean;
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

function paragraphLines(paragraph: NormalizedTextParagraph, availableWidth: number, scale: number): Omit<NativeTextLine, "y">[] {
  const firstRun = paragraph.runs[0];
  if (firstRun === undefined) return [];
  const bullet = bulletText(paragraph);
  const firstWidth = Math.max(1, availableWidth + paragraph.style.hanging);
  const laterWidth = Math.max(1, availableWidth);
  const lines: Omit<NativeTextLine, "y">[] = [];
  let first = true;
  let segments: NativeTextSegment[] = [];
  let width = 0;
  let hasContent = false;
  let pendingSpace: TextRun | undefined;

  const startLine = (): void => {
    segments = [];
    width = 0;
    hasContent = false;
    if (first && bullet !== "") {
      appendSegment(segments, firstRun, bullet);
      width = textWidth(bullet, firstRun, scale);
    }
  };
  const finishLine = (): void => {
    const lineFontSize = segments.reduce((largest, segment) => Math.max(largest, scaledFontSize(segment.run, scale)), scaledFontSize(firstRun, scale));
    lines.push({ first, fontSize: lineFontSize, lineHeight: lineFontSize * paragraph.style.lineSpacing, paragraph, segments });
    first = false;
    pendingSpace = undefined;
    startLine();
  };
  const appendWord = (word: string, run: TextRun): void => {
    const space = pendingSpace === undefined || segments.length === 0 ? "" : " ";
    const maximum = first ? firstWidth : laterWidth;
    const candidateWidth = width + textWidth(space, pendingSpace ?? run, scale) + textWidth(word, run, scale);
    if (hasContent && candidateWidth > maximum) finishLine();
    if (pendingSpace !== undefined && segments.length > 0) {
      appendSegment(segments, pendingSpace, " ");
      width += textWidth(" ", pendingSpace, scale);
    }
    for (const character of word) {
      const characterMeasure = characterWidth(character, run, scale);
      const activeMaximum = first ? firstWidth : laterWidth;
      if (hasContent && width + characterMeasure > activeMaximum) finishLine();
      appendSegment(segments, run, character);
      width += characterMeasure;
      hasContent = true;
    }
    pendingSpace = undefined;
  };

  startLine();
  for (const run of paragraph.runs) {
    for (const token of run.text.split(/(\s+)/u)) {
      if (token === "") continue;
      if (/\s/u.test(token)) {
        if (token.includes("\n")) finishLine();
        else pendingSpace = run;
      } else appendWord(token, run);
    }
  }
  if (segments.length > 0 || lines.length === 0) finishLine();
  return lines;
}

export function canRenderTextNatively(paragraphs: NormalizedTextParagraph[]): boolean {
  // Mixed runs with independent actions stay on the XHTML path until Layout
  // owns resolved cross-run line fragments as a stable contract.
  return paragraphs.every((paragraph) => paragraph.runs.length === 1 && paragraph.runs[0]?.action === undefined);
}

function textAnchor(align: NormalizedTextParagraph["style"]["align"]): "start" | "middle" | "end" {
  if (align === "center") return "middle";
  if (align === "right") return "end";
  return "start";
}

function nativeRunAttributes(run: TextRun, scale: number, theme: NormalizedPresentationTheme): string {
  const decorations = [run.style.underline ? "underline" : "", run.style.strike ? "line-through" : ""].filter(Boolean).join(" ");
  return [
    `font-family="${escapeXml(fontFamily(run.style.fontFamily, theme))}"`,
    `font-size="${formatNumber(scaledFontSize(run, scale))}"`,
    `font-weight="${run.style.bold ? 700 : 400}"`,
    `font-style="${run.style.italic ? "italic" : "normal"}"`,
    `text-decoration="${decorations || "none"}"`,
    `fill="${colorValue(run.style.color, theme)}"`,
    `lang="${escapeXml(run.style.language)}"`,
  ].join(" ");
}

export function textSvg(
  paragraphs: NormalizedTextParagraph[],
  frame: NormalizedTextFrameStyle,
  box: Box,
  context: TextContext,
): string {
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
  let cursor = 0;
  for (const paragraph of paragraphs) {
    cursor += paragraph.style.spaceBefore;
    const availableWidth = Math.max(1, contentWidth - paragraph.style.indent);
    for (const line of paragraphLines(paragraph, availableWidth, scale)) {
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
    const left = box.x + frame.margin.left + style.indent - (line.first ? style.hanging : 0);
    const right = box.x + box.width - frame.margin.right;
    const anchor = textAnchor(style.align);
    const x = anchor === "middle" ? (left + right) / 2 : anchor === "end" ? right : left;
    const y = box.y + frame.margin.top + verticalOffset + line.y + line.fontSize * TEXT_BASELINE_RATIO;
    const content = line.segments.map((segment) => `<tspan ${nativeRunAttributes(segment.run, scale, context.theme)}>${escapeXml(segment.text)}</tspan>`).join("");
    return `<text x="${x}" y="${y}" text-anchor="${anchor}" data-pptkit-text-line="${lineIndex}">${content}</text>`;
  }).join("");
}

function actionAttributes(action: ElementAction | undefined, context: TextContext): { open: string; close: string } {
  if (action === undefined) return { open: "", close: "" };
  if (action.type === "slide") {
    context.warnings.push({
      code: "slide-action-not-linked",
      message: `Slide action targeting ${action.slideId} is metadata-only in standalone SVG output.`,
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
    open: `<a href="${escapeXml(url)}" rel="noopener noreferrer"${action.tooltip === undefined ? "" : ` title="${escapeXml(action.tooltip)}"`}>`,
    close: "</a>",
  };
}

function fontFamily(value: NormalizedTextParagraph["runs"][number]["style"]["fontFamily"], theme: NormalizedPresentationTheme): string {
  return typeof value === "string" ? value : theme.fonts[value.theme];
}

function cssFontFamily(value: string): string {
  return `&quot;${escapeXml(value.replace(/\\/g, "\\\\").replace(/"/g, '\\"'))}&quot;`;
}

function bulletText(paragraph: NormalizedTextParagraph): string {
  const bullet = paragraph.style.bullet;
  if (bullet.type === "none") return "";
  if (bullet.type === "bullet") return `${bullet.character} `;
  const value = bullet.startAt;
  if (bullet.style === "arabicParen") return `${value}) `;
  if (bullet.style === "alphaLowerPeriod") return `${String.fromCharCode(96 + Math.min(26, value))}. `;
  if (bullet.style === "alphaUpperPeriod") return `${String.fromCharCode(64 + Math.min(26, value))}. `;
  return `${value}. `;
}

export function textHtml(
  paragraphs: NormalizedTextParagraph[],
  frame: NormalizedTextFrameStyle,
  context: TextContext,
): string {
  if (frame.autoFit.mode === "resize") {
    context.warnings.push({
      code: "text-autofit-resize-unsupported",
      message: "Text autoFit resize cannot change SVG element geometry and was rendered inside the authored box.",
      slideId: context.slideId,
      elementId: context.elementId,
    });
  }
  const scale = frame.autoFit.mode === "shrink" ? frame.autoFit.fontScale : 1;
  const justify = frame.verticalAlign === "middle" ? "center" : frame.verticalAlign === "bottom" ? "flex-end" : "flex-start";
  const rootStyle = [
    "width:100%",
    "height:100%",
    "box-sizing:border-box",
    "display:flex",
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
      `padding-left:${style.indent}px`,
      `text-indent:${-style.hanging}px`,
      `line-height:${style.lineSpacing}`,
      `padding-top:${style.spaceBefore}px`,
      `padding-bottom:${style.spaceAfter}px`,
      `text-align:${style.align}`,
      frame.wrap ? "white-space:pre-wrap" : "white-space:pre",
    ].join(";");
    const runs = paragraph.runs.map((run) => {
      const decorations = [run.style.underline ? "underline" : "", run.style.strike ? "line-through" : ""].filter(Boolean).join(" ");
      const runStyle = [
        `font-family:${cssFontFamily(fontFamily(run.style.fontFamily, context.theme))}`,
        `font-size:${formatNumber(run.style.fontSize * scale)}px`,
        `font-weight:${run.style.bold ? 700 : 400}`,
        `font-style:${run.style.italic ? "italic" : "normal"}`,
        `text-decoration:${decorations || "none"}`,
        `color:${colorValue(run.style.color, context.theme)}`,
      ].join(";");
      const action = actionAttributes(run.action, context);
      return `${action.open}<span lang="${escapeXml(run.style.language)}" style="${runStyle}">${escapeXml(run.text)}</span>${action.close}`;
    }).join("");
    const bullet = bulletText(paragraph);
    return `<div style="${paragraphStyle}">${bullet === "" ? "" : `<span aria-hidden="true">${escapeXml(bullet)}</span>`}${runs}</div>`;
  }).join("");
  return `<div xmlns="http://www.w3.org/1999/xhtml" style="${rootStyle}">${content}</div>`;
}
