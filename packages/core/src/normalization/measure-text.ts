import type { NormalizedTextParagraph } from "../types/element.js";
import type { NormalizedTextFrameStyle, NormalizedTextRunStyle } from "../types/style.js";

const DEFAULT_FONT_SIZE = 18;
const MIN_MEASURABLE_WIDTH = 1;
const MIN_TEXT_HEIGHT_SAFETY_PADDING = 3;
const TEXT_HEIGHT_SAFETY_PADDING_RATIO = 0.2;

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

function characterWidth(character: string, style: NormalizedTextRunStyle): number {
  if (/\s/u.test(character)) return style.fontSize * 0.28;
  if (isWideCharacter(character)) return style.fontSize;
  if (/[,.;:!?()[\]{}'"`]/u.test(character)) return style.fontSize * 0.35;
  return style.fontSize * (style.bold ? 0.56 : 0.52) * (style.italic ? 1.02 : 1);
}

function estimateRunWidth(text: string, style: NormalizedTextRunStyle): number {
  let width = 0;
  for (const character of text) width += characterWidth(character, style);
  return width;
}

function paragraphFontSize(paragraph: NormalizedTextParagraph): number {
  return paragraph.runs.reduce((largest, run) => Math.max(largest, run.style.fontSize), DEFAULT_FONT_SIZE);
}

function contentFontSize(paragraphs: readonly NormalizedTextParagraph[]): number {
  return paragraphs.reduce(
    (largest, paragraph) => Math.max(largest, paragraphFontSize(paragraph)),
    DEFAULT_FONT_SIZE,
  );
}

function paragraphTextWidth(paragraph: NormalizedTextParagraph): number {
  return paragraph.runs.reduce((width, run) => width + estimateRunWidth(run.text, run.style), 0);
}

function paragraphLineWidth(paragraph: NormalizedTextParagraph, availableWidth: number): number {
  const fontSize = paragraphFontSize(paragraph);
  const indentation = Math.max(0, paragraph.style.indent - paragraph.style.hanging);
  const bulletWidth = paragraph.style.bullet.type === "none" ? 0 : Math.max(9, fontSize * 0.75);
  return Math.max(MIN_MEASURABLE_WIDTH, availableWidth - indentation - bulletWidth);
}

function paragraphHeight(paragraph: NormalizedTextParagraph, availableWidth: number, wrap: boolean): number {
  const fontSize = paragraphFontSize(paragraph);
  const lineHeight = fontSize * paragraph.style.lineSpacing;
  const lineWidth = paragraphLineWidth(paragraph, availableWidth);
  const lineCount = wrap
    ? Math.max(1, Math.ceil(paragraphTextWidth(paragraph) / lineWidth))
    : 1;

  return (
    lineCount * lineHeight +
    paragraph.style.spaceBefore +
    paragraph.style.spaceAfter
  );
}

/**
 * Estimates the height of normalized text using deterministic font metrics.
 * This is intentionally environment-independent; it is not a replacement for
 * browser or native font measurement.
 */
export function estimateTextHeight(
  width: number,
  paragraphs: readonly NormalizedTextParagraph[],
  frame: NormalizedTextFrameStyle,
): number {
  const availableWidth = Math.max(
    MIN_MEASURABLE_WIDTH,
    width - frame.margin.left - frame.margin.right,
  );
  const contentHeight = paragraphs.length === 0
    ? DEFAULT_FONT_SIZE
    : paragraphs.reduce(
        (height, paragraph) => height + paragraphHeight(paragraph, availableWidth, frame.wrap),
        0,
      );
  // Font ascent/descent and line-box behavior vary by font and host renderer.
  // Leave a small trailing buffer so the final glyph is not clipped when the
  // deterministic estimate is slightly smaller than the renderer's metrics.
  const safetyPadding = Math.max(
    MIN_TEXT_HEIGHT_SAFETY_PADDING,
    contentFontSize(paragraphs) * TEXT_HEIGHT_SAFETY_PADDING_RATIO,
  );

  return frame.margin.top + contentHeight + frame.margin.bottom + safetyPadding;
}
