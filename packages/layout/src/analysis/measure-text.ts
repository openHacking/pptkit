import type { Box, NormalizedPresentationTheme, NormalizedTextFrameStyle, NormalizedTextParagraph } from "@pptkit/core";
import { resolveTextFontSegments } from "@pptkit/core";
import type { LayoutDiagnostic, MeasurementParagraph, TextMeasurementOptions, TextMeasurementRequest } from "../types/analysis.js";

export interface MeasurementContext {
  elementId: string;
  path: string;
  slideId?: string;
  layoutId?: string;
  masterId?: string;
}

function freeze<T>(value: T): T {
  if (value !== null && typeof value === "object") {
    for (const child of Object.values(value)) freeze(child);
    Object.freeze(value);
  }
  return value;
}

function requireSynchronous(value: unknown): void {
  if (value !== null && (typeof value === "object" || typeof value === "function") && "then" in value && typeof value.then === "function") {
    // A JavaScript caller can bypass the synchronous TypeScript contract. Consume
    // rejected promises so an invalid provider cannot cause an unhandled rejection.
    void Promise.resolve(value).catch(() => undefined);
    throw new TypeError("Measurement and font providers must be synchronous; load resources before resolving Layout");
  }
}

function validBounds(value: unknown): value is Box {
  if (value === null || typeof value !== "object") return false;
  const box = value as Box;
  return [box.x, box.y, box.width, box.height, box.x + box.width, box.y + box.height].every(Number.isFinite)
    && box.width >= 0 && box.height >= 0;
}

export function measureTextBody(
  content: readonly NormalizedTextParagraph[], frame: NormalizedTextFrameStyle, contentBox: Box,
  theme: NormalizedPresentationTheme, options: TextMeasurementOptions,
  context: MeasurementContext, diagnostics: LayoutDiagnostic[],
): Box | undefined {
  const direction = frame.textDirection;
  const vertical = direction === "eastAsianVertical" || direction === "mongolianVertical" || direction === "stackedVertical" || direction === "stackedVerticalRtl" ? direction : undefined;
  const measureEastAsian = vertical === "eastAsianVertical" ? options.measureEastAsianVertical : undefined;
  const measureVertical = options.measureVertical;
  if (vertical && !measureEastAsian && !measureVertical) {
    diagnostics.push({ ...context, code: vertical === "eastAsianVertical" ? "east-asian-vertical-measurement-unavailable" : "vertical-text-measurement-unavailable", severity: "warning", message: "This vertical writing mode requires a vertical measurement provider." });
    return undefined;
  }
  if (frame.textDirection === "rotate90" || frame.textDirection === "rotate270") {
    diagnostics.push({ ...context, code: "text-direction-sizing-unsupported", severity: "warning", message: "Automatic sizing does not support rotated text frames." });
    return undefined;
  }
  const scale = frame.autoFit.mode === "shrink" ? frame.autoFit.fontScale : 1;
  const paragraphs: MeasurementParagraph[] = [];
  let fontFailed = false;
  for (const [paragraphIndex, paragraph] of content.entries()) {
    const runs = [];
    for (const [runIndex, run] of paragraph.runs.entries()) {
      let segments;
      try { segments = resolveTextFontSegments(run.text, run.style, theme); }
      catch (error) {
        diagnostics.push({ ...context, path: `${context.path}.${paragraphIndex}.runs.${runIndex}`, code: "text-font-selection-failed", severity: "warning", message: `Font selection failed: ${error instanceof Error ? error.message : String(error)}` });
        return undefined;
      }
      for (const segment of segments) {
        const family = segment.family;
        let resolved = family;
        const field = segment.scriptFont !== undefined ? `scriptFontFamilies.${segment.scriptFont}` : run.style.rightToLeft ? (run.style.complexScriptFontFamily === undefined ? "fontFamily" : "complexScriptFontFamily") : segment.script === "eastAsian" && run.style.eastAsianFontFamily !== undefined ? "eastAsianFontFamily" : segment.script === "complexScript" && run.style.complexScriptFontFamily !== undefined ? "complexScriptFontFamily" : "fontFamily";
        const path = `${context.path}.${paragraphIndex}.runs.${runIndex}.style.${field}`;
        if (options.resolveFont !== undefined) {
          try {
            const request = freeze({ family, fontSize: run.style.fontSize * scale, bold: run.style.bold, italic: run.style.italic, language: run.style.language, text: segment.text });
            const result = options.resolveFont(request);
            requireSynchronous(result);
            if (!result || typeof result.family !== "string" || !result.family.trim()) {
              fontFailed = true;
              diagnostics.push({ ...context, path, code: "font-unavailable", severity: "warning", message: `Font "${family}" is unavailable. Supply a font resolver with a supported fallback.` });
            } else {
              resolved = result.family;
              if (resolved !== family) diagnostics.push({ ...context, path, code: "font-fallback", severity: "warning", message: `Font "${family}" was measured using "${resolved}". Native applications may substitute a different font.` });
            }
          } catch (error) {
            fontFailed = true;
            diagnostics.push({ ...context, path, code: "font-resolution-failed", severity: "warning", message: `Font resolution failed: ${error instanceof Error ? error.message : String(error)}` });
          }
        }
        const { eastAsianFontFamily: _eastAsian, complexScriptFontFamily: _complexScript, scriptFontFamilies: _scriptMap, ...style } = structuredClone(run.style);
        runs.push({ text: segment.text, style: { ...style, fontFamily: resolved, fontSize: run.style.fontSize * scale } });
      }
    }
    paragraphs.push({ runs, style: structuredClone(paragraph.style) });
  }
  if (fontFailed) return undefined;
  const request: TextMeasurementRequest = freeze({ paragraphs, availableWidth: contentBox.width, wrap: frame.wrap });
  try {
    const logical = { paragraphs, availableInlineSize: contentBox.height, availableBlockSize: contentBox.width, wrap: frame.wrap };
    const measured = measureEastAsian ? measureEastAsian(freeze(logical))
      : vertical && measureVertical ? measureVertical(freeze({ ...logical, textDirection: vertical }))
        : options.measure(request);
    requireSynchronous(measured);
    if (!measured || !validBounds(measured.bounds)) {
      diagnostics.push({ ...context, code: "text-measurement-invalid", severity: "warning", message: "Text measurer must return finite bounds with nonnegative dimensions. Authored geometry was retained." });
      return undefined;
    }
    const bounds = { ...measured.bounds };
    const remaining = Math.max(0, (vertical ? contentBox.width : contentBox.height) - (bounds.y + bounds.height));
    const offset = frame.verticalAlign === "middle" ? remaining / 2 : frame.verticalAlign === "bottom" ? remaining : 0;
    if (vertical) {
      const rightToLeft = vertical === "eastAsianVertical" || vertical === "stackedVerticalRtl";
      return { x: contentBox.x + (rightToLeft ? contentBox.width - (offset + bounds.y + bounds.height) : offset + bounds.y),
      y: contentBox.y + bounds.x, width: bounds.height, height: bounds.width };
    }
    return { x: contentBox.x + bounds.x, y: contentBox.y + offset + bounds.y, width: bounds.width, height: bounds.height };
  } catch (error) {
    diagnostics.push({ ...context, code: "text-measurement-failed", severity: "warning", message: `Text measurement failed: ${error instanceof Error ? error.message : String(error)}. Authored geometry was retained.` });
    return undefined;
  }
}
