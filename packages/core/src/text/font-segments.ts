import type { NormalizedTextRunStyle } from "../types/style.js";
import type { NormalizedPresentationTheme } from "../types/theme.js";
import { fontScriptBucket, fontScriptCode, fontScriptCandidates, hanFontScripts, type TextFontScript, type TextFontScriptCode } from "./font-scripts.js";

export type { TextFontScript } from "./font-scripts.js";
export interface TextFontSegment {
  readonly text: string;
  /** UTF-16 offsets into the original text; no rewriting or run identity changes. */
  readonly start: number;
  readonly end: number;
  readonly family: string;
  readonly script: TextFontScript;
  /** Present only when an explicit map entry supplied this family's value. */
  readonly scriptFont?: TextFontScriptCode;
}
export interface TextFontSegmentOptions {
  /** Positive safe integer, default 100000. Checked before adding a segment. */
  readonly maxSegments?: number;
}

const continuation = /[\p{Mark}\u200C\u200D\uFE00-\uFE0F]/u;
const neutral = /[\p{Punctuation}\p{Separator}]/u;
type FontStyle = Pick<NormalizedTextRunStyle, "fontFamily" | "eastAsianFontFamily" | "complexScriptFontFamily" | "scriptFontFamilies"> & Partial<Pick<NormalizedTextRunStyle, "language" | "rightToLeft">>;

/** Deterministic explicit-family selection; does not shape text or query host fonts. */
export function resolveTextFontSegments(text: string, style: FontStyle, theme: NormalizedPresentationTheme, options?: TextFontSegmentOptions): readonly TextFontSegment[] {
  const maximum = options?.maxSegments ?? 100_000;
  if (!Number.isSafeInteger(maximum) || maximum < 1) throw new RangeError("maxSegments must be a positive safe integer.");
  const family = (script: TextFontScript): string => {
    const value = (style.rightToLeft ? style.complexScriptFontFamily : script === "eastAsian" ? style.eastAsianFontFamily : script === "complexScript" ? style.complexScriptFontFamily : undefined) ?? style.fontFamily;
    return typeof value === "string" ? value : theme.fonts[value.theme];
  };
  const map = style.scriptFontFamilies, mapped = map !== undefined && Object.keys(map).length > 0;
  const han = mapped ? hanFontScripts(style.language) : [];
  // Preserve the historical single-family path, including empty insertion runs.
  if (!mapped && style.eastAsianFontFamily === undefined && style.complexScriptFontFamily === undefined) return [{ text, start: 0, end: text.length, family: family("default"), script: "default" }];
  const segments: TextFontSegment[] = [];
  let prior: TextFontScript = "default", offset = 0;
  let priorCode: TextFontScriptCode = "Zyyy";
  for (const character of text) {
    const newline = /[\r\n\u2028\u2029]/u.test(character);
    const specific = fontScriptBucket(character);
    const script: TextFontScript = mapped && newline ? "default" : continuation.test(character) && offset > 0 ? prior : specific !== "default" ? specific : neutral.test(character) ? prior : "default";
    const code: TextFontScriptCode = mapped && newline ? "Zyyy" : mapped && ((continuation.test(character) && offset > 0) || neutral.test(character)) ? priorCode : mapped ? fontScriptCode(character) : "Zyyy";
    const scriptFont = mapped ? fontScriptCandidates(code, han).find(candidate => map?.[candidate] !== undefined) : undefined;
    const value = scriptFont === undefined ? undefined : map?.[scriptFont];
    const selected = value === undefined ? family(script) : typeof value === "string" ? value : theme.fonts[value.theme];
    const previous = segments.at(-1), end = offset + character.length;
    if (previous && previous.script === script && previous.family === selected && previous.scriptFont === scriptFont) segments[segments.length - 1] = { ...previous, end };
    else {
      if (segments.length === maximum) throw new RangeError("Text font selection exceeds maxSegments.");
      segments.push({ text: "", start: offset, end, family: selected, script, ...(scriptFont === undefined ? {} : { scriptFont }) });
    }
    prior = newline ? "default" : script;
    priorCode = newline ? "Zyyy" : code;
    offset = end;
  }
  if (!segments.length) return [{ text: "", start: 0, end: 0, family: family("default"), script: "default" }];
  return segments.map(segment => ({ ...segment, text: text.slice(segment.start, segment.end) }));
}
