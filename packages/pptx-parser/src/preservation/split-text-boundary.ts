import { PackageInspectionError } from "../types/archive.js";

/** Private Unicode preflight only: it does not locate or write any source XML. */
export class TextSplitBoundaryError extends Error {
  constructor(readonly code: "split-invalid" | "split-limit-exceeded" | "split-unavailable", message: string) {
    super(message); this.name = "TextSplitBoundaryError";
  }
}

type SegmenterConstructor = new (locale: string | undefined, options: { granularity: "grapheme" }) => {
  segment(text: string): Iterable<{ index: number }>;
};

/** A UTF-16 offset must be an interior runtime grapheme boundary. */
export async function splitAtGraphemeBoundary(text: string, offset: number, options: {
  readonly maxCharacters?: number; readonly maxGraphemes?: number; readonly signal?: AbortSignal; readonly checkpoint?: () => void;
} = {}): Promise<Readonly<{ left: string; right: string }>> {
  if (typeof text !== "string" || !Number.isSafeInteger(offset) || !options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Invalid text split request.");
  const { maxCharacters = 2 * 1024 * 1024, maxGraphemes = 1000000, signal, checkpoint } = options;
  if (!Number.isSafeInteger(maxCharacters) || maxCharacters <= 0 || !Number.isSafeInteger(maxGraphemes) || maxGraphemes <= 0) throw new RangeError("Invalid text split limits.");
  if (checkpoint !== undefined && typeof checkpoint !== "function") throw new TypeError("Invalid text split checkpoint.");
  const check = () => { checkpoint?.(); if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Text split preflight was aborted."); checkpoint?.(); };
  check();
  if (text.length > maxCharacters) throw new TextSplitBoundaryError("split-limit-exceeded", "Text exceeds the split character limit.");
  if (offset <= 0 || offset >= text.length) throw new TextSplitBoundaryError("split-invalid", "Split offset must be inside the run.");
  let nextCheck = 8192;
  for (let i = 0; i < text.length; i++) {
    if (i >= nextCheck) { nextCheck = i + 8192; await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
    const value = text.charCodeAt(i);
    if (value >= 0xd800 && value <= 0xdbff) {
      const next = text.charCodeAt(++i);
      if (!(next >= 0xdc00 && next <= 0xdfff)) throw new TextSplitBoundaryError("split-invalid", "Text contains an unpaired surrogate.");
    } else if (value >= 0xdc00 && value <= 0xdfff) throw new TextSplitBoundaryError("split-invalid", "Text contains an unpaired surrogate.");
  }
  const Segmenter = (Intl as unknown as { Segmenter?: SegmenterConstructor }).Segmenter;
  if (typeof Segmenter !== "function") throw new TextSplitBoundaryError("split-unavailable", "Runtime grapheme segmentation is unavailable.");
  let count = 0;
  for (const segment of new Segmenter(undefined, { granularity: "grapheme" }).segment(text)) {
    if (++count > maxGraphemes) throw new TextSplitBoundaryError("split-limit-exceeded", "Text exceeds the split grapheme work limit.");
    if (count % 256 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
    if (segment.index === offset) { check(); return Object.freeze({ left: text.slice(0, offset), right: text.slice(offset) }); }
    if (segment.index > offset) break;
  }
  check();
  throw new TextSplitBoundaryError("split-invalid", "Split offset is not a grapheme boundary.");
}
