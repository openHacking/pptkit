import { importSimpleSvg, SvgImportError, type SvgImportResult } from "./simple-svg.js";
import { createSvgRasterFallback, type SvgRasterFallback, type SvgRasterFallbackInput } from "./raster-fallback.js";

export interface SvgImportWithFallbackOptions {
  /** Caller-owned identity; it is not a hash and is not verified against the source. */
  readonly sourceId: string;
  /** Explicitly authorizes use of these trusted raster bytes if conversion fails. */
  readonly fallback?: Omit<SvgRasterFallbackInput, "sourceId">;
}
export type SvgImportFailure = Extract<SvgImportResult, { status: "unsupported" }> | { readonly status: "parse-error"; readonly code: "invalid-xml" | "input-limit"; readonly message: string };
export type SvgImportWithFallbackResult =
  | (Extract<SvgImportResult, { status: "editable" }> & { readonly sourceId: string })
  | (SvgImportFailure & { readonly sourceId: string })
  | (SvgRasterFallback & { readonly reason: SvgImportFailure });

/** Uses a caller-supplied fallback only after a reported unsupported or parse-limit outcome. */
export function importSvgWithFallback(source: string, options: SvgImportWithFallbackOptions): SvgImportWithFallbackResult {
  if (!options || typeof options.sourceId !== "string" || !options.sourceId.trim() || options.sourceId.length > 1024) throw new TypeError("A nonempty SVG source identity is required.");
  let result: SvgImportResult;
  let failure: SvgImportFailure;
  try {
    result = importSimpleSvg(source);
    if (result.status === "editable") return { ...result, sourceId: options.sourceId };
    failure = result;
  } catch (error) {
    if (!(error instanceof SvgImportError) || error.code === "invalid-input") throw error;
    failure = { status: "parse-error", code: error.code, message: error.message };
  }
  if (options.fallback === undefined) return { ...failure, sourceId: options.sourceId };
  return { ...createSvgRasterFallback({ ...options.fallback, sourceId: options.sourceId }), reason: failure };
}
