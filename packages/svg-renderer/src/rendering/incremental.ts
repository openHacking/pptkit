import type { LayoutResult } from "@pptkit/layout";
import type { SvgRenderOptions, SvgRenderResult, SvgRenderWarning, SvgRenderedSlide } from "../types/public.js";
import { checkRenderAbort } from "./abort.js";
import { chargeSvgBytes } from "./byte-budget.js";
import { captureSvgPreviewSnapshot, planSvgPreviewInvalidation, type SvgPreviewSnapshot } from "./invalidation.js";
import { captureReadingOrders } from "./reading-order.js";
import { documentSlideTargets, renderResolvedLayout } from "./render.js";
import { captureEmbeddedFonts } from "./embedded-fonts.js";

export interface IncrementalSvgRenderOptions extends Omit<SvgRenderOptions, "slideIndices"> {
  /** Change when external resolver results or other caller-owned dependencies change. */
  environmentRevision: string;
}
export interface IncrementalSvgCache {
  readonly schemaVersion: "pptkit-svg-incremental-cache-v1";
  readonly snapshot: SvgPreviewSnapshot;
  readonly entries: readonly { readonly slide: Readonly<SvgRenderedSlide>; readonly warnings: readonly Readonly<SvgRenderWarning>[] }[];
}
export interface IncrementalSvgResult {
  readonly result: SvgRenderResult;
  readonly cache: IncrementalSvgCache;
  readonly renderedSlideIds: readonly string[];
}

function copyWarning(warning: Readonly<SvgRenderWarning>): SvgRenderWarning {
  return { ...warning, ...(warning.fontDependencies === undefined ? {} : { fontDependencies: warning.fontDependencies.map(font => ({ ...font })) }) };
}

function freezeWarning(warning: SvgRenderWarning): Readonly<SvgRenderWarning> {
  const copy = copyWarning(warning);
  if (copy.fontDependencies !== undefined) {
    copy.fontDependencies.forEach(font => Object.freeze(font));
    Object.freeze(copy.fontDependencies);
  }
  return Object.freeze(copy);
}

/** Render all pages, reusing a previously returned cache when captured dependencies match. */
export async function renderIncrementalSvg(layout: LayoutResult, options: IncrementalSvgRenderOptions, previous?: IncrementalSvgCache): Promise<IncrementalSvgResult> {
  if (!options || "slideIndices" in options) throw new TypeError("Incremental SVG rendering requires options and renders all slides.");
  const { signal, maxSvgBytes, idNamespace, resolveAsset, environmentRevision } = options;
  checkRenderAbort(signal);
  const embeddedFonts = captureEmbeddedFonts(options.embeddedFonts).faces;
  if (maxSvgBytes !== undefined && (!Number.isSafeInteger(maxSvgBytes) || maxSvgBytes < 0)) throw new RangeError("maxSvgBytes must be a nonnegative safe integer.");
  const diagnostics = options.pngAnimationDiagnostics === undefined ? "off" : options.pngAnimationDiagnostics;
  if (diagnostics !== "off" && diagnostics !== "warn") throw new TypeError("pngAnimationDiagnostics must be off or warn.");
  const portability = options.portabilityDiagnostics === undefined ? "off" : options.portabilityDiagnostics;
  if (portability !== "off" && portability !== "warn") throw new TypeError("portabilityDiagnostics must be off or warn.");
  const mixedTextMode = options.mixedTextMode === undefined ? "browser" : options.mixedTextMode;
  if (mixedTextMode !== "browser" && mixedTextMode !== "native-no-wrap" && mixedTextMode !== "native") throw new TypeError("mixedTextMode must be browser, native-no-wrap or native.");
  const tableMode = options.tableMode === undefined ? "browser" : options.tableMode;
  if (tableMode !== "browser" && tableMode !== "native") throw new TypeError("tableMode must be browser or native.");
  const backgroundMode = options.backgroundMode === undefined ? "slide" : options.backgroundMode;
  if (backgroundMode !== "slide" && backgroundMode !== "transparent") throw new TypeError("backgroundMode must be slide or transparent.");
  const slideActionMode = options.slideActionMode === undefined ? "metadata" : options.slideActionMode;
  if (slideActionMode !== "metadata" && slideActionMode !== "document-links") throw new TypeError("slideActionMode must be metadata or document-links.");
  if (resolveAsset !== undefined && typeof resolveAsset !== "function") throw new TypeError("resolveAsset must be a function when provided.");
  if (typeof environmentRevision !== "string" || !environmentRevision) throw new TypeError("environmentRevision must be a nonempty string.");
  if (previous && previous.schemaVersion !== "pptkit-svg-incremental-cache-v1") throw new TypeError("Unsupported incremental SVG cache.");
  const working = structuredClone(layout);
  const targets = slideActionMode === "document-links" ? documentSlideTargets(working, working.slides.map((_, index) => index), idNamespace) : undefined;
  const orders = captureReadingOrders(working, options.readingOrder);
  const readingOrder = [...orders.values()].flat();
  const snapshotOptions = { readingOrder, environmentRevision: JSON.stringify([environmentRevision, diagnostics, portability, mixedTextMode, tableMode, backgroundMode, embeddedFonts, slideActionMode, targets === undefined ? null : [...targets]]), ...(idNamespace === undefined ? {} : { idNamespace }) };
  const snapshot = captureSvgPreviewSnapshot(working, snapshotOptions);
  const plan = previous ? planSvgPreviewInvalidation(previous.snapshot, snapshot) : undefined;
  const reusable = new Set(plan?.reusable.map(item => item.index) ?? []);
  const entries: { slide: Readonly<SvgRenderedSlide>; warnings: readonly Readonly<SvgRenderWarning>[] }[] = [];
  const renderedSlideIds: string[] = [];
  const baseWarnings = (working.analysis?.diagnostics ?? []).map(({ severity, bounds, available, ...warning }) => warning);
  const withoutAnalysis = { ...working }; delete withoutAnalysis.analysis;
  let bytes = 0;
  for (let index = 0; index < working.slides.length; index++) {
    checkRenderAbort(signal);
    let entry = reusable.has(index) ? previous?.entries[index] : undefined;
    if (entry && (entry.slide.index !== index || entry.slide.slideId !== working.slides[index]!.id)) throw new TypeError("Incremental SVG cache entry does not match its snapshot.");
    if (!entry) {
      const slideOrder = orders.get(working.slides[index]!.id);
      const single = await renderResolvedLayout(withoutAnalysis, { slideIndices: [index], readingOrder: slideOrder ?? [], ...(idNamespace === undefined ? {} : { idNamespace }), ...(resolveAsset === undefined ? {} : { resolveAsset }), ...(signal === undefined ? {} : { signal }), pngAnimationDiagnostics: diagnostics, portabilityDiagnostics: portability, mixedTextMode, tableMode, backgroundMode, embeddedFonts, slideActionMode }, undefined, targets);
      entry = Object.freeze({ slide: Object.freeze({ ...single.slides[0]! }), warnings: Object.freeze(single.warnings.map(freezeWarning)) });
      renderedSlideIds.push(working.slides[index]!.id);
    }
    if (maxSvgBytes !== undefined) bytes = chargeSvgBytes(entry.slide.svg, bytes, maxSvgBytes, index);
    entries.push(entry);
  }
  checkRenderAbort(signal);
  const after = captureSvgPreviewSnapshot(working, snapshotOptions);
  const finalPlan = planSvgPreviewInvalidation(snapshot, after);
  if (finalPlan.render.length || finalPlan.removedSlideIds.length) throw new TypeError("A resolver mutated the captured Layout during incremental rendering.");
  const warnings: SvgRenderWarning[] = [...baseWarnings, ...entries.flatMap(entry => entry.warnings.map(copyWarning))];
  const result: SvgRenderResult = { width: working.size.width, height: working.size.height, slides: entries.map(entry => ({ ...entry.slide })), warnings, status: warnings.length ? "rendered-with-warnings" : "rendered" };
  const cache = Object.freeze({ schemaVersion: "pptkit-svg-incremental-cache-v1" as const, snapshot, entries: Object.freeze(entries) });
  return Object.freeze({ result, cache, renderedSlideIds: Object.freeze(renderedSlideIds) });
}
