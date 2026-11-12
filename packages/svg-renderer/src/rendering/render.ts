import type { NormalizedAsset } from "@pptkit/core";
import type { LayoutResult } from "@pptkit/layout";
import type { SvgRenderOptions, SvgRenderResult } from "../types/public.js";
import { escapeXml, safeId } from "./escape.js";
import { backgroundRect, renderElement, type RenderContext } from "./element.js";
import { chargeSvgBytes } from "./byte-budget.js";
import { captureReadingOrders, readingOrderIds, readingOrderOwnership, wrapReadingOrder } from "./reading-order.js";
import { checkRenderAbort } from "./abort.js";
import { captureEmbeddedFonts } from "./embedded-fonts.js";
import type { BrowserRasterFontCoverageRequest } from "./browser-raster.js";

const invalidXmlCharacters = /[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/gu;

export function documentSlideTargets(layout: LayoutResult, indices: readonly number[], namespace: string | undefined): Map<string, string> {
  const prefix = namespace === undefined ? "" : `pptkit-ns-${Array.from(namespace, char => char.codePointAt(0)!.toString(16)).join("-")}-`;
  const targets = new Map(indices.map(index => [layout.slides[index]!.id, `${prefix}pptkit-link-slide-${index}`]));
  if (targets.size !== indices.length) throw new TypeError("Selected slide IDs must be unique for document links.");
  return targets;
}

export async function renderResolvedLayout(layout: LayoutResult, options: SvgRenderOptions = {}, fontCoverage?: (request: BrowserRasterFontCoverageRequest) => void, targetOverride?: ReadonlyMap<string, string>): Promise<SvgRenderResult> {
  const signal = options.signal;
  checkRenderAbort(signal);
  const embeddedFonts = captureEmbeddedFonts(options.embeddedFonts);
  const namespace = options.idNamespace;
  const byteLimit = options.maxSvgBytes;
  if (byteLimit !== undefined && (!Number.isSafeInteger(byteLimit) || byteLimit < 0)) throw new RangeError("maxSvgBytes must be a nonnegative safe integer.");
  const pngAnimationDiagnostics = options.pngAnimationDiagnostics === undefined ? "off" : options.pngAnimationDiagnostics;
  if (pngAnimationDiagnostics !== "off" && pngAnimationDiagnostics !== "warn") throw new TypeError("pngAnimationDiagnostics must be off or warn.");
  const portabilityDiagnostics = options.portabilityDiagnostics === undefined ? "off" : options.portabilityDiagnostics;
  if (portabilityDiagnostics !== "off" && portabilityDiagnostics !== "warn") throw new TypeError("portabilityDiagnostics must be off or warn.");
  const mixedTextMode = options.mixedTextMode === undefined ? "browser" : options.mixedTextMode;
  if (mixedTextMode !== "browser" && mixedTextMode !== "native-no-wrap" && mixedTextMode !== "native") throw new TypeError("mixedTextMode must be browser, native-no-wrap or native.");
  const tableMode = options.tableMode === undefined ? "browser" : options.tableMode;
  if (tableMode !== "browser" && tableMode !== "native") throw new TypeError("tableMode must be browser or native.");
  const backgroundMode = options.backgroundMode === undefined ? "slide" : options.backgroundMode;
  if (backgroundMode !== "slide" && backgroundMode !== "transparent") throw new TypeError("backgroundMode must be slide or transparent.");
  const slideActionMode = options.slideActionMode === undefined ? "metadata" : options.slideActionMode;
  if (slideActionMode !== "metadata" && slideActionMode !== "document-links") throw new TypeError("slideActionMode must be metadata or document-links.");
  if (namespace !== undefined && (typeof namespace !== "string" || namespace.length === 0)) {
    throw new TypeError("idNamespace must be a nonempty string when provided.");
  }
  const selection = options.slideIndices;
  let indices: number[];
  if (selection === undefined) indices = Array.from({ length: layout.slides.length }, (_, index) => index);
  else {
    if (!Array.isArray(selection)) throw new TypeError("slideIndices must be an array of distinct valid slide indices.");
    indices = Array.from(selection);
    if (indices.some(index => !Number.isSafeInteger(index) || index < 0 || index >= layout.slides.length) || new Set(indices).size !== indices.length) throw new RangeError("slideIndices must contain distinct valid slide indices.");
  }
  const orders = captureReadingOrders(layout, options.readingOrder);
  const selectedIds = new Set(indices.map(index => layout.slides[index]!.id));
  // Encode code points injectively so untrusted input cannot break SVG or collide
  // with a different namespace. The caller owns uniqueness across render results.
  const idPrefix = namespace === undefined ? "" : `pptkit-ns-${Array.from(namespace, char => char.codePointAt(0)!.toString(16)).join("-")}-`;
  const slideActionTargets = slideActionMode === "document-links" ? targetOverride ?? documentSlideTargets(layout, indices, namespace) : undefined;
  const warnings: SvgRenderResult["warnings"] = (layout.analysis?.diagnostics ?? []).filter(diagnostic => selection === undefined || diagnostic.slideId === undefined || selectedIds.has(diagnostic.slideId)).map(({ severity, bounds, available, ...warning }) => warning);
  const assets = new Map<string, NormalizedAsset>(layout.assets.map((asset) => [asset.id, asset]));
  const masters = new Map((layout.masters ?? []).map(master => [master.id, master]));
  const layouts = new Map(layout.layouts.map((item) => [item.id, item]));
  const slides: SvgRenderResult["slides"] = [];
  let svgBytes = 0;

  for (const index of indices) {
    checkRenderAbort(signal);
    const slide = layout.slides[index]!;
    if (embeddedFonts.faces.length > 0) warnings.push({ code: "svg-embedded-font-unverified", slideId: slide.id, message: "Caller-supplied font programs were embedded; their validity, coverage and viewer support are not verified." });
    if (backgroundMode === "slide" && slide.background.type === "linearGradient") warnings.push({ code: "background-gradient-approximation", slideId: slide.id, path: `slides.${index}.background`, message: "SVG gradient color and opacity interpolation can differ from PowerPoint." });
    if (slide.eventTimeline !== undefined) warnings.push({ code: "event-timeline-static", slideId: slide.id, path: `slides.${index}.eventTimeline`, message: "Static SVG retains authored geometry; foreground timing is not played or embedded." });
    const slideLayout = layouts.get(slide.layoutId);
    if (slideLayout?.masterId !== undefined && !slideLayout.theme) throw new TypeError(`Layout "${slideLayout.id}" has no resolved master theme.`);
    const theme = slideLayout?.theme ?? layout.theme;
    const defs: string[] = embeddedFonts.style ? [embeddedFonts.style] : [];
    const context: RenderContext = {
      slideId: slide.id,
      ...(signal ? { signal } : {}),
      theme,
      assets,
      ...(options.resolveAsset === undefined ? {} : { resolveAsset: options.resolveAsset }),
      assetCache: new Map(),
      pngAnimationDiagnostics,
      portabilityDiagnostics,
      mixedTextMode,
      tableMode,
      ...(slideActionTargets === undefined ? {} : { slideActionTargets }),
      ...(fontCoverage === undefined ? {} : { fontCoverage }),
      warnings,
      defs,
      definitionIds: { prefix: `${idPrefix}pptkit-def-${index}-${Array.from(slide.id, char => char.codePointAt(0)!.toString(16)).join("-")}`, next: 0 },
    };
    const renderedElements: string[] = [];
    if (slide.showInheritedShapes !== false) {
      const master = slideLayout?.masterId === undefined ? undefined : masters.get(slideLayout.masterId);
      if (slideLayout?.showMasterShapes !== false) for (const element of master?.elements ?? []) renderedElements.push(await renderElement(element, { ...context, theme: master!.theme }));
      for (const element of slideLayout?.elements ?? []) renderedElements.push(await renderElement(element, context));
    }
    const order = orders.get(slide.id), groups = new Map<string, ReadonlyMap<string, string>>();
    let rootOrder: ReadonlyMap<string, string> | undefined;
    order?.forEach(({ groupId, elementIds }, index) => {
      const ids = readingOrderIds(elementIds, `${context.definitionIds.prefix}-reading-${index}`);
      if (groupId === undefined) rootOrder = ids;
      else groups.set(groupId, ids);
    });
    const localContext = groups.size ? { ...context, groupReadingOrders: groups } : context;
    for (const element of slide.elements) renderedElements.push(wrapReadingOrder(element, await renderElement(element, localContext), rootOrder));
    checkRenderAbort(signal);
    const body = renderedElements.join("");
    const title = `${layout.metadata.title || "PPTKit presentation"} — slide ${index + 1}`;
    const rootId = slideActionTargets?.get(slide.id) ?? `${idPrefix}${safeId(`pptkit-slide-${slide.id}`)}`;
    const accessibility = order ? `role="group"${readingOrderOwnership(rootOrder)}` : 'role="img"';
    const background = backgroundMode === "slide" ? backgroundRect({ x: 0, y: 0, width: layout.size.width, height: layout.size.height }, slide.background, theme, `${rootId}-background`) : "";
    let svg = `<svg id="${rootId}" xmlns="http://www.w3.org/2000/svg" xmlns:xhtml="http://www.w3.org/1999/xhtml" viewBox="0 0 ${layout.size.width} ${layout.size.height}" width="${layout.size.width}" height="${layout.size.height}" ${accessibility} aria-labelledby="${rootId}-title" data-pptkit-slide-id="${escapeXml(slide.id)}"><title id="${rootId}-title">${escapeXml(title)}</title><defs>${defs.join("")}</defs>${background}${body}</svg>`;
    invalidXmlCharacters.lastIndex = 0;
    if (invalidXmlCharacters.test(svg)) {
      warnings.push({ code: "svg-xml-character-replacement", slideId: slide.id, path: `slides.${index}`, message: "Characters forbidden by XML 1.0 were replaced with U+FFFD in this SVG slide." });
      svg = svg.replace(invalidXmlCharacters, "\ufffd");
    }
    if (byteLimit !== undefined) svgBytes = chargeSvgBytes(svg, svgBytes, byteLimit, index);
    slides.push({ slideId: slide.id, index, hidden: slide.hidden, svg });
  }

  checkRenderAbort(signal);
  return {
    width: layout.size.width,
    height: layout.size.height,
    slides,
    warnings,
    status: warnings.length === 0 ? "rendered" : "rendered-with-warnings",
  };
}
