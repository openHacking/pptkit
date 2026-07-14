import type { NormalizedAsset } from "@pptkit/core";
import type { LayoutResult } from "@pptkit/layout";
import type { SvgRenderOptions, SvgRenderResult } from "../types/public.js";
import { escapeXml, safeId } from "./escape.js";
import { backgroundRect, renderElement, type RenderContext } from "./element.js";

export async function renderResolvedLayout(layout: LayoutResult, options: SvgRenderOptions = {}): Promise<SvgRenderResult> {
  const warnings: SvgRenderResult["warnings"] = [];
  const assets = new Map<string, NormalizedAsset>(layout.assets.map((asset) => [asset.id, asset]));
  const layouts = new Map(layout.layouts.map((item) => [item.id, item]));
  const slides: SvgRenderResult["slides"] = [];

  for (let index = 0; index < layout.slides.length; index += 1) {
    const slide = layout.slides[index]!;
    const defs: string[] = [];
    const context: RenderContext = {
      slideId: slide.id,
      theme: layout.theme,
      assets,
      ...(options.resolveAsset === undefined ? {} : { resolveAsset: options.resolveAsset }),
      assetCache: new Map(),
      warnings,
      defs,
    };
    const slideLayout = layouts.get(slide.layoutId);
    const elements = [...(slideLayout?.elements ?? []), ...slide.elements];
    const renderedElements: string[] = [];
    for (const element of elements) renderedElements.push(await renderElement(element, context));
    const body = renderedElements.join("");
    const title = `${layout.metadata.title || "PPTKit presentation"} — slide ${index + 1}`;
    const rootId = safeId(`pptkit-slide-${slide.id}`);
    const svg = `<svg id="${rootId}" xmlns="http://www.w3.org/2000/svg" xmlns:xhtml="http://www.w3.org/1999/xhtml" viewBox="0 0 ${layout.size.width} ${layout.size.height}" width="${layout.size.width}" height="${layout.size.height}" role="img" aria-labelledby="${rootId}-title" data-pptkit-slide-id="${escapeXml(slide.id)}"><title id="${rootId}-title">${escapeXml(title)}</title><defs>${defs.join("")}</defs>${backgroundRect({ x: 0, y: 0, width: layout.size.width, height: layout.size.height }, slide.background, layout.theme)}${body}</svg>`;
    slides.push({ slideId: slide.id, index, hidden: slide.hidden, svg });
  }

  return {
    width: layout.size.width,
    height: layout.size.height,
    slides,
    warnings,
    status: warnings.length === 0 ? "rendered" : "rendered-with-warnings",
  };
}
