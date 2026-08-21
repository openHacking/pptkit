import type { NormalizedAsset, NormalizedPresentationTheme } from "@pptkit/core";
import type { LayoutElement } from "@pptkit/layout";
import type { SvgAssetResolver, SvgRenderWarning } from "../types/public.js";
import { escapeXml } from "./escape.js";

export interface RenderContext {
  slideId: string;
  theme: NormalizedPresentationTheme;
  assets: ReadonlyMap<string, NormalizedAsset>;
  resolveAsset?: SvgAssetResolver;
  assetCache: Map<string, Promise<string | undefined>>;
  warnings: SvgRenderWarning[];
  defs: string[];
}

export function accessibility(element: LayoutElement): string {
  const identity = ` data-pptkit-element-id="${escapeXml(element.id)}"`;
  if (element.accessibility.decorative) return `${identity} aria-hidden="true"`;
  return `${identity} role="img" aria-label="${escapeXml(element.accessibility.description ?? element.name)}"`;
}
