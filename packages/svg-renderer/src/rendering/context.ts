import type { NormalizedAsset, NormalizedPresentationTheme } from "@pptkit/core";
import type { LayoutElement } from "@pptkit/layout";
import type { SvgAssetResolver, SvgRenderWarning } from "../types/public.js";
import { escapeXml } from "./escape.js";
import type { BrowserRasterFontCoverageRequest } from "./browser-raster.js";

export interface RenderContext {
  signal?: AbortSignal;
  slideId: string;
  /** Ancestor artwork is decorative; descendant actions stay out of sequential focus. */
  insideDecorativeGroup?: true;
  groupReadingOrders?: ReadonlyMap<string, ReadonlyMap<string, string>>;
  theme: NormalizedPresentationTheme;
  assets: ReadonlyMap<string, NormalizedAsset>;
  resolveAsset?: SvgAssetResolver;
  pngAnimationDiagnostics: "off" | "warn";
  portabilityDiagnostics: "off" | "warn";
  mixedTextMode: "browser" | "native-no-wrap" | "native";
  tableMode: "browser" | "native";
  slideActionTargets?: ReadonlyMap<string, string>;
  fontDependencies?: Map<string, { family: string; bold: boolean; italic: boolean }>;
  fontCoverage?: (request: BrowserRasterFontCoverageRequest) => void;
  assetCache: Map<string, Promise<string | undefined>>;
  warnings: SvgRenderWarning[];
  defs: string[];
  definitionIds: { prefix: string; next: number };
}

export function recordFont(context: RenderContext, family: string, bold: boolean, italic: boolean, text: string, fontSize: number, language?: string, elementId?: string): void {
  context.fontDependencies?.set(JSON.stringify([family, bold, italic]), { family, bold, italic });
  if (text && context.fontCoverage) context.fontCoverage({ family, bold, italic, text: text.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/gu, "\ufffd"), fontSize, slideId: context.slideId, ...(elementId === undefined ? {} : { elementId }), ...(language === undefined ? {} : { language }) });
}

/** Unique within a rendered presentation, including inherited and grouped objects. */
export function nextDefinitionId(context: RenderContext): string {
  return `${context.definitionIds.prefix}-${context.definitionIds.next++}`;
}

/** Keep element identity on the SVG frame even when its XHTML child owns semantics. */
export function accessibility(element: LayoutElement, labelOwner: "frame" | "child" = "frame"): string {
  const identity = ` data-pptkit-element-id="${escapeXml(element.id)}"`;
  if (element.accessibility.decorative) return `${identity} aria-hidden="true"`;
  if (labelOwner === "child") return identity;
  const role = element.type === "group" ? "group" : "img";
  return `${identity} role="${role}" aria-label="${escapeXml(element.accessibility.description ?? element.name)}"`;
}
