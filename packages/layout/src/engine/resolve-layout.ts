import { normalizePresentation, type PresentationDocument } from "@pptkit/core";
import type { LayoutResult } from "../types/layout.js";
import { resolveNormalizedLayout } from "./resolve-normalized-layout.js";
import type { LayoutOptions } from "../types/analysis.js";

export function resolveLayout(document: PresentationDocument, options?: LayoutOptions): LayoutResult {
  return resolveNormalizedLayout(normalizePresentation(document), options);
}
