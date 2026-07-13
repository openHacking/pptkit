import { normalizePresentation, type PresentationDocument } from "@pptkit/core";
import type { LayoutResult } from "../types/layout.js";
import { resolveNormalizedLayout } from "./resolve-normalized-layout.js";

export function resolveLayout(document: PresentationDocument): LayoutResult {
  return resolveNormalizedLayout(normalizePresentation(document));
}
