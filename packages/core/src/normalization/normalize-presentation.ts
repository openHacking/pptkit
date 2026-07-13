import type { NormalizedPresentation } from "../types/normalized.js";
import type { PresentationDocument } from "../types/presentation.js";
import { PresentationNormalizer } from "./presentation-normalizer.js";

export function normalizePresentation(document: PresentationDocument): NormalizedPresentation {
  return new PresentationNormalizer().normalize(document);
}
