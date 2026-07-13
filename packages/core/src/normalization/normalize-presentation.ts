import { PresentationValidationError } from "../types/diagnostic.js";
import type { NormalizedPresentation, PresentationDocument } from "../types/presentation.js";
import { validatePresentation } from "../validation/validate-presentation.js";
import { PresentationNormalizer } from "./presentation-normalizer.js";

export function normalizePresentation(document: PresentationDocument): NormalizedPresentation {
  const diagnostics = validatePresentation(document);
  const errors = diagnostics.filter((diagnostic) => diagnostic.severity === "error");
  if (errors.length > 0) throw new PresentationValidationError(errors);
  return new PresentationNormalizer().normalize(document);
}
