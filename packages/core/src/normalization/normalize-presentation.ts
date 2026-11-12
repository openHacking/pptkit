import { PresentationValidationError } from "../types/diagnostic.js";
import type { NormalizedPresentation, PresentationData } from "../types/presentation.js";
import { validatePresentation } from "../validation/validate-presentation.js";
import { PresentationNormalizer } from "./presentation-normalizer.js";
import type { PresentationInheritanceOrigin, PresentationInheritanceResolution } from "../types/inheritance.js";
import { deepFreeze } from "../utils/clone.js";

export function normalizePresentation(document: PresentationData): NormalizedPresentation {
  assertValid(document);
  return new PresentationNormalizer().normalize(document);
}

function assertValid(document: PresentationData): void {
  const diagnostics = validatePresentation(document);
  const errors = diagnostics.filter((diagnostic) => diagnostic.severity === "error");
  if (errors.length > 0) throw new PresentationValidationError(errors);
}

/** Resolve once and return the ordinary IR with a separate immutable theme/background origin snapshot. */
export function resolvePresentationInheritance(document: PresentationData): PresentationInheritanceResolution {
  assertValid(document);
  const origins: PresentationInheritanceOrigin[] = [];
  const presentation = new PresentationNormalizer().normalize(document, origin => origins.push(origin));
  return Object.freeze({ presentation, origins: deepFreeze(origins) });
}
