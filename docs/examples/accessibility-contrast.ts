import { inspectSlideTextContrast, type PresentationDocument } from '@pptkit/core';

export function reviewTextContrast(document: PresentationDocument) {
  return inspectSlideTextContrast(document).findings;
}
