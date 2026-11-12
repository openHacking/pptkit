import { inspectMissingDescriptions, type PresentationDocument } from '@pptkit/core';

export function reviewDescriptions(document: PresentationDocument) {
  const report = inspectMissingDescriptions(document);
  return report.findings.map(({ elementId, path, slideId, layoutId, masterId }) => ({ elementId, path, slideId, layoutId, masterId }));
}
