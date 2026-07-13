import { PresentationDocumentImpl } from "../document/presentation-document.js";
import type { PresentationDocument, PresentationInit } from "../types/presentation.js";

export function createPresentation(init: PresentationInit = {}): PresentationDocument {
  return new PresentationDocumentImpl(init);
}
