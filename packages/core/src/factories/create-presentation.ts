import { PresentationDocumentImpl } from "../document/presentation-document.js";
import type { ManagedPresentationDocument, PresentationInit } from "../types/presentation.js";

export function createPresentation(init: PresentationInit = {}): ManagedPresentationDocument {
  return new PresentationDocumentImpl(init);
}
