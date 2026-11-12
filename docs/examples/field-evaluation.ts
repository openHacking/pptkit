import { evaluatePresentationField } from "@pptkit/core";

export const pageNumber = evaluatePresentationField(
  { id: "number", kind: "slide-number" },
  { pageOrder: ["cover", "body"], pageId: "body", firstPageNumber: 1 },
);
