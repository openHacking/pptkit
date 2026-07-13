import { createPresentation, type PresentationDocument } from "@pptkit/core";
import type { ExampleInputData } from "./example-types.js";

export function createExamplePresentation(input: ExampleInputData): PresentationDocument {
  const presentation = createPresentation({ title: input.title });

  for (const slide of input.slides) {
    presentation.addSlide({
      ...(slide.id !== undefined ? { id: slide.id } : {}),
      elements: slide.elements.map((element, index) => ({
        type: "text",
        text: element,
        box: { x: 48, y: 48 + index * 32, width: 640, height: 24 },
      })),
    });
  }

  return presentation;
}
