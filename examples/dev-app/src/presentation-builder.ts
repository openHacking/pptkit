import { createPresentation, type Box, type PresentationDocument, type PresentationElementInput } from "@pptkit/core";
import type { ExampleElementSpec, ExampleInputData } from "./example-types.js";

function defaultBox(element: Exclude<ExampleElementSpec, string>, y: number): Box {
  if (element.type === "image") {
    return { x: 48, y, width: 320, height: 180 };
  }

  if (element.type === "shape") {
    if (element.shape === "line") {
      return { x: 48, y, width: 320, height: 4 };
    }

    return { x: 48, y, width: 240, height: 120 };
  }

  return { x: 48, y, width: 640, height: 24 };
}

function createSlideElements(
  presentation: PresentationDocument,
  elements: ExampleInputData["slides"][number]["elements"],
): PresentationElementInput[] {
  let nextY = 48;

  return elements.map((element) => {
    if (typeof element === "string") {
      const box = { x: 48, y: nextY, width: 640, height: 24 };
      nextY += box.height + 24;
      return { type: "text", text: element, box };
    }

    const box = element.box ?? defaultBox(element, nextY);
    nextY = box.y + box.height + 24;

    if (element.type === "image") {
      const asset = presentation.registerAsset({
        kind: "image",
        source: element.asset.source,
        ...(element.asset.id !== undefined ? { id: element.asset.id } : {}),
        ...(element.asset.mimeType !== undefined ? { mimeType: element.asset.mimeType } : {}),
        ...(element.asset.width !== undefined ? { width: element.asset.width } : {}),
        ...(element.asset.height !== undefined ? { height: element.asset.height } : {}),
        ...(element.asset.altText !== undefined ? { altText: element.asset.altText } : {}),
        ...(element.asset.dedupeKey !== undefined ? { dedupeKey: element.asset.dedupeKey } : {}),
      });

      return {
        type: "image",
        assetId: asset.id,
        box,
        ...(element.altText !== undefined || element.asset.altText !== undefined
          ? { altText: element.altText ?? element.asset.altText }
          : {}),
      };
    }

    if (element.type === "shape") {
      return {
        type: "shape",
        shape: element.shape,
        box,
        ...(element.style !== undefined ? { style: element.style } : {}),
      };
    }

    return {
      type: "text",
      text: element.text,
      box,
      ...(element.style !== undefined ? { style: element.style } : {}),
    };
  });
}

export function createExamplePresentation(input: ExampleInputData): PresentationDocument {
  const presentation = createPresentation({
    title: input.title,
    ...(input.size !== undefined ? { size: input.size } : {}),
  });

  for (const slide of input.slides) {
    presentation.addSlide({
      ...(slide.id !== undefined ? { id: slide.id } : {}),
      ...(slide.background !== undefined ? { background: slide.background } : {}),
      elements: createSlideElements(presentation, slide.elements),
    });
  }

  return presentation;
}
