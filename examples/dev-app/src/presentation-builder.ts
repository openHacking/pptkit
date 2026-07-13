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
      return { type: "text", content: element, box };
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
        ...(element.asset.altText !== undefined ? { accessibility: { description: element.asset.altText } } : {}),
        ...(element.asset.dedupeKey !== undefined ? { dedupeKey: element.asset.dedupeKey } : {}),
      });

      const description = element.altText ?? element.asset.altText;
      return {
        type: "image",
        assetId: asset.id,
        box,
        ...(description !== undefined ? { accessibility: { description } } : {}),
      };
    }

    if (element.type === "shape") {
      const stroke = element.style?.stroke === undefined
        ? undefined
        : {
            paint: { type: "solid" as const, color: element.style.stroke },
            ...(element.style.strokeWidth !== undefined ? { width: element.style.strokeWidth } : {}),
          };
      if (element.shape === "line") {
        return {
          type: "connector",
          start: { x: box.x, y: box.y },
          end: { x: box.x + box.width, y: box.y + box.height },
          ...(stroke !== undefined ? { style: stroke } : {}),
        };
      }
      return {
        type: "shape",
        shape: element.shape,
        box,
        ...(element.style !== undefined ? {
          style: {
            ...(element.style.fill !== undefined ? { fill: { type: "solid" as const, color: element.style.fill } } : {}),
            ...(stroke !== undefined ? { stroke } : {}),
          },
        } : {}),
      };
    }

    return {
      type: "text",
      content: element.text,
      box,
      ...(element.style !== undefined ? {
        frame: { ...(element.style.autoFit !== undefined ? { autoFit: element.style.autoFit } : {}) },
      } : {}),
      ...(element.style !== undefined ? {
        content: [{
          style: {
            ...(element.style.align !== undefined ? { align: element.style.align } : {}),
            ...(element.style.lineSpacing !== undefined ? { lineSpacing: element.style.lineSpacing } : {}),
          },
          runs: [{
            text: element.text,
            style: {
              ...(element.style.fontSize !== undefined ? { fontSize: element.style.fontSize } : {}),
              ...(element.style.fontFamily !== undefined ? { fontFamily: element.style.fontFamily } : {}),
              ...(element.style.fontWeight !== undefined ? { bold: element.style.fontWeight === "bold" } : {}),
              ...(element.style.color !== undefined ? { color: element.style.color } : {}),
            },
          }],
        }],
      } : {}),
    };
  });
}

export function createExamplePresentation(input: ExampleInputData): PresentationDocument {
  const presentation = createPresentation({
    metadata: { title: input.title },
    ...(input.size !== undefined ? { size: input.size } : {}),
  });

  for (const slide of input.slides) {
    presentation.addSlide({
      ...(slide.id !== undefined ? { id: slide.id } : {}),
      ...(slide.background !== undefined ? { background: { type: "solid", color: slide.background } } : {}),
      elements: createSlideElements(presentation, slide.elements),
    });
  }

  return presentation;
}
