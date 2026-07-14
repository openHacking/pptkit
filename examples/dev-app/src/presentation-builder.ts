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

  if (element.type === "group") {
    return { x: 48, y, width: 320, height: 180 };
  }

  if (element.type === "table") {
    return { x: 48, y, width: 624, height: 180 };
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

    if (element.type === "image") {
      const box = element.box ?? defaultBox(element, nextY);
      nextY = box.y + box.height + 24;
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
      const box = element.box ?? defaultBox(element, nextY);
      nextY = box.y + box.height + 24;
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
        ...(element.text !== undefined ? {
          text: {
            content: element.text.content,
            ...(element.text.textStylePreset !== undefined ? { textStylePreset: element.text.textStylePreset } : {}),
            ...(element.text.frame !== undefined ? { frame: element.text.frame } : {}),
          },
        } : {}),
        ...(element.style !== undefined ? {
          style: {
            ...(element.style.fill !== undefined ? { fill: { type: "solid" as const, color: element.style.fill } } : {}),
            ...(stroke !== undefined ? { stroke } : {}),
          },
        } : {}),
      };
    }

    if (element.type === "group") {
      const box = element.box ?? defaultBox(element, nextY);
      nextY = box.y + box.height + 24;
      return {
        type: "group",
        box,
        coordinateSize: element.coordinateSize,
        children: createSlideElements(presentation, element.children),
      };
    }

    if (element.type === "table") {
      const box = element.box ?? defaultBox(element, nextY);
      nextY = box.y + box.height + 24;
      return {
        type: "table",
        box,
        columns: element.columns,
        rows: element.rows.map((row) => ({
          ...(row.height !== undefined ? { height: row.height } : {}),
          cells: row.cells.map((cell) => ({
            content: cell.content,
            ...(cell.rowSpan !== undefined ? { rowSpan: cell.rowSpan } : {}),
            ...(cell.colSpan !== undefined ? { colSpan: cell.colSpan } : {}),
          })),
        })),
      };
    }

    const box = element.box ?? defaultBox(element, nextY);
    nextY = box.y + (box.height ?? 24) + 24;
    const style = element.style;
    const content = style === undefined
      ? element.text
      : element.text.split(/\r\n|\r|\n/).map((text) => ({
          style: {
            ...(style.align !== undefined ? { align: style.align } : {}),
            ...(style.indent !== undefined ? { indent: style.indent } : {}),
            ...(style.hanging !== undefined ? { hanging: style.hanging } : {}),
            ...(style.lineSpacing !== undefined ? { lineSpacing: style.lineSpacing } : {}),
            ...(style.spaceBefore !== undefined ? { spaceBefore: style.spaceBefore } : {}),
            ...(style.spaceAfter !== undefined ? { spaceAfter: style.spaceAfter } : {}),
            ...(style.bullet !== undefined ? { bullet: style.bullet } : {}),
          },
          runs: [{
            text,
            style: {
              ...(style.fontSize !== undefined ? { fontSize: style.fontSize } : {}),
              ...(style.fontFamily !== undefined ? { fontFamily: style.fontFamily } : {}),
              ...(style.fontWeight !== undefined ? { bold: style.fontWeight === "bold" } : {}),
              ...(style.italic !== undefined ? { italic: style.italic } : {}),
              ...(style.underline !== undefined ? { underline: style.underline } : {}),
              ...(style.strike !== undefined ? { strike: style.strike } : {}),
              ...(style.color !== undefined ? { color: style.color } : {}),
              ...(style.language !== undefined ? { language: style.language } : {}),
            },
          }],
        }));

    return {
      type: "text",
      content,
      box,
      ...(element.textStylePreset !== undefined ? { textStylePreset: element.textStylePreset } : {}),
      ...(style?.autoFit !== undefined ? { frame: { autoFit: style.autoFit } } : {}),
    };
  });
}

export function createExamplePresentation(input: ExampleInputData): PresentationDocument {
  const presentation = createPresentation({
    metadata: { title: input.title },
    ...(input.size !== undefined ? { size: input.size } : {}),
    ...(input.textStylePresets !== undefined ? { textStylePresets: input.textStylePresets } : {}),
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
