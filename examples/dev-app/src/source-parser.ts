import type {
  ExampleElementSpec,
  ExampleImageAssetSpec,
  ExampleImageElementSpec,
  ExampleInputData,
  ExampleGroupElementSpec,
  ExampleTableElementSpec,
  ExampleShapeElementSpec,
  ExampleTextElementSpec,
} from "./example-types.js";
import type { TextFrameStyleInput, TextStylePresetInput, TextStylePresetMap } from "@pptkit/core";

export class ExampleSourceError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ExampleSourceError";
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isFiniteNumber(value: unknown): value is number {
  return typeof value === "number" && Number.isFinite(value);
}

function isPositiveInteger(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value) && value > 0;
}

function parseBox(value: unknown, label: string) {
  if (!isRecord(value)) {
    throw new ExampleSourceError(`${label} must be an object.`);
  }

  const { x, y, width, height } = value;
  if (!isFiniteNumber(x) || !isFiniteNumber(y) || !isFiniteNumber(width) || !isFiniteNumber(height)) {
    throw new ExampleSourceError(`${label} requires finite numeric x, y, width, and height fields.`);
  }

  return { x, y, width, height };
}

function parseTextBox(value: unknown, label: string) {
  if (!isRecord(value)) {
    throw new ExampleSourceError(`${label} must be an object.`);
  }

  const { x, y, width, height } = value;
  if (!isFiniteNumber(x) || !isFiniteNumber(y) || !isFiniteNumber(width) || (height !== undefined && !isFiniteNumber(height))) {
    throw new ExampleSourceError(`${label} requires finite numeric x, y, and width fields; height is optional.`);
  }

  return { x, y, width, ...(height !== undefined ? { height } : {}) };
}

function parseRecord(value: unknown, label: string) {
  if (!isRecord(value)) {
    throw new ExampleSourceError(`${label} must be an object.`);
  }

  return value;
}

function parseTextStylePresets(value: unknown): TextStylePresetMap {
  const source = parseRecord(value, "Source field textStylePresets");
  const presets: Record<string, TextStylePresetInput> = {};
  for (const [name, presetValue] of Object.entries(source)) {
    if (name.trim() === "") {
      throw new ExampleSourceError("Source field textStylePresets cannot contain an empty preset name.");
    }
    const preset = parseRecord(presetValue, `Source field textStylePresets.${name}`);
    if (preset.frame !== undefined && !isRecord(preset.frame)) {
      throw new ExampleSourceError(`Source field textStylePresets.${name}.frame must be an object.`);
    }
    if (preset.paragraph !== undefined && !isRecord(preset.paragraph)) {
      throw new ExampleSourceError(`Source field textStylePresets.${name}.paragraph must be an object.`);
    }
    if (preset.run !== undefined && !isRecord(preset.run)) {
      throw new ExampleSourceError(`Source field textStylePresets.${name}.run must be an object.`);
    }
    presets[name] = {
      ...(preset.frame !== undefined ? { frame: preset.frame as TextFrameStyleInput } : {}),
      ...(preset.paragraph !== undefined ? { paragraph: preset.paragraph } : {}),
      ...(preset.run !== undefined ? { run: preset.run } : {}),
    };
  }
  return presets;
}

function parseSize(value: unknown) {
  const size = parseRecord(value, "Source field size");
  const result: { width?: number; height?: number; unit?: "pt" } = {};

  for (const field of ["width", "height"] as const) {
    const fieldValue = size[field];
    if (fieldValue !== undefined && (!isFiniteNumber(fieldValue) || fieldValue <= 0)) {
      throw new ExampleSourceError(`Source field size.${field} must be a positive finite number.`);
    }
    if (typeof fieldValue === "number") result[field] = fieldValue;
  }

  if (size.unit !== undefined && size.unit !== "pt") {
    throw new ExampleSourceError('Source field size.unit must be "pt".');
  }
  if (size.unit === "pt") result.unit = "pt";
  return result;
}

function parseImageAsset(value: unknown, label: string): ExampleImageAssetSpec {
  if (!isRecord(value)) {
    throw new ExampleSourceError(`${label} must be an object.`);
  }

  if (value.id !== undefined && typeof value.id !== "string") {
    throw new ExampleSourceError(`${label}.id must be a string.`);
  }

  if (!isRecord(value.source)) {
    throw new ExampleSourceError(`${label}.source must be an object.`);
  }

  if (value.source.type !== "path" && value.source.type !== "url") {
    throw new ExampleSourceError(`${label}.source.type must be "path" or "url".`);
  }

  if (typeof value.source.value !== "string" || value.source.value.trim() === "") {
    throw new ExampleSourceError(`${label}.source.value must be a non-empty string.`);
  }

  for (const [field, fieldValue] of [
    ["mimeType", value.mimeType],
    ["altText", value.altText],
    ["dedupeKey", value.dedupeKey],
  ] as const) {
    if (fieldValue !== undefined && typeof fieldValue !== "string") {
      throw new ExampleSourceError(`${label}.${field} must be a string.`);
    }
  }

  for (const [field, fieldValue] of [
    ["width", value.width],
    ["height", value.height],
  ] as const) {
    if (fieldValue !== undefined && !isFiniteNumber(fieldValue)) {
      throw new ExampleSourceError(`${label}.${field} must be a finite number.`);
    }
  }

  const mimeType = typeof value.mimeType === "string" ? value.mimeType : undefined;
  const altText = typeof value.altText === "string" ? value.altText : undefined;
  const dedupeKey = typeof value.dedupeKey === "string" ? value.dedupeKey : undefined;
  const width = typeof value.width === "number" ? value.width : undefined;
  const height = typeof value.height === "number" ? value.height : undefined;

  return {
    ...(value.id !== undefined ? { id: value.id } : {}),
    source: {
      type: value.source.type,
      value: value.source.value,
    },
    ...(mimeType !== undefined ? { mimeType } : {}),
    ...(width !== undefined ? { width } : {}),
    ...(height !== undefined ? { height } : {}),
    ...(altText !== undefined ? { altText } : {}),
    ...(dedupeKey !== undefined ? { dedupeKey } : {}),
  };
}

function parseElement(value: unknown, slideIndex: number, elementIndex: number): ExampleElementSpec {
  const label = `Slide ${slideIndex + 1} element ${elementIndex + 1}`;

  if (typeof value === "string") {
    return value;
  }

  if (!isRecord(value) || typeof value.type !== "string") {
    throw new ExampleSourceError(`${label} must be a string or typed element object.`);
  }

  if (value.type === "text") {
    if (typeof value.text !== "string") {
      throw new ExampleSourceError(`${label}.text must be a string.`);
    }
    if (value.textStylePreset !== undefined && typeof value.textStylePreset !== "string") {
      throw new ExampleSourceError(`${label}.textStylePreset must be a string.`);
    }

    const element: ExampleTextElementSpec = {
      type: "text",
      text: value.text,
      ...(typeof value.textStylePreset === "string" ? { textStylePreset: value.textStylePreset } : {}),
      ...(value.box !== undefined ? { box: parseTextBox(value.box, `${label}.box`) } : {}),
      ...(value.style !== undefined ? { style: parseRecord(value.style, `${label}.style`) } : {}),
    };
    return element;
  }

  if (value.type === "shape") {
    if (value.shape !== "rect" && value.shape !== "roundRect" && value.shape !== "ellipse" && value.shape !== "line") {
      throw new ExampleSourceError(`${label}.shape must be "rect", "roundRect", "ellipse", or "line".`);
    }

    if (value.text !== undefined && !isRecord(value.text)) {
      throw new ExampleSourceError(`${label}.text must be an object.`);
    }
    if (isRecord(value.text) && typeof value.text.content !== "string") {
      throw new ExampleSourceError(`${label}.text.content must be a string.`);
    }
    if (isRecord(value.text) && value.text.textStylePreset !== undefined && typeof value.text.textStylePreset !== "string") {
      throw new ExampleSourceError(`${label}.text.textStylePreset must be a string.`);
    }
    if (isRecord(value.text) && value.text.frame !== undefined && !isRecord(value.text.frame)) {
      throw new ExampleSourceError(`${label}.text.frame must be an object.`);
    }

    const element: ExampleShapeElementSpec = {
      type: "shape",
      shape: value.shape,
      ...(isRecord(value.text) ? {
        text: {
          content: value.text.content as string,
          ...(typeof value.text.textStylePreset === "string" ? { textStylePreset: value.text.textStylePreset } : {}),
          ...(isRecord(value.text.frame) ? { frame: value.text.frame as TextFrameStyleInput } : {}),
        },
      } : {}),
      ...(value.box !== undefined ? { box: parseBox(value.box, `${label}.box`) } : {}),
      ...(value.style !== undefined ? { style: parseRecord(value.style, `${label}.style`) } : {}),
    };
    return element;
  }

  if (value.type === "image") {
    const asset = parseImageAsset(value.asset, `${label}.asset`);

    if (value.altText !== undefined && typeof value.altText !== "string") {
      throw new ExampleSourceError(`${label}.altText must be a string.`);
    }

    const element: ExampleImageElementSpec = {
      type: "image",
      asset,
      ...(value.box !== undefined ? { box: parseBox(value.box, `${label}.box`) } : {}),
      ...(value.altText !== undefined ? { altText: value.altText } : {}),
    };
    return element;
  }

  if (value.type === "group") {
    if (!isRecord(value.coordinateSize) || !isPositiveInteger(value.coordinateSize.width) || !isPositiveInteger(value.coordinateSize.height)) {
      throw new ExampleSourceError(`${label}.coordinateSize requires finite numeric width and height fields.`);
    }
    if (!Array.isArray(value.children)) {
      throw new ExampleSourceError(`${label}.children must be an array.`);
    }
    const element: ExampleGroupElementSpec = {
      type: "group",
      coordinateSize: { width: value.coordinateSize.width, height: value.coordinateSize.height },
      children: value.children.map((child, childIndex) => parseElement(child, slideIndex, elementIndex + childIndex + 1)),
      ...(value.box !== undefined ? { box: parseBox(value.box, `${label}.box`) } : {}),
    };
    return element;
  }

  if (value.type === "table") {
    if (!Array.isArray(value.columns) || value.columns.length === 0 || value.columns.some((column) => !isFiniteNumber(column) || column <= 0)) {
      throw new ExampleSourceError(`${label}.columns must contain positive finite numbers.`);
    }
    if (!Array.isArray(value.rows) || value.rows.length === 0) {
      throw new ExampleSourceError(`${label}.rows must be a non-empty array.`);
    }
    const rows = value.rows.map((row, rowIndex) => {
      if (!isRecord(row) || !Array.isArray(row.cells)) {
        throw new ExampleSourceError(`${label}.rows.${rowIndex}.cells must be an array.`);
      }
      if (row.height !== undefined && (!isFiniteNumber(row.height) || row.height <= 0)) {
        throw new ExampleSourceError(`${label}.rows.${rowIndex}.height must be a positive finite number.`);
      }
      return {
        ...(row.height !== undefined ? { height: row.height as number } : {}),
        cells: row.cells.map((cell, cellIndex) => {
          if (!isRecord(cell) || typeof cell.content !== "string") {
            throw new ExampleSourceError(`${label}.rows.${rowIndex}.cells.${cellIndex}.content must be a string.`);
          }
          if (cell.rowSpan !== undefined && !isPositiveInteger(cell.rowSpan)) {
            throw new ExampleSourceError(`${label}.rows.${rowIndex}.cells.${cellIndex}.rowSpan must be a positive integer.`);
          }
          if (cell.colSpan !== undefined && !isPositiveInteger(cell.colSpan)) {
            throw new ExampleSourceError(`${label}.rows.${rowIndex}.cells.${cellIndex}.colSpan must be a positive integer.`);
          }
          return {
            content: cell.content,
            ...(cell.rowSpan !== undefined ? { rowSpan: cell.rowSpan as number } : {}),
            ...(cell.colSpan !== undefined ? { colSpan: cell.colSpan as number } : {}),
          };
        }),
      };
    });
    const element: ExampleTableElementSpec = {
      type: "table",
      columns: value.columns,
      rows,
      ...(value.box !== undefined ? { box: parseBox(value.box, `${label}.box`) } : {}),
    };
    return element;
  }

  throw new ExampleSourceError(`${label}.type must be "text", "shape", "image", "group", or "table".`);
}

export function parseExampleSource(source: string): ExampleInputData {
  let value: unknown;

  try {
    value = JSON.parse(source);
  } catch {
    throw new ExampleSourceError("Source must be valid JSON.");
  }

  if (!isRecord(value)) {
    throw new ExampleSourceError("Source must be a JSON object.");
  }

  if (typeof value.title !== "string" || value.title.trim() === "") {
    throw new ExampleSourceError('Source requires a non-empty string "title".');
  }

  if (value.summary !== undefined && typeof value.summary !== "string") {
    throw new ExampleSourceError('Source field "summary" must be a string.');
  }

  if (!Array.isArray(value.slides) || value.slides.length === 0) {
    throw new ExampleSourceError('Source requires a non-empty "slides" array.');
  }

  const slides = value.slides.map((slide, index) => {
    if (!isRecord(slide)) {
      throw new ExampleSourceError(`Slide ${index + 1} must be an object.`);
    }

    if (slide.id !== undefined && typeof slide.id !== "string") {
      throw new ExampleSourceError(`Slide ${index + 1} field "id" must be a string.`);
    }

    if (typeof slide.title !== "string" || slide.title.trim() === "") {
      throw new ExampleSourceError(`Slide ${index + 1} requires a non-empty string "title".`);
    }

    if (slide.background !== undefined && (typeof slide.background !== "string" || slide.background.trim() === "")) {
      throw new ExampleSourceError(`Slide ${index + 1} field "background" must be a non-empty string.`);
    }

    if (!Array.isArray(slide.elements)) {
      throw new ExampleSourceError(`Slide ${index + 1} field "elements" must be an array.`);
    }

    return {
      ...(slide.id !== undefined ? { id: slide.id } : {}),
      title: slide.title,
      ...(typeof slide.background === "string" ? { background: slide.background } : {}),
      elements: slide.elements.map((element, elementIndex) => parseElement(element, index, elementIndex)),
    };
  });

  return {
    title: value.title,
    summary: value.summary ?? "",
    ...(value.size !== undefined ? { size: parseSize(value.size) } : {}),
    ...(value.textStylePresets !== undefined ? { textStylePresets: parseTextStylePresets(value.textStylePresets) } : {}),
    slides,
  };
}
