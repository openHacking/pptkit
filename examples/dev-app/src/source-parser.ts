import type { ExampleInputData } from "./example-types.js";

export class ExampleSourceError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ExampleSourceError";
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
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

    if (!Array.isArray(slide.elements) || slide.elements.some((element) => typeof element !== "string")) {
      throw new ExampleSourceError(`Slide ${index + 1} field "elements" must be an array of strings.`);
    }

    return {
      ...(slide.id !== undefined ? { id: slide.id } : {}),
      title: slide.title,
      elements: slide.elements,
    };
  });

  return {
    title: value.title,
    summary: value.summary ?? "",
    slides,
  };
}
