import { resolveLayout } from "@pptkit/layout";
import type {
  ExampleElementSpec,
  ExampleDefinition,
  ExampleReport,
  NormalizedDocumentReport,
} from "./example-types.js";
import { parseExampleSource } from "./source-parser.js";
import { createExamplePresentation } from "./presentation-builder.js";

function describeElement(element: ExampleElementSpec): string {
  if (typeof element === "string") {
    return element;
  }

  if (element.type === "text") {
    return element.text;
  }

  if (element.type === "shape") {
    return element.text === undefined ? `Shape: ${element.shape}` : `Shape: ${element.shape} — ${element.text.content}`;
  }

  if (element.type === "group") {
    return `Group: ${element.children.length} children`;
  }

  if (element.type === "table") {
    return `Table: ${element.columns.length} columns × ${element.rows.length} rows`;
  }

  const assetLabel =
    element.altText ??
    element.asset.altText ??
    element.asset.id ??
    element.asset.source.value;
  return `Image: ${assetLabel}`;
}

function buildNormalizedDocument(input: ReturnType<ExampleDefinition["createInput"]>): NormalizedDocumentReport {
  const document = createExamplePresentation(input);

  const slides = input.slides.map((slide, index) => {
    const created = document.slides[index]!;
    return {
      id: created.id,
      title: slide.title,
      elements: slide.elements.map((element) => describeElement(element)),
    };
  });

  return {
    title: input.title,
    summary: input.summary,
    slideCount: slides.length,
    slides,
  };
}

export async function buildExampleReport(example: ExampleDefinition, sourceOverride?: string): Promise<ExampleReport> {
  const input = sourceOverride === undefined ? example.createInput() : parseExampleSource(sourceOverride);
  const normalizedDocument = buildNormalizedDocument(input);
  const presentation = createExamplePresentation(input);

  const renderResult = resolveLayout(presentation);
  const exportResult = {
    output: "",
    slideCount: normalizedDocument.slideCount,
    status: "not-exported",
  } as const;

  const diagnostics = [
    `Normalize: ${example.expectedCapabilities.normalize}`,
    `Render: ${example.expectedCapabilities.render}`,
    `Export: ${example.expectedCapabilities.exportPptx}`,
  ];

  diagnostics.push("PPTX export is available from the Export PPTX action.");

  return {
    example: sourceOverride === undefined ? example : { ...example, source: { ...example.source, content: sourceOverride } },
    normalizedDocument,
    renderResult,
    exportResult,
    diagnostics,
  };
}
