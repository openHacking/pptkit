import { createPresentation } from "../../../packages/core/dist/index.js";
import { resolveLayout } from "../../../packages/layout/dist/index.js";
import type {
  ExampleDefinition,
  ExampleReport,
  NormalizedDocumentReport,
  VisualPreviewSlide,
} from "./example-types.js";
import { parseExampleSource } from "./source-parser.js";

function buildNormalizedDocument(input: ReturnType<ExampleDefinition["createInput"]>): NormalizedDocumentReport {
  const document = createPresentation({ title: input.title });

  const slides = input.slides.map((slide) => {
    const created = document.addSlide({
      ...(slide.id !== undefined ? { id: slide.id } : {}),
      elements: slide.elements.map((element, index) => ({
        type: "text",
        text: element,
        box: {
          x: 48,
          y: 48 + index * 32,
          width: 640,
          height: 24,
        },
      })),
    });

    return {
      id: created.id,
      title: slide.title,
      elements: [...slide.elements],
    };
  });

  return {
    title: input.title,
    summary: input.summary,
    slideCount: slides.length,
    slides,
  };
}

function buildVisualPreview(normalizedDocument: NormalizedDocumentReport): VisualPreviewSlide[] {
  return normalizedDocument.slides.map((slide) => ({
    id: slide.id,
    title: slide.title,
    body: slide.elements,
  }));
}

export async function buildExampleReport(example: ExampleDefinition, sourceOverride?: string): Promise<ExampleReport> {
  const input = sourceOverride === undefined ? example.createInput() : parseExampleSource(sourceOverride);
  const normalizedDocument = buildNormalizedDocument(input);
  const presentation = createPresentation({ title: normalizedDocument.title });

  for (const slide of normalizedDocument.slides) {
    presentation.addSlide({
      id: slide.id,
      elements: slide.elements.map((element, index) => ({
        type: "text",
        text: element,
        box: {
          x: 48,
          y: 48 + index * 32,
          width: 640,
          height: 24,
        },
      })),
    });
  }

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

  if (example.expectedCapabilities.render !== "implemented") {
    diagnostics.push(
      "Visual preview uses a structural slide mock instead of a production renderer.",
    );
  }

  diagnostics.push("PPTX export is available from the Export PPTX action.");

  return {
    example: sourceOverride === undefined ? example : { ...example, source: { ...example.source, content: sourceOverride } },
    normalizedDocument,
    visualPreview: {
      status: "structural-preview",
      slides: buildVisualPreview(normalizedDocument),
    },
    renderResult,
    exportResult,
    diagnostics,
  };
}
