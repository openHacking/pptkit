import { createPresentation } from "../../../packages/core/dist/index.js";
import { resolveLayout } from "../../../packages/layout/dist/index.js";
import { exportPptx } from "../../../packages/pptx-exporter/dist/index.js";
import type {
  ExampleDefinition,
  ExampleReport,
  NormalizedDocumentReport,
  VisualPreviewSlide,
} from "./example-types.js";

function buildNormalizedDocument(example: ExampleDefinition): NormalizedDocumentReport {
  const input = example.createInput();
  const document = createPresentation({ title: input.title });

  const slides = input.slides.map((slide) => {
    const created = document.addSlide({
      ...(slide.id !== undefined ? { id: slide.id } : {}),
      elements: slide.elements.map((element) => ({
        type: element,
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

export async function buildExampleReport(example: ExampleDefinition): Promise<ExampleReport> {
  const normalizedDocument = buildNormalizedDocument(example);
  const presentation = createPresentation({ title: normalizedDocument.title });

  for (const slide of normalizedDocument.slides) {
    presentation.addSlide({
      id: slide.id,
      elements: slide.elements.map((element) => ({
        type: element,
      })),
    });
  }

  const renderResult = resolveLayout(presentation);
  const exportResult = await exportPptx(presentation, {
    output: `${example.id}.pptx`,
  });

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

  if (exportResult.status === "not-implemented") {
    diagnostics.push("PPTX export remains a placeholder boundary and does not create a file yet.");
  }

  return {
    example,
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
