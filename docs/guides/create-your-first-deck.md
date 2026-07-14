# Create Your First Deck

This walkthrough builds a two-slide presentation with a theme, reusable layout, placeholder-bound rich text, a local image, a slide action, validation, and Node.js file output.

## Before you begin

Complete [Install PPTKit](../getting-started/installation.md) and place an image at `./assets/overview.png` relative to the process that runs the example. Supplying its pixel dimensions allows Layout to resolve `cover` cropping deterministically.

## Complete example

<!-- doc-test: docs/examples/first-deck.ts -->
```ts
import {
  createPresentation,
  PresentationValidationError,
  validatePresentation,
} from "@pptkit/core";
import { writePptx } from "@pptkit/pptx-exporter/node";

const presentation = createPresentation({
  metadata: {
    title: "Product Review",
    author: "Example Team",
    subject: "Quarterly product review",
  },
  theme: {
    name: "Product",
    colors: { accent1: "2457D6", accent2: "F97316" },
    fonts: { heading: "Aptos Display", body: "Aptos" },
  },
});

presentation.defineSlideLayout({
  id: "content",
  name: "Content",
  elements: [{
    type: "shape",
    shape: "rect",
    box: { x: 0, y: 0, width: 12, height: 540 },
    style: { fill: { type: "solid", color: { theme: "accent1" } } },
  }],
  placeholders: [
    {
      key: "title",
      kind: "title",
      box: { x: 48, y: 36, width: 820, height: 64 },
      textStyle: {
        run: { fontFamily: { theme: "heading" }, fontSize: 32, bold: true },
      },
    },
    {
      key: "body",
      kind: "body",
      box: { x: 48, y: 124, width: 390, height: 300 },
    },
  ],
});

const overviewImage = presentation.registerAsset({
  kind: "image",
  source: { type: "path", value: "./assets/overview.png" },
  mimeType: "image/png",
  width: 1600,
  height: 900,
  accessibility: { description: "Product overview dashboard" },
  dedupeKey: "overview-dashboard",
});

const summary = presentation.addSlide({ id: "summary", layoutId: "content" });
summary.addElement({
  type: "text",
  content: "Product Review",
  placeholderKey: "title",
});
summary.addElement({
  type: "text",
  placeholderKey: "body",
  content: [
    {
      style: { bullet: { type: "bullet" }, spaceAfter: 8 },
      runs: [{ text: "Activation increased 18%", style: { fontSize: 22 } }],
    },
    {
      style: { bullet: { type: "bullet" }, spaceAfter: 8 },
      runs: [{ text: "Retention remains the next focus", style: { fontSize: 22 } }],
    },
  ],
});
summary.addElement({
  type: "image",
  assetId: overviewImage.id,
  box: { x: 480, y: 124, width: 400, height: 250 },
  fit: "cover",
});

const detail = presentation.addSlide({ id: "detail", layoutId: "content" });
detail.addElement({ type: "text", content: "Next Actions", placeholderKey: "title" });
detail.addElement({
  type: "table",
  box: { x: 48, y: 130, width: 760, height: 190 },
  columns: [300, 230, 230],
  rows: [
    {
      cells: [{
        content: "Execution plan",
        colSpan: 3,
        style: { fill: { type: "solid", color: { theme: "accent1" } } },
      }],
    },
    { cells: [{ content: "Workstream" }, { content: "Owner" }, { content: "Status" }] },
    { cells: [{ content: "Onboarding" }, { content: "Growth" }, { content: "In progress" }] },
  ],
  action: { type: "slide", slideId: summary.id, tooltip: "Back to summary" },
});

const diagnostics = validatePresentation(presentation);
for (const diagnostic of diagnostics) {
  console.error(diagnostic.severity, diagnostic.code, diagnostic.path, diagnostic.message);
}

if (diagnostics.some((diagnostic) => diagnostic.severity === "error")) {
  throw new Error("The presentation contains validation errors.");
}

try {
  const result = await writePptx(presentation, {
    output: "./output/product-review.pptx",
  });
  console.log(result.status, result.output, result.warnings);
} catch (error) {
  if (error instanceof PresentationValidationError) {
    console.error(error.diagnostics);
  } else {
    throw error;
  }
}
```

## What the example demonstrates

- The theme is partial; Core materializes every omitted theme role.
- The layout owns reusable decoration and placeholder geometry.
- Slides bind content by stable `layoutId` and `placeholderKey` values.
- The image source is registered once and referenced by asset ID.
- Source dimensions allow Layout to resolve centered `cover` cropping.
- Navigation uses a stable slide ID rather than a page number.
- Validation reports all document problems before asset loading and packaging.
- An unreadable image becomes an exporter warning; invalid presentation structure remains a hard error.

## Drawing order and editing

Elements draw in their array order. Use `moveElement()`, `insertElement()`, or `duplicateElement()` to change authoring state; do not mutate `slide.elements`.

```ts
const copy = presentation.duplicateSlide("detail");
presentation.moveSlide(copy.id, 1);
presentation.removeSlide(copy.id);
```

## Continue reading

- [Core presentation operations](../api/core/presentation.md)
- [Element reference](../api/core/elements.md)
- [Themes and layouts](../api/core/themes-and-layouts.md)
- [Validation and IR](../api/core/validation-and-ir.md)
- [PPTX exporter](../api/pptx-exporter.md)
