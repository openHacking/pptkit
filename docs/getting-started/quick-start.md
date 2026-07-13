# Quick Start

This guide describes the first-run experience for PPTKit package development.

## Goal

Build a presentation document with `@pptkit/core`, then pass it into downstream packages for layout and export work.

## Example

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { writePptx } from "@pptkit/pptx-exporter/node";

const presentation = createPresentation({
  title: "Hello PPTKit",
});

const chart = presentation.registerAsset({
  kind: "image",
  source: {
    type: "path",
    value: "./assets/chart.png",
  },
  mimeType: "image/png",
  dedupeKey: "quarterly-chart",
});

presentation.addSlide({
  elements: [
    {
      type: "text",
      text: "Hello, PPTKit",
      box: {
        x: 48,
        y: 48,
        width: 400,
        height: 40,
      },
    },
    {
      type: "image",
      assetId: chart.id,
      box: {
        x: 48,
        y: 120,
        width: 320,
        height: 180,
      },
    },
  ],
});

const normalized = normalizePresentation(presentation);
const result = await writePptx(presentation, {
  output: "./hello-pptkit.pptx",
});

console.log(normalized.slides.length, result.status);
```

## Mental Model

The current developer flow is:

1. Create a presentation model with `@pptkit/core`.
2. Register reusable assets and reference them from typed elements.
3. Normalize the document into a stable downstream structure.
4. Arrange and export through `@pptkit/layout` and `@pptkit/pptx-exporter`.

## Why This Matters

This quick start should make four things obvious:

- The project has clear package responsibilities.
- Assets are part of the core document model, not an exporter-only concern.
- The user does not need to wire together low-level PPT XML details by hand.
- Runtime-neutral generation returns editable PPTX bytes, while the Node.js adapter writes files and reports recoverable issues through `result.warnings`.

## Next Steps

Continue with:

- Installation
- Hello World
- `@pptkit/core` API reference
- Layout guides
- Export options
