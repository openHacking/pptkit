# Quick Start

This guide describes the intended first-run experience for PPTKit.

The APIs shown below are provisional. They exist to document the target developer experience before the first implementation is published.

## Goal

Generate an editable PPTX from structured slide content with a small, explicit package surface.

## Planned Example

```ts
import { createPresentation } from "@pptkit/core";
import { exportPptx } from "@pptkit/pptx-exporter";

const presentation = createPresentation({
  title: "Hello PPTKit",
});

presentation.addSlide({
  elements: [
    {
      type: "text",
      text: "Hello, PPTKit",
      x: 48,
      y: 48,
      width: 400,
      height: 40,
    },
  ],
});

await exportPptx(presentation, {
  output: "./hello-pptkit.pptx",
});
```

## Mental Model

The initial developer flow is expected to be:

1. Create a presentation model with `@pptkit/core`.
2. Arrange content using shared layout primitives from `@pptkit/layout`.
3. Export the result through `@pptkit/pptx-exporter`.

## Why This Matters

The quick-start experience should make three things obvious:

- The project has clear package responsibilities.
- Editable presentation output is a first-class goal.
- The user does not need to wire together low-level PPT XML details by hand.

## Next Steps

After the first public release, this guide should link to:

- Installation
- Hello World
- Layout guides
- SVG usage guides
- Export options
