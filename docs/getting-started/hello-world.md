# Hello World: Core Only

This smallest example demonstrates Core authoring and normalization. It does not write a file; use the [Quick Start](quick-start.md) for an end-to-end PPTX example.

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";

const presentation = createPresentation({
  metadata: { title: "Hello World", author: "Example Team" },
});

const slide = presentation.addSlide();
slide.addElement({
  type: "text",
  content: "Hello World",
  box: { x: 64, y: 64, width: 320, height: 48 },
});

const normalized = normalizePresentation(presentation);
console.log(normalized.irVersion, normalized.slides.length);
```

`normalizePresentation()` validates the document and returns detached Canonical IR v1 with IDs and defaults materialized. It throws `PresentationValidationError` with all error diagnostics when the authoring document is invalid.

Read [Validation and IR](../api/core/validation-and-ir.md) for the complete contract.
