# Hello World

This page shows the smallest complete `@pptkit/core` example currently available in the repository.

## Example

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";

const presentation = createPresentation({
  title: "Hello World",
});

presentation.addSlide({
  elements: [
    {
      type: "text",
      text: "Hello World",
      box: {
        x: 64,
        y: 64,
        width: 320,
        height: 36,
      },
      style: {
        fontSize: 28,
        fontWeight: "bold",
      },
    },
  ],
});

const normalized = normalizePresentation(presentation);

console.log(normalized);
```

## What You Get

- one presentation document with default size metadata
- one slide with a typed text element
- one normalized output object ready for layout or export

## Current Limitation

If you need a `.pptx` file, pass the presentation into `@pptkit/pptx-exporter`.

That package currently returns a structured placeholder result and does not write a real PowerPoint file yet.
