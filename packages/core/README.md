# `@pptkit/core`

Authoring-facing document model for PPTKit.

## What It Provides

- `createPresentation()` for document construction
- typed `text`, `image`, and `shape` slide elements
- asset registration and lookup
- `normalizePresentation()` for downstream-safe document output

## Example

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";

const presentation = createPresentation({
  title: "Quarterly Update",
});

presentation.addSlide({
  elements: [
    {
      type: "text",
      text: "Quarterly Update",
      box: { x: 48, y: 48, width: 360, height: 32 },
    },
  ],
});

const normalized = normalizePresentation(presentation);

console.log(normalized.slides.length);
```

## Documentation

- API reference: [docs/api/core.md](../../docs/api/core.md)
- Architecture: [docs/architecture/core-authoring-model.md](../../docs/architecture/core-authoring-model.md)

## Internal Architecture

The package keeps its root API intentionally small. Internally, contracts live in `types`, mutable authoring state in `document`, and side-effect-free validation and normalization flows in their own modules. Internal classes are not part of the supported public API.
