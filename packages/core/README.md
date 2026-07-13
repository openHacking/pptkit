# `@pptkit/core`

Authoring model and Canonical Presentation IR for PPTKit.

## Capabilities

- method-managed presentations, slides, elements, assets, and slide layouts
- rich text paragraphs and runs
- theme colors and fonts
- text, image, shape, connector, group, and table elements
- placeholders, speaker notes, hidden slides, sections, tags, and custom data
- structured validation and fully materialized IR v1 normalization

## Example

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";

const presentation = createPresentation({
  metadata: { title: "Quarterly Update", author: "Example Team" },
  theme: { colors: { accent1: "2457D6" } },
});

const slide = presentation.addSlide();
slide.addElement({
  type: "text",
  content: [{
    runs: [
      { text: "Quarterly ", style: { fontSize: 32 } },
      { text: "Update", style: { fontSize: 32, bold: true, color: { theme: "accent1" } } },
    ],
  }],
  box: { x: 48, y: 48, width: 500, height: 60 },
});

const normalized = normalizePresentation(presentation);
console.log(normalized.irVersion); // 1
```

Collections are exposed as readonly snapshots. Use document and slide methods to add, insert, move, remove, or duplicate content.

See [the Core API reference](../../docs/api/core.md) and [authoring architecture](../../docs/architecture/core-authoring-model.md).
