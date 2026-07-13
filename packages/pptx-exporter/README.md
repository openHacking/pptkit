# `@pptkit/pptx-exporter`

PPTX export boundary for PPTKit.

## Current Status

This package writes a minimal editable `.pptx` package from the normalized PPTKit model.

The public entry point only orchestrates normalization, layout, package assembly, ZIP encoding, and output. Asset loading, OOXML generation, package parts, ZIP encoding, and filesystem output live in separate internal modules so they can evolve and be tested independently.

Supported elements are text, images, rectangles, ellipses, and lines. Image assets may use local paths or HTTP(S) URLs. Recoverable asset failures are returned as structured warnings while the remaining document is exported.

```ts
const result = await exportPptx(presentation, { output: "./deck.pptx" });
// result.status: "written" | "written-with-warnings"
```

The exporter intentionally does not yet cover themes, rich text runs, groups, tables, animations, speaker notes, or PPTX parsing.
