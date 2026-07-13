# `@pptkit/pptx-exporter`

PPTX export boundary for PPTKit.

## Current Status

This package generates a minimal editable `.pptx` package from the normalized PPTKit model in browsers and Node.js.

The default entry point is runtime-neutral and returns PPTX bytes. Filesystem output is an explicit Node.js adapter exposed from `@pptkit/pptx-exporter/node`.

Supported elements are text, images, rectangles, ellipses, and lines. The default entry loads URL assets, including HTTP(S), data, and blob URLs. The Node.js entry additionally loads local paths. Recoverable asset failures are returned as structured warnings while the remaining document is exported.

```ts
import { generatePptx } from "@pptkit/pptx-exporter";

const result = await generatePptx(presentation);
// result.bytes: Uint8Array
```

```ts
import { writePptx } from "@pptkit/pptx-exporter/node";

const result = await writePptx(presentation, { output: "./deck.pptx" });
```

The exporter intentionally does not yet cover themes, rich text runs, groups, tables, animations, speaker notes, or PPTX parsing.
