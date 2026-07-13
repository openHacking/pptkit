# Quick Start

This guide takes you from a local checkout to an editable `.pptx` file using the current pre-release API.

## Prerequisites

- Node.js 20 or newer
- pnpm 10
- a local checkout of this repository

PPTKit packages are not published yet. Install and build the workspace before running an example:

```bash
pnpm install
pnpm build
```

## Create and export a presentation

Save the following as `quick-start.ts` in a workspace package or another project configured to resolve the workspace packages.

<!-- doc-test: docs/examples/quick-start.ts -->
```ts
import { createPresentation, validatePresentation } from "@pptkit/core";
import { writePptx } from "@pptkit/pptx-exporter/node";

const presentation = createPresentation({
  metadata: { title: "Hello PPTKit", author: "Example Team" },
  theme: { colors: { accent1: "2457D6" } },
});

const slide = presentation.addSlide();
slide.addElement({
  type: "text",
  content: [{
    runs: [
      { text: "Hello ", style: { fontSize: 36 } },
      {
        text: "PPTKit",
        style: { fontSize: 36, bold: true, color: { theme: "accent1" } },
      },
    ],
  }],
  box: { x: 64, y: 64, width: 520, height: 72 },
});

const diagnostics = validatePresentation(presentation);
if (diagnostics.some((diagnostic) => diagnostic.severity === "error")) {
  throw new Error(JSON.stringify(diagnostics, null, 2));
}

const result = await writePptx(presentation, {
  output: "./hello-pptkit.pptx",
});

console.log(result.status, result.output, result.warnings);
```

Run it with your TypeScript runner or compile it with your project.

The repository keeps the checked source at `docs/examples/quick-start.ts`. To execute that exact file with the workbench's TypeScript runner:

```bash
pnpm --filter @pptkit/dev-app exec tsx \
  --tsconfig ../../docs/examples/tsconfig.json \
  ../../docs/examples/quick-start.ts
```

The command writes `hello-pptkit.pptx` in `examples/dev-app/`.

The Node exporter creates parent directories when necessary and returns:

- `status`: `written` or `written-with-warnings`
- `output`: the written path
- `slideCount` and `byteLength`
- `warnings`: recoverable asset-loading problems

Core validation errors are not warnings. `writePptx()` rejects with `PresentationValidationError` before packaging when the document is invalid.

## Why the code uses methods

Slides and elements are inserted through methods so Core can own stable IDs and validate ordering. The exposed collections are readonly snapshots:

```ts
const second = presentation.addSlide();
second.addElement({
  type: "shape",
  shape: "roundRect",
  box: { x: 80, y: 120, width: 300, height: 160 },
  style: {
    fill: { type: "solid", color: { theme: "accent1" }, opacity: 0.15 },
    stroke: { paint: { type: "solid", color: { theme: "accent1" } }, width: 2 },
  },
});
```

Element order is drawing order. Use `moveElement()` instead of a `zIndex` field.

## Browser generation

Browser-capable runtimes use the default exporter entry and receive bytes instead of writing a file:

```ts
import { generatePptx } from "@pptkit/pptx-exporter";

const result = await generatePptx(presentation);
const blob = new Blob([result.bytes], {
  type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
});
```

The default entry loads URL assets. The Node entry additionally loads local path assets.

## Next steps

- [Create Your First Deck](../guides/create-your-first-deck.md) adds layouts, placeholders, images, multiple slides, and error handling.
- [Core API](../api/core.md) documents every authoring method and input family.
- [Elements](../api/core/elements.md) covers images, shapes, connectors, groups, and tables.
- [Themes and layouts](../api/core/themes-and-layouts.md) explains inheritance and placeholders.
- [PPTX exporter](../api/pptx-exporter.md) documents runtime loading, results, warnings, and failures.
