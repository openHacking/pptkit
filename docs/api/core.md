# `@pptkit/core`

`@pptkit/core` is the format-independent authoring and document-contract package for PPTKit. It owns presentation construction, stable identities, validation, themes, layouts, assets, and Canonical Presentation IR v1. It does not read files, fetch URLs, calculate final layout, or write PPTX packages.

## Installation status

PPTKit preview packages are available on npm. Inside this repository, workspace packages import Core directly:

```ts
import {
  createPresentation,
  normalizePresentation,
  validatePresentation,
} from "@pptkit/core";
```

The package installation command is:

```bash
pnpm add @pptkit/core
```

## First presentation

```ts
import { createPresentation, normalizePresentation } from "@pptkit/core";

const presentation = createPresentation({
  metadata: {
    title: "Launch Plan",
    author: "Example Team",
    language: "en-US",
  },
  theme: {
    colors: { accent1: "2457D6" },
    fonts: { heading: "Aptos Display", body: "Aptos" },
  },
});

const slide = presentation.addSlide();
slide.addElement({
  type: "text",
  content: "Launch Plan",
  box: { x: 48, y: 48, width: 500, height: 60 },
});

const normalized = normalizePresentation(presentation);
console.log(normalized.irVersion, normalized.slides.length);
```

Authoring inputs may omit most defaults and IDs. Normalization validates the complete document, generates a detached IR, and materializes the values required by layout and exporters.

## Chart elements

Chart elements express categorical data as bar, line, or pie visualizations. A
chart carries its categories, one or more data series, an optional title, and
axis/legend visibility flags.

```ts
slide.addElement({
  type: "chart",
  chartType: "bar",
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series: [
    { name: "Revenue", values: [120, 150, 180, 210] },
    { name: "Costs", values: [80, 90, 100, 120] },
  ],
  title: "Quarterly performance",
  showLegend: true,
  xAxis: { show: true, labels: true },
  yAxis: { show: true },
  box: { x: 48, y: 140, width: 624, height: 320 },
});
```

Normalization validates chart structure and materializes defaults:

- Pie charts require exactly one series; additional series are a validation error.
- Every series `values` array length must match the `categories` length.
- Series colors default from theme accents when omitted, so charts stay consistent
  with the presentation theme without explicit per-series colors.

The normalized contract is documented in [Canonical Presentation IR v1](../architecture/canonical-ir-v1.md).

## API reference

- [Presentations and slides](core/presentation.md) — initialization, metadata, method-managed collections, ordering, duplication, and slide semantics.
- [Elements](core/elements.md) — shared element fields plus text, image, shape, connector, group, and table inputs.
- [Text and styles](core/text-and-styles.md) — paragraphs, runs, paints, strokes, transforms, bullets, links, and defaults.
- [Themes and layouts](core/themes-and-layouts.md) — theme roles, reusable layouts, placeholders, and inheritance.
- [Assets](core/assets.md) — registration, lookup, deduplication, and runtime boundaries.
- [Validation and IR](core/validation-and-ir.md) — diagnostics, normalization failures, and IR v1 guarantees.

The exact normalized schema is documented separately in [Canonical Presentation IR v1](../architecture/canonical-ir-v1.md).

## Public exports

Runtime exports:

| Export | Purpose |
| --- | --- |
| `createPresentation(init?)` | Creates method-managed authoring state. |
| `validatePresentation(document)` | Returns every validation diagnostic found in a document. |
| `normalizePresentation(document)` | Produces detached Canonical IR v1 or throws one validation error containing all error diagnostics. |
| `PresentationValidationError` | Error class exposing a readonly `diagnostics` collection. |

Core also exports its public TypeScript types, including authoring inputs, normalized structures, geometry, themes, styles, diagnostics, assets, and element unions.

## Package boundary

Use Core when code needs to describe presentation intent. Use [`@pptkit/layout`](layout.md) to resolve connector anchors and image fitting, and use [`@pptkit/pptx-exporter`](pptx-exporter.md) to load assets and generate or write PPTX output.
