# `@pptkit/core`

`@pptkit/core` is the format-independent authoring and document-contract package for PPTKit. It owns presentation construction, stable identities, validation, themes, layouts, assets, and Canonical Presentation IR v2. It does not read files, fetch URLs, calculate final layout, or write PPTX packages.

[`resolvePresentationInheritance`](inheritance-origins.md) returns the ordinary resolved IR with a separate immutable theme/background origin snapshot.

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

Chart elements are a `chartType`-discriminated union. Core uses the general
`bar` family name and expresses direction separately through `orientation`.

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
  orientation: "vertical",
  grouping: "clustered",
  legend: { visible: true, position: "right" },
  axes: {
    category: { visible: true, labels: true },
    value: { scale: { min: 0, max: 250, majorUnit: 50 } },
  },
  box: { x: 48, y: 140, width: 624, height: 320 },
});
```

Normalization validates chart structure and materializes defaults:

- [Scatter charts](chart-scatter.md) use independent X/Y arrays and two numeric axes.
- Pie charts require exactly one series; doughnut charts support multiple concentric series. See [doughnut charts](chart-doughnut.md) for geometry and import limits.
- Every series `values` array length must match `categories`. Each position contains a finite number or explicit `null` for a blank point; sparse holes, `undefined`, NaN and infinities remain invalid. Zero and negative zero are values, not blanks. Sparse category, series and explicit pie-color arrays are invalid. Normalized consumers must now handle `null` in chart values.
- `blankDisplay` accepts `gap` or `zero` for bar, line, pie, doughnut, and standard area charts, and `span` for line and scatter charts. Stacked area charts accept only `zero` and default to it, including when values are null. See [area charts](chart-area.md). It controls display without changing `null` values. When omitted, charts containing blanks default to `gap`; dense charts retain the previous `zero` default. Legacy normalized charts without this property use the same defaults.
- Series colors default from theme accents when omitted, so charts stay consistent
  with the presentation theme without explicit per-series colors.
- Bar charts default to vertical, clustered layout. `categoryGap` and
  `seriesGap` are normalized explicitly. For non-percentage stacked bars, positive and negative cumulative sums must each remain finite; overflowing sums produce `bar-stack-overflow` rather than invalid preview coordinates.
- Cartesian charts receive explicit category/value axes, including line,
  labels, major ticks, gridlines, text styles, and fixed or automatic scale. Category
  major ticks default to `none` for bar charts (whose native axis boundaries fall between
  labels) and `outside` for line charts.
- Line markers accept `"auto"`, `false`, or an explicit shape/style. Automatic
  shapes cycle deterministically through diamond, square, circle, triangle,
  star, x, plus, and dash.
- Pie and doughnut `pointColors` resolve to one `#RRGGBB` value per category and pie charts
  accept exactly one series.

The normalized contract is documented in [Canonical Presentation IR v2](../architecture/canonical-ir-v2.md).

## API reference

- [Presentations and slides](core/presentation.md) — initialization, metadata, method-managed collections, ordering, duplication, and slide semantics.
- [Elements](core/elements.md) — shared element fields plus text, image, shape, connector, group, and table inputs.
- [Text and styles](core/text-and-styles.md) — paragraphs, runs, paints, strokes, transforms, bullets, links, and defaults.
- [Themes and layouts](core/themes-and-layouts.md) — theme roles, reusable layouts, placeholders, and inheritance.
- [Assets](core/assets.md) — registration, lookup, deduplication, and runtime boundaries.
- [Validation and IR](core/validation-and-ir.md) — diagnostics, normalization failures, and IR v2 guarantees.

The exact normalized schema is documented separately in [Canonical Presentation IR v2](../architecture/canonical-ir-v2.md).

## Public exports

Runtime exports:

| Export | Purpose |
| --- | --- |
| `createPresentation(init?)` | Creates method-managed authoring state. |
| `validatePresentation(document)` | Returns every validation diagnostic found in a document. |
| `normalizePresentation(document)` | Produces detached Canonical IR v2 or throws one validation error containing all error diagnostics. |
| `compileFiniteTimeline(root, limits?)` | Compiles abstract finite span/sequence/parallel intervals; see [finite time](finite-timeline.md). |
| `sampleFiniteTimeline(plan, timeMs, limits?)` | Validates a detached/transported schedule and returns immutable progress at absolute time. |
| `compileEventTimeline(input, limits?)` | Compiles detached foreground gate stages and whole-object inventories; see [event replay](event-timeline.md). |
| `replayEventTimeline(plan, request, limits?)` | Recomputes activations/observations from an explicit visit, current inventory and ordered log. |
| `EventTimelineError` | Located invalid/unsupported event data, stale targets, quotas, overflow and corrupt-plan errors. |
| `FiniteTimelineError` | Located errors for invalid/unsupported timing, quotas, overflow and corrupt schedules. |
| `PresentationValidationError` | Error class exposing a readonly `diagnostics` collection. |

Core also exports its public TypeScript types, including authoring inputs, normalized structures, geometry, themes, styles, diagnostics, assets, and element unions.

## Package boundary

Use Core when code needs to describe presentation intent. Use [`@pptkit/layout`](layout.md) to resolve connector anchors and image fitting, and use [`@pptkit/pptx-exporter`](pptx-exporter.md) to load assets and generate or write PPTX output.


Table structure changes use [editTable](table-editing.md), returning a frozen candidate and anchor mapping for explicit `replaceElement` commit.
