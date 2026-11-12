# `@pptkit/layout`

`@pptkit/layout` converts validated Canonical Presentation IR into a detached, export-ready layout result. It resolves connector anchors, image `contain`/`cover` geometry, charts, and opt-in composition/table/intrinsic sizing and pagination. Fixed-width text auto-height is resolved earlier by Core normalization; Layout uses injected providers for environment-dependent text measurement and does not write files or serialize OOXML.

## `resolveLayout(document)`

```ts
declare function resolveLayout(
  document: PresentationDocument,
  options?: LayoutOptions,
): LayoutResult;
```

This is the authoring-facing convenience entry point. It calls `normalizePresentation(document)` and therefore throws `PresentationValidationError` when the document contains error diagnostics.

```ts
import { createPresentation } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";

const presentation = createPresentation();
const slide = presentation.addSlide();

slide.addElement({
  id: "source",
  type: "shape",
  shape: "rect",
  box: { x: 80, y: 120, width: 160, height: 80 },
});
slide.addElement({
  id: "target",
  type: "shape",
  shape: "ellipse",
  box: { x: 420, y: 120, width: 120, height: 80 },
});
slide.addElement({
  type: "connector",
  start: { elementId: "source", anchor: "right" },
  end: { elementId: "target", anchor: "left" },
});

const layout = resolveLayout(presentation);
console.log(layout.status, layout.slideCount);
```

## `resolveNormalizedLayout(normalized)`

```ts
declare function resolveNormalizedLayout(
  normalized: NormalizedPresentation,
  options?: LayoutOptions,
): LayoutResult;
```

Use this entry when a pipeline has already normalized the document. It avoids performing normalization twice. The function assumes the supplied IR satisfies the Core IR v2 contract.

## `LayoutResult`

```ts
interface LayoutResult {
  size: PresentationSize;
  metadata: NormalizedPresentationMetadata;
  theme: NormalizedPresentationTheme;
  assets: NormalizedAsset[];
  layouts: LayoutSlideLayout[];
  slides: LayoutSlide[];
  slideCount: number;
  status: "resolved";
  analysis?: LayoutAnalysis;
  pagination?: LayoutPagination;
}
```

The result is detached from both authoring state and the input normalized document.

## Resolution behavior

### Connectors

Element-reference endpoints become absolute points using `top`, `right`, `bottom`, `left`, or `center` anchors. The default anchor is `center`. The layout result also contains a box derived from the start point, route points, and end point.

Core validation normally prevents missing connector references. An unresolved endpoint passed directly to `resolveNormalizedLayout()` throws instead of inventing a coordinate. Connector bounds and ordinary chart data extents are accumulated iteratively, avoiding the runtime's function-argument limit for long arrays. Resolution still materializes document data and chart geometry; this is not a general memory or input-size guarantee.

### Images

- `stretch` keeps the requested box.
- `contain` uses registered source dimensions to create a centered contained box, then emits `fit: "stretch"` for export.
- `cover` derives centered normalized crop values, then emits `fit: "crop"`.
- `crop` preserves caller-provided normalized crop values.

Explicit fractional `crop` preview scales and clips the image in frame coordinates and does not require declared source dimensions. Very narrow crops use their actual fractions; nonfinite derived SVG coordinates raise `RangeError`. PPTX separately quantizes crop fractions to 1/100000, so extreme crops can differ; `image-crop-quantized-empty` warns when rounding removes an entire axis. If source dimensions are unavailable or their geometry cannot be represented with finite, nonzero coordinates and visible fractions, `contain` and `cover` remain unresolved rather than inventing an aspect ratio. SVG preview delegates those fits to the viewer using `meet` and `slice`, respectively. PPTX export instead stretches an unresolved image and emits `image-fit-unresolved` with its owner, element and asset identities. This also applies when a zero-sized frame prevents resolution; providing dimensions only from an export resolver cannot retroactively change layout.

### Groups

Group children remain in their local coordinate system. Layout recursively resolves child connectors and images; the exporter applies the group box and `coordinateSize` transform.

### Charts

Chart layout is resolved once for preview and export. The resolved contract contains the
plot box, value scale, category anchors, series geometry, and legend item boxes. A right
legend resolves as one column and a bottom legend resolves as one row. Line charts use
`categoryAxisPlacement: "onCategory"`, so category labels, points, and major tick marks
share anchors. Bar charts use `"betweenCategories"` to keep the first and last bars fully
inside the plot; their label anchors and boundary tick anchors are stored separately.
Scatter layouts expose `xScale`, the vertical `valueScale`, and per-series `scatter` points and segments with original `sourceIndex` and explicit zero-displayed blanks; see [scatter charts](chart-scatter.md). Doughnut layouts expose concentric rings and their annular slices; see [doughnut charts](chart-doughnut.md). Format-specific renderers map these semantics to native output without re-running chart
layout.

## Current limits

The package provides opt-in injected measurement and overflow diagnostics. A browser font-backed provider is available through the separate `@pptkit/layout/dom` entry. It provides explicit bounded grid and affine frame constraints, but does not provide automatic overflow correction, arbitrary simultaneous constraint solving or general text/story flow across slides. Explicit [table pagination](table-pagination.md) can split top-level tables into continuation slides. Core's deterministic fixed-width text height estimate is part of normalization, not a Layout measurement service. Native chart implementations may apply their own font metrics, but exporters must preserve the resolved legend flow and category-axis placement semantics.


## Opt-in measurement and diagnostics

Both resolution functions accept `LayoutOptions`. Calling without options (or with `{}`) retains the previous geometry and omits `analysis`. Measurement analysis never resizes elements or changes the renderer's line breaks. Explicit [composition rules](composition.md) can reposition fixed-size frames before analysis.

```ts
interface LayoutOptions {
  composition?: readonly LayoutComposition[];
  tableSizing?: readonly TableSizingRule[];
  intrinsicSizing?: readonly IntrinsicSizingRule[];
  tablePagination?: readonly TablePaginationRule[];
  textMeasurement?: {
    measure: TextMeasurer;
    resolveFont?: FontResolver;
  };
  diagnostics?: {
    textOverflow?: boolean;
    slideBounds?: boolean;
    tolerance?: number;
    allowedOffSlideElementIds?: readonly string[];
    includeHidden?: boolean;
  };
}
```

`measure` is a synchronous caller-supplied service. Load fonts before resolving Layout. Its frozen, detached request contains normalized paragraphs/runs, `availableWidth` excluding margins, and `wrap`. Theme fonts have been resolved to family names and shrink `fontScale` has been applied to run font sizes. The provider owns shaping, glyph coverage, wrapping, bullet markers, indentation, horizontal alignment, paragraph spacing, and line metrics. It returns `{ bounds: { x, y, width, height } }` in logical points relative to the content origin, before frame vertical alignment. Bounds must be finite with nonnegative dimensions; negative x/y are allowed for overhanging glyphs.

An optional synchronous `resolveFont` receives family, effective font size, bold/italic, language and text; it returns `{ family }` or `undefined` when unavailable. A changed family produces `font-fallback`. Failure or absence produces `font-resolution-failed` or `font-unavailable` and skips that text body. Without a resolver, the measurer is responsible for font availability. Provider exceptions and invalid results produce `text-measurement-failed` or `text-measurement-invalid`; asynchronous providers are rejected with a diagnostic, not awaited.

`analysis.measurements` contains `elementId`, source `path`, slide or layout identity, the element-local `contentBox`, and measured `bounds` after margins and vertical alignment. Table cells use the containing table's local coordinates. Templates are analyzed once with `layoutId`, rather than once per slide. Measurements and diagnostics are detached data.

`textOverflow` compares measured text, shape text and table cells with their content areas. For parallelograms, the content area uses the preset text rectangle before frame margins; text metrics and wrapping still depend on the supplied measurer. `resolveShapeTextBox` exposes the same untransformed rectangle for normalized shape geometry, returning a copy of the frame for other shapes. Zero-width parallelograms use the degenerate frame. If no measurer is supplied, `text-measurement-unavailable` explicitly states that text was not checked. Native chart label flow is not checked and produces `chart-text-overflow-unchecked`. Quarter-turn table-cell text uses the horizontal provider with the physical content height as available width, then rotates its bounds into cell coordinates. It reports `table-text-direction-measurement-approximate` because native margins and wrapping can differ. Automatic row sizing uses a conservative unwrapped extent and verifies the result with a wrapped measurement; it reports `table-text-direction-sizing-approximate` and, when the conservative extent exceeds a finite span maximum, tests wrapping at that maximum before deciding whether to retain authored geometry. Table measurement uses explicit row/column sizes and spans; mismatch with the frame produces `table-geometry-mismatch`, and impossible spans produce `table-span-invalid`. This diagnostic pass does not resize cells. Explicit [table sizing](table-sizing.md) resolves selected row heights before analysis.

`slideBounds` checks element frames and available measured text bounds in slide coordinates, including rotation, flips, and ancestor group scaling. It excludes stroke extents, shadows and other effects. Nonfinite transformed bounds produce `geometry-unmeasurable`. Hidden slides/elements/subtrees are skipped unless `includeHidden` is true. `allowedOffSlideElementIds` suppresses boundary warnings for an element and its descendants, but does not suppress text overflow. The default tolerance is 0.01 logical points; a negative/nonfinite tolerance or malformed configuration throws.

The separate opt-in `shadowBounds` estimates outer-shadow slide overflow using the SVG preview's three-sigma blur envelope, plus half the shape stroke width, transformed with the shape and ancestor groups. It emits `shadow-slide-overflow-approximate` with estimated bounds. It skips zero-opacity shadows and zero-opacity shape or ancestor groups, and respects hidden and allowed-off-slide settings. This conservative frame-based estimate does not inspect the painted silhouette or text and is not a native PowerPoint bounds guarantee; other effects remain unchecked.

Diagnostics include code, warning severity, message, object identity and source path. Overflow diagnostics also include `bounds` and `available` in the relevant coordinate system. `status` remains `resolved`; callers inspect `analysis.diagnostics` to decide their own acceptance policy. Exporter and renderer forward these diagnostics as warnings when analysis is requested.

Diagnostic batches are merged in their existing phase order without passing the full batch as function arguments. Field refresh likewise traverses wide groups without using function-argument expansion. Both operations still retain data in memory; neither establishes a general document-size or memory bound.

The opt-in [browser font-backed adapter](browser-measurement.md) implements this provider contract. Shared resolved line layout and native Office fidelity remain separate work. These contracts do not establish native PowerPoint font or wrapping parity.

## Checked integration example

<!-- doc-test: docs/examples/layout-analysis.ts -->
```ts
import type { PresentationDocument } from "@pptkit/core";
import { resolveLayout, type TextMeasurer } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";

// Supply a synchronous provider backed by fonts loaded by your application.
export async function inspectAndExport(document: PresentationDocument, measure: TextMeasurer) {
  const layout = {
    textMeasurement: { measure },
    diagnostics: { textOverflow: true, slideBounds: true },
  };
  const resolved = resolveLayout(document, layout);
  const diagnostics = resolved.analysis?.diagnostics ?? [];
  // The application decides whether warnings should block its workflow.
  if (diagnostics.length > 0) return { diagnostics };
  return { diagnostics, generated: await generatePptx(document, { layout }) };
}
```

## Explicit composition

See [alignment, distribution and stack composition](composition.md) for `LayoutOptions.composition`, coordinate/selection rules and failure behavior.

## Table sizing

See [measured table rows](table-sizing.md) for explicit automatic rows, merged-span constraints, overflow policies and `LayoutTableElement.cellLayout`.

## Intrinsic sizing

See [explicit content and parent sizing](intrinsic-sizing.md) for text/image/group dimensions, bounds, dependency cycles and atomic failure behavior.

## Table pagination

See [pagination and continuation](table-pagination.md) for repeated headers, merged blocks, continuation layouts, deterministic IDs and source row maps.

## Connector identity provenance

Resolved connectors optionally retain `connections.start/end` with the authored target `elementId` and explicit anchor, alongside their resolved point endpoints. These detached, format-independent references enable [native connector export](native-connectors.md); they do not change preview geometry or introduce OOXML IDs into Layout. Free endpoints have no connection identity.


Connector anchors use the target's local frame sites, applying target flips then rotation about the frame center. Resolution stays in the immediate group's coordinates; ancestor transforms apply downstream once. Explicit free endpoints/routes remain unchanged. Intrinsic group fitting uses the same transformed connector geometry as final resolution. Cross-group reference resolution remains outside this contract.

Optional [slide-bound event plans](slide-event-timeline.md) retain exact semantic targets and detached schedules in Layout. Owner/current target validation precedes transforms. Pagination that would split a timed slide remains unapplied with `pagination-event-timeline-unsupported`; a single-page table transform preserves the plan.

See [direct table borders](table-borders.md) for `resolveTableBorders`, `LayoutOptions.tableBorderLimits`, exact opt-in cell geometry, versioned edge metadata and frame/grid mismatch warnings.

Core resolves [table region styles](table-style.md) before measurement/pagination. Source body band formatting and repeated headers retain effective styles across fragments.

The [script font map extension](script-fonts.md#script-font-map-extension) selects language-aware concrete measurement fragments without changing author runs. Map fallback/missing diagnostics identify `scriptFontFamilies.<key>`; provider styles omit map rules.
