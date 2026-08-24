# `@pptkit/layout`

`@pptkit/layout` converts validated Canonical Presentation IR into a detached, export-ready layout result. It currently resolves connector anchors and image `contain`/`cover` geometry. Fixed-width text auto-height is resolved earlier by Core normalization; Layout itself does not perform environment-dependent text measurement, paginate tables, write files, or serialize OOXML.

## `resolveLayout(document)`

```ts
declare function resolveLayout(
  document: PresentationDocument,
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
}
```

The result is detached from both authoring state and the input normalized document.

## Resolution behavior

### Connectors

Element-reference endpoints become absolute points using `top`, `right`, `bottom`, `left`, or `center` anchors. The default anchor is `center`. The layout result also contains a box derived from the start point, route points, and end point.

Core validation normally prevents missing connector references. If invalid IR is passed directly to `resolveNormalizedLayout()`, an unresolved endpoint falls back to `{ x: 0, y: 0 }`; callers should not rely on this fallback as a validation mechanism.

### Images

- `stretch` keeps the requested box.
- `contain` uses registered source dimensions to create a centered contained box, then emits `fit: "stretch"` for export.
- `cover` derives centered normalized crop values, then emits `fit: "crop"`.
- `crop` preserves caller-provided normalized crop values.

If source dimensions are unavailable, `contain` and `cover` remain unresolved rather than inventing an aspect ratio.

### Groups

Group children remain in their local coordinate system. Layout recursively resolves child connectors and images; the exporter applies the group box and `coordinateSize` transform.

### Charts

Chart layout is resolved once for preview and export. The resolved contract contains the
plot box, value scale, category anchors, series geometry, and legend item boxes. A right
legend resolves as one column and a bottom legend resolves as one row. Line charts use
`categoryAxisPlacement: "onCategory"`, so category labels, points, and major tick marks
share anchors. Bar charts use `"betweenCategories"` to keep the first and last bars fully
inside the plot; their label anchors and boundary tick anchors are stored separately.
Format-specific renderers map these semantics to native output without re-running chart
layout.

## Current limits

The current package does not perform environment-dependent text measurement, overflow management, constraints, automatic placement, table pagination, or cross-slide pagination. Core's deterministic fixed-width text height estimate is part of normalization, not a Layout measurement service. Native chart implementations may apply their own font metrics, but exporters must preserve the resolved legend flow and category-axis placement semantics.
