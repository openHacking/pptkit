# `@pptkit/svg-renderer`

`@pptkit/svg-renderer` creates browser-oriented SVG previews from a PPTKit
`PresentationDocument` or an already resolved `LayoutResult`. It returns one standalone
SVG string per slide and never reads files, writes output, parses PPTX, or mutates its
input.

## Render an authoring document

```ts
declare function renderPresentationToSvg(
  document: PresentationDocument,
  options?: SvgRenderOptions,
): Promise<SvgRenderResult>;
```

The convenience entry normalizes the authoring document once, calls
`resolveNormalizedLayout()`, and renders the detached layout result.

## Render a layout result

```ts
declare function renderLayoutToSvg(
  layout: LayoutResult,
  options?: SvgRenderOptions,
): Promise<SvgRenderResult>;
```

Use this entry when a pipeline already owns normalized and resolved state. Both entry
points produce the same output for the same presentation.

## Results and warnings

```ts
interface SvgRenderResult {
  width: number;
  height: number;
  slides: SvgRenderedSlide[];
  warnings: SvgRenderWarning[];
  status: "rendered" | "rendered-with-warnings";
}

interface SvgRenderedSlide {
  slideId: string;
  index: number;
  hidden: boolean;
  svg: string;
}

interface SvgRenderWarning {
  code: string;
  message: string;
  slideId?: string;
  elementId?: string;
  assetId?: string;
}
```

Hidden slides remain in the result and are marked with `hidden: true`. Hidden elements
are omitted. Layout elements render before slide-local elements, preserving their
drawing order.

## Asset resolution

URL assets are used directly. Path assets produce a visible placeholder and an
`asset-path-unsupported` warning unless a resolver supplies a browser-safe URL:

```ts
type SvgAssetResolver = (
  asset: NormalizedAsset,
) => string | undefined | Promise<string | undefined>;

interface SvgRenderOptions {
  resolveAsset?: SvgAssetResolver;
}
```

Resolvers may return HTTP(S), data, or blob URLs. Resolver failures become warnings and
do not prevent other slides from rendering. Applications own blob URL lifetime.

## Supported preview semantics

- SVG geometry for the seven Core shapes, connectors, arrows, image transforms, crop,
  opacity, nested groups, backgrounds, fills, and strokes.
- Native SVG `<text>/<tspan>` output for simple text, with deterministic line breaks
  and baselines calibrated toward DrawingML output.
- XHTML `foreignObject` fallback for mixed-run rich text, shape text, and tables with
  row/column spans.
- Stable IDs derived from slide and element identity, accessibility labels, URL actions,
  and metadata for cross-slide actions.
- Explicit warnings for unsupported path assets, resolver failures, degraded image
  crop, `autoFit: resize`, unsafe URLs, and standalone cross-slide actions.

## Fidelity and security boundary

The output targets current Chromium, Firefox, and Safari. Native text is the portable
default for simple content. Remaining `foreignObject` content is not guaranteed in
non-browser SVG viewers, PowerPoint SVG import, or arbitrary rasterizers.
Browser and PowerPoint font engines differ, so previews help find clipping, overlap,
missing assets, structural mistakes, and obvious visual regressions; they are not
pixel-identical PowerPoint evidence.

Core point values are emitted as CSS pixel values inside XHTML `foreignObject` content.
This keeps text, margins, and tables in the same SVG user-unit coordinate system as the
surrounding point-based geometry; using CSS `pt` there would apply the browser's 4/3
point-to-pixel conversion a second time and make text overflow its authored bounds.

Simple text does not delegate wrapping or vertical placement to CSS. The renderer emits
explicit native SVG lines, applies normalized shrink scaling, and positions alphabetic
baselines from the authored top inset. Mixed-run and shape-text `foreignObject` nodes
keep vertical overflow visible to avoid silently cutting off glyphs; table cells remain
clipped to their explicit bounds.

`foreignObject` is not a universal requirement or a fidelity guarantee. It is a bounded
fallback until Layout can expose resolved mixed-run lines. Keeping the native and
fallback paths explicit avoids pretending that browser CSS and PowerPoint share font
metrics.

All authored text and attributes are escaped and unsafe URL protocols are omitted. The
renderer does not accept arbitrary SVG or HTML fragments. Applications should still
apply their normal Content Security Policy when inserting generated SVG into a page.

See the checked [browser example](../examples/svg-preview.ts) and
[Preview and Editing](../architecture/preview-and-editing.md).
