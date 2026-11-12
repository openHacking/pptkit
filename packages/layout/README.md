# `@pptkit/layout`

Detached layout resolution for Canonical Presentation IR v2.

`resolveNormalizedLayout()` clones normalized Core state into export-ready layout state. It resolves connector element anchors, recursively resolves group children, calculates image `contain` boxes and `cover` crops from asset dimensions, and keeps slide-layout content separate from slide-local content.

`resolveLayout(document)` is the authoring convenience entry point; it normalizes once and then delegates to the canonical resolver.

The package does not read files, download assets, serialize OOXML, or write output.

See the [Layout API reference](../../docs/api/layout.md) and [Layout Engine architecture](../../docs/architecture/layout-engine.md).

Opt-in `LayoutOptions` support synchronous text/font measurement and detached overflow diagnostics. See the [Layout analysis contract](../../docs/api/layout.md#opt-in-measurement-and-diagnostics). The opt-in [DOM adapter](../../docs/api/browser-measurement.md) supplies browser font-backed measurement. [Explicit alignment/distribution/stacks](../../docs/api/composition.md) preserve fixed sizes. Explicit [intrinsic sizing](../../docs/api/intrinsic-sizing.md), [grids](../../docs/api/grid.md), [frame constraints](../../docs/api/constraints.md) and [table pagination](../../docs/api/table-pagination.md) are available as opt-in transformations.

See [intrinsic sizing](../../docs/api/intrinsic-sizing.md) for explicit content/parent dimensions and bounded group fitting.

[Direct table borders](../../docs/api/table-borders.md) add pure `resolveTableBorders`, cumulative `tableBorderLimits` and opt-in versioned physical edge/cell geometry shared with renderers.
