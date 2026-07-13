# `@pptkit/layout`

Detached layout resolution for Canonical Presentation IR v1.

`resolveNormalizedLayout()` clones normalized Core state into export-ready layout state. It resolves connector element anchors, recursively resolves group children, calculates image `contain` boxes and `cover` crops from asset dimensions, and keeps slide-layout content separate from slide-local content.

`resolveLayout(document)` is the authoring convenience entry point; it normalizes once and then delegates to the canonical resolver.

The package does not read files, download assets, serialize OOXML, or write output.

See the [Layout API reference](../../docs/api/layout.md) and [Layout Engine architecture](../../docs/architecture/layout-engine.md).
