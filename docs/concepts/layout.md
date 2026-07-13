# Layout

Layout is a first-class transformation between normalized presentation semantics and format-specific output.

## Implemented behavior

`@pptkit/layout` currently:

- consumes authoring documents or Canonical IR v1
- returns a detached export-ready layout result
- resolves connector element anchors to points
- derives connector bounds
- resolves image `contain` and `cover` when source dimensions are known
- recursively processes group children while preserving local coordinates

## Why it is separate

Keeping layout outside Core and exporters makes placement behavior independently testable and reusable. Core remains focused on document meaning, while exporters remain focused on format translation and packaging.

## Roadmap behavior

Text measurement, overflow diagnostics, constraints, automatic composition, table measurement, and pagination belong in Layout but are not implemented yet. They must remain predictable transformations rather than hidden “magic” inside export.

See the [`@pptkit/layout` API](../api/layout.md) for callable behavior and [Layout Engine](../architecture/layout-engine.md) for contributor constraints.
