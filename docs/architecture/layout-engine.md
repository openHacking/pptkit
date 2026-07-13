# Layout Engine

`@pptkit/layout` owns presentation placement and sizing transformations that should not live in authoring state or file serialization.

## Current contract

The package accepts a `PresentationDocument` through `resolveLayout()` or Canonical IR through `resolveNormalizedLayout()`. It returns detached metadata, theme, assets, layouts, slides, and resolved elements.

Implemented resolution covers connector anchors/bounds, image contain/cover geometry, and recursive group processing. Layout neither mutates its input nor writes files.

## Expansion direction

Layout is the owner for future:

- text measurement interfaces and font metrics
- overflow detection and diagnostics
- alignment, spacing, stacking, and grid-like composition
- intrinsic sizing and constraints
- table measurement and pagination
- slide boundary checks and cross-page continuation

These capabilities are not current behavior and should be added as explicit, testable transformations.

## Controlled, not magical

Every layout capability must define its inputs, deterministic output, supported constraints, failure/overflow behavior, and measurement dependencies. PPTKit can learn from web layout without promising DOM or CSS parity.

## Non-goals

- full browser CSS compatibility
- arbitrary DOM fidelity
- implicit exporter-only auto-layout
- environment-dependent measurement hidden inside Core

See [Layout API](../api/layout.md) and [Rendering Pipeline](rendering-pipeline.md).
