# Package Overview

PPTKit is a TypeScript monorepo organized by cohesive responsibility. Public authoring semantics flow from Core through Layout into format-specific exporters.

## Current packages

| Package | Public entry points | Responsibility |
| --- | --- | --- |
| `@pptkit/core` | package root | Method-managed authoring model, assets, themes, layouts, validation, and Canonical IR v1. |
| `@pptkit/layout` | package root | Detached geometric resolution over authoring or normalized documents. |
| `@pptkit/pptx-exporter` | package root, `/node` | Editable PPTX generation, asset loading, OOXML packaging, and Node file output. |
| `@pptkit/cli` | package root/bin | Thin command orchestration; presentation-authoring commands are not stable yet. |

## Data flow

```text
application
    │
    ▼
@pptkit/core authoring state
    │ validate + normalize
    ▼
Canonical Presentation IR v1
    │ resolve
    ▼
@pptkit/layout result
    │ serialize + package
    ▼
@pptkit/pptx-exporter → .pptx
```

The exporter orchestrates normalization and layout for convenience, so ordinary callers can pass a `PresentationDocument` directly to `generatePptx()` or `writePptx()`.

## Planned packages

- `@pptkit/pptx-parser` — package-aware PPTX parsing with explicit editable, preserved, fallback, and unsupported tiers.
- `@pptkit/svg-parser` — SVG ingestion into format-independent visual structures.
- `@pptkit/svg-renderer` — SVG output from supported presentation structures.

These packages are roadmap direction and have no public API in the current repository.

## Design rules

- Core contains no filesystem, network, ZIP, XML, or relationship behavior.
- Layout contains no file output or OOXML serialization.
- Exporters consume normalized/layout contracts and do not own authoring defaults.
- Other packages are imported through public exports rather than private source paths.
- New capabilities ship as vertical slices across the owning packages, tests, examples, and documentation.

See [Package Boundaries](../architecture/package-boundaries.md) for contributor-level dependency rules.
