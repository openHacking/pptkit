# Package Overview

PPTKit is a TypeScript monorepo organized by cohesive responsibility. Public authoring semantics flow from Core through Layout into format-specific exporters.

## Current packages

| Package | Public entry points | Responsibility |
| --- | --- | --- |
| `@pptkit/core` | package root | Method-managed authoring model, assets, themes, layouts, validation, Canonical IR v2, detached [finite-time arithmetic](finite-timeline.md), and [foreground event replay](event-timeline.md). |
| `@pptkit/layout` | package root | Detached geometric resolution over authoring or normalized documents. |
| `@pptkit/pptx-exporter` | package root, `/node` | Editable PPTX generation, asset loading, OOXML packaging, and Node file output. |
| `@pptkit/svg-renderer` | package root | Modern-browser SVG preview generation over resolved Layout state. |
| `@pptkit/pdf-renderer` | package root | PDF bytes from resolved Layout via a caller-owned Chromium browser. |
| `@pptkit/media-inspection` | package root | Bounded, environment-neutral media structure checks. |
| `@pptkit/pptx-parser` | package root, `/browser`, `/node` | Experimental bounded ZIP/XML/OPC, MCE and [PPTX structural inspection](pptx-inspection.md); explicit [import/no-op sessions](import-session.md), [partial content import](content-import.md), opt-in [existing text-box run saves](text-edit-save.md), [V2 shape/notes run saves](text-owner-save.md) and [V3 table-anchor run saves](table-text-save.md). The Node entry adds stable [file import and exclusive/snapshot-guarded publication](file-save.md). Broader preservation edits and rebuild remain planned. |
| `@pptkit/cli` | package root/bin | Read-only inspect/compare orchestration with public Parser reports; authoring commands remain unimplemented. |

## Data flow

```text
application
    │
    ▼
@pptkit/core authoring state
    │ validate + normalize
    ▼
Canonical Presentation IR v2
    │ resolve
    ▼
@pptkit/layout result
    │ output transform
    ├───────────────┐
    ▼               ▼
@pptkit/pptx-exporter   @pptkit/svg-renderer
    ▼               ▼
  .pptx        per-slide SVG preview
                    │
                    ▼
             @pptkit/pdf-renderer → PDF bytes
```

The exporter orchestrates normalization and layout for convenience, so ordinary callers can pass a `PresentationDocument` directly to `generatePptx()` or `writePptx()`.

## Planned packages

- `@pptkit/svg-parser` — SVG ingestion into format-independent visual structures.

The SVG parser remains roadmap direction and has no public API in the current repository. PPTX partial content import uses editable/preserved/fallback/unsupported reports in the existing Parser package; complete inheritance and preservation-aware editing remain future work.

## Design rules

- Core contains no filesystem, network, ZIP, XML, or relationship behavior.
- Layout contains no file output or OOXML serialization.
- Exporters consume normalized/layout contracts and do not own authoring defaults.
- Other packages are imported through public exports rather than private source paths.
- New capabilities ship as vertical slices across the owning packages, tests, examples, and documentation.

See [Package Boundaries](../architecture/package-boundaries.md) for contributor-level dependency rules.

Parser optionally compares [direct SmartArt paragraph vectors](diagram-text-comparison.md), retaining source indices and unknown freshness without Core write authority.

Parser also exposes the pure [capability observation view](import-capabilities.md); CLI can select it explicitly without adding source-save authority.


Parser also exposes opt-in [native timing inventory](native-timing-inventory.md), with bounded source observations separate from Core authored timelines.

[Direct table borders](table-borders.md) span Core styles, Layout edge resolution, Parser direct projection and SVG/PPTX output; source-style writes and full theme inheritance remain separate gates.

[Table region styles](table-style.md) are Core authoring/normalization; Layout owns shared edges/measurement/pagination and Exporter reports materialized direct-native formatting. Native table style inheritance stays in Parser/format work.

[Source table styles](table-style-import.md) reuse Parser theme/MCE infrastructure and stay independent from Core custom authoring rules. The two cascades have explicit separate priorities and native gates.
