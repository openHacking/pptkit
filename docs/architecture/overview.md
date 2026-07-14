# Architecture Overview

PPTKit is a TypeScript presentation document engine with explicit boundaries between authoring, normalization, layout, parsing, preview, and file export. Core, Layout, editable PPTX export, and foundational browser SVG preview are implemented; parsing remains roadmap work.

## Goal

Provide a developer-facing model for generating editable presentation output without forcing applications to manipulate low-level file-format structures.

## Three-layer model

PPTKit keeps three representations distinct:

1. The **Authoring Model** is ergonomic, method-managed application state.
2. **Canonical Presentation IR** is detached, validated, explicit, and format-independent.
3. A **PPTX Package Model** is format-specific parts, relationships, XML, media, and ZIP structure inside parser/exporter boundaries.

```text
application
    │
    ▼
Authoring Model (@pptkit/core)
    │ validate + normalize
    ▼
Canonical Presentation IR v1
    │
    ▼
Layout Result (@pptkit/layout)
    │
    ├────────► SVG preview (@pptkit/svg-renderer)
    │
    ▼
PPTX Package Model (@pptkit/pptx-exporter) → editable .pptx
```

A future parse flow starts at package-level data, classifies editable/preserved/fallback/unsupported structures, and maps supported semantics toward Canonical IR. Export and parse are related but are not assumed to be perfect inverses.

## Current package mapping

- `@pptkit/core` owns authoring semantics, identity, validation, and IR v1.
- `@pptkit/layout` resolves connector anchors and image fitting into detached layout output.
- `@pptkit/pptx-exporter` loads assets and serializes editable PPTX packages; its Node adapter writes files.
- `@pptkit/svg-renderer` creates browser-oriented per-slide hybrid SVG previews from Layout state.
- `@pptkit/cli` is a thin workflow shell without a stable authoring command surface yet.

The planned PPTX and SVG parser packages are described in [Package Boundaries](package-boundaries.md).

## Why the separation matters

Without these boundaries, mutable authoring assumptions leak into exporters, layout becomes hidden file-writing behavior, and parsers are forced to pretend arbitrary OOXML is equivalent to the public authoring model.

The separation allows PPTKit to:

- validate one canonical contract
- test layout independently from packaging
- support multiple output paths without duplicating authoring rules
- preserve unknown parsed content without contaminating common authoring APIs
- discuss editable output and visual fidelity as explicit tradeoffs

## Architecture priorities

1. Stable, intentional public semantics.
2. Complete validation and explicit normalized defaults.
3. Responsibility-based dependency direction.
4. Independently testable layout and serialization stages.
5. Honest degradation and preservation behavior.
6. Cross-application output verification.

## Current maturity

The repository is pre-release. Core IR v1, editable PPTX output, and a foundational browser preview pipeline exist, including rich text, themes, layouts/placeholders, images, shapes, connectors, groups, tables, notes, and metadata. Higher-fidelity text measurement, automatic layout, parsing, round-trip preservation, cross-browser visual baselines, and publication hardening remain active roadmap areas.
