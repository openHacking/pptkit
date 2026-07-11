# Rendering Pipeline

This document describes the intended pipeline from authoring input to preview, export, and parse outputs.

## Pipeline Stages

The architecture should treat the system as multiple related paths built around normalized document state.

Shared stages include:

1. Authoring input is expressed through the `Authoring Model`.
2. The system normalizes that input into the `Canonical Presentation IR`.
3. Layout logic computes normalized placement and sizing information.
4. Directional transforms feed preview, export, or parse results.

## Authoring Path

```text
Authoring API
    |
    v
Authoring Model
    |
    v
Normalization
    |
    v
Canonical Presentation IR
```

## SVG Normalization Path

```text
Raw SVG
   |
   v
@pptkit/svg-parser
   |
   v
Normalized SVG structures
   |
   v
@pptkit/svg-renderer
   |
   v
Canonical Presentation IR or export-ready drawing structures
```

This path exists to keep SVG handling explicit instead of letting format-specific logic leak across the system.

## Preview Path

```text
Canonical Presentation IR
    |
    v
Layout Resolution
    |
    v
Preview Transform
    |
    v
Preview Output
```

The preview path should remain derived from normalized document state rather than becoming the system of record.

## Export Path

```text
Canonical Presentation IR
    |
    v
Layout Resolution
    |
    v
PPTX Export
    |
    v
PPTX Package Model
    |
    v
.pptx
```

## Parse Path

```text
.pptx
    |
    v
PPTX Package Model
    |
    v
Package Inspection
    |
    v
Normalization
    |
    v
Canonical Presentation IR
```

## Why Use a Pipeline

A staged pipeline gives the project better control over:

- Validation
- Error reporting
- Intermediate testing
- Feature fallback behavior
- Future support for other import or export paths

## Failure Handling

The pipeline should prefer predictable degradation over silent corruption.

Examples:

- Unsupported SVG features should surface clear warnings or documented fallback behavior.
- Export validation should fail early when the model cannot be represented safely.
- Parse failures should preserve diagnostics that help contributors reproduce issues.
- Unknown OOXML structures should be preserved where practical instead of silently discarded.

## Long-Term Value

Once this pipeline is stable, the project can grow more safely into:

- Better layout systems
- Richer asset support
- Import workflows
- Plugin hooks
- Benchmarking and regression tests
