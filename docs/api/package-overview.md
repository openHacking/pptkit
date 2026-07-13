# Package Overview

This page summarizes the current public API surface by package.

## Current Packages

- `@pptkit/core`
- `@pptkit/layout`
- `@pptkit/pptx-exporter`
- `@pptkit/cli`

Current status:

- `@pptkit/core` provides the formal authoring model, asset registry, and normalization boundary.
- `@pptkit/layout` consumes normalized core documents and returns a detached export-ready layout IR.
- `@pptkit/pptx-exporter` consumes the layout IR and writes a minimal editable `.pptx` package with structured warnings.
- `@pptkit/cli` remains a thin local workflow shell.

## Planned Follow-Up Packages

- `@pptkit/pptx-parser`
- `@pptkit/svg-parser`
- `@pptkit/svg-renderer`

## API Design Goals

- Small entry points
- Explicit naming
- Stable data flow between packages
- Minimal accidental coupling between authoring, layout, and file-format code

## Documentation Strategy

As additional packages are implemented, this section should expand into:

- Package-level references
- Function and type documentation
- Usage examples
- Migration notes for breaking changes
