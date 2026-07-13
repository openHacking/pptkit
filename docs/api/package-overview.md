# Package Overview

This page summarizes the current public API surface by package.

## Current Packages

- `@pptkit/core`
- `@pptkit/layout`
- `@pptkit/pptx-exporter`
- `@pptkit/cli`

Current status:

- `@pptkit/core` provides the formal authoring model, asset registry, and normalization boundary.
- `@pptkit/layout` consumes normalized core documents through a placeholder layout pass.
- `@pptkit/pptx-exporter` consumes normalized core documents and returns a structured placeholder export result.
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
