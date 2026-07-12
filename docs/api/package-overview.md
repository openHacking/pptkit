# Package Overview

This page summarizes the intended public API surface by package.

## Bootstrap Packages

- `@pptkit/core`
- `@pptkit/layout`
- `@pptkit/pptx-exporter`
- `@pptkit/cli`

These packages now exist in the workspace as provisional implementation shells.

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

As the packages are implemented, this section should expand into:

- Package-level references
- Function and type documentation
- Usage examples
- Migration notes for breaking changes
