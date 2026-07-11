# Package Boundaries

This document defines the intended responsibilities of the initial PPTKit packages.

## Core Rule

Each package should have one primary responsibility.

If a package starts combining content modeling, layout, parsing, and file export, the architecture will become harder to test, harder to explain, and harder to evolve.

These boundaries should align with the three-layer model:

- `Authoring Model`
- `Canonical Presentation IR`
- `PPTX Package Model`

## Planned Packages

### `@pptkit/core`

Owns authoring-facing and shared document structures.

Expected responsibilities:

- Authoring Model construction
- Shared document structures
- Element definitions
- Shared types
- Public construction APIs
- Normalization entry points

### `@pptkit/layout`

Owns layout primitives and composition logic.

Expected responsibilities:

- Layout resolution against canonical structures
- Box and positioning abstractions
- Layout rules and constraints
- Size calculation helpers
- Normalized placement output for preview and export paths

### `@pptkit/pptx-exporter`

Owns transformation from normalized presentation data into `.pptx` output and package structures.

Expected responsibilities:

- PPTX Package Model generation
- Slide XML writing
- Asset packaging
- Export configuration
- Export validation

### `@pptkit/pptx-parser`

Owns transformation from `.pptx` input into normalized PPTKit-readable structures.

Expected responsibilities:

- PPTX Package Model inspection
- XML extraction and normalization
- Slide and asset inspection
- Parse diagnostics
- Preservation of unknown or partially understood structures where practical

### `@pptkit/svg-parser`

Owns directional SVG parsing into normalized structures.

Expected responsibilities:

- SVG structure parsing
- Shape and path normalization
- Attribute extraction

### `@pptkit/svg-renderer`

Owns translation of normalized SVG structures into presentation-compatible drawing instructions.

Expected responsibilities:

- Directional SVG-to-document translation support
- Rendering preparation for export or preview paths
- Fallback behavior for unsupported SVG features
- Output-oriented SVG transformation

### `@pptkit/cli`

Owns command-line entry points for local workflows.

Expected responsibilities:

- File-based generation commands
- Parse and inspect commands
- Local diagnostics
- Scriptable terminal UX

## Boundary Rules

- `@pptkit/core` should not contain PPTX XML or package serialization details.
- `@pptkit/layout` should not write files.
- `@pptkit/pptx-exporter` should not become the owner of the public authoring model.
- `@pptkit/pptx-parser` should not depend on exporter internals as if export and parse were mirror images.
- `@pptkit/svg-parser` and `@pptkit/svg-renderer` should not duplicate the full responsibility of the PPTX exporter.
- CLI concerns should stay out of core package APIs.

## Anti-Patterns To Avoid

- No element-type-per-package split such as text/image/shape/table packages.
- No raw PPTX XML concerns inside public authoring APIs.
- No browser editor state built directly on top of raw PPTX package structures.

## Why This Matters

These boundaries keep the project understandable for contributors and make it easier to:

- Review architecture changes
- Test components in isolation
- Avoid accidental lock-in to a single output format
- Add future capabilities without rewriting the foundation
