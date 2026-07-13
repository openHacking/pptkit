# Document Models

PPTKit should separate document state into three layers:

- `Authoring Model`
- `Canonical Presentation IR`
- `PPTX Package Model`

These layers should remain distinct even when the same package touches more than one of them.

## Authoring Model

The `Authoring Model` is the user-facing representation used by APIs and higher-level tooling.

It should be:

- ergonomic
- expressive
- easy to construct
- suitable for builders and future helpers

It should not be forced to mirror OOXML structures or package-level storage details.

Typical responsibilities:

- presentation construction
- slide creation
- element composition
- user-oriented defaults
- high-level authoring helpers

## Canonical Presentation IR

The `Canonical Presentation IR` is the normalized document form that the rest of the system should rely on.

It should be:

- stable
- normalized
- explicit
- independent from any one file format

Typical responsibilities:

- normalized slide and element structure
- resolved identifiers and references
- consistent geometry and style representation
- validation-friendly document state
- shared contract for preview, export, and parse-related transforms

It should not expose package-level OOXML concerns or try to preserve every authoring convenience directly.

## PPTX Package Model

The `PPTX Package Model` is the format-specific representation used for parsing and exporting PowerPoint packages.

It should be:

- format-aware
- package-oriented
- explicit about OOXML constraints
- allowed to model relationships, parts, and package structure directly

Typical responsibilities:

- package parts
- relationships
- content types
- slide XML and related resources
- media and embedded asset packaging

It should not become the project’s public authoring model.

## Why They Must Stay Separate

If the `Authoring Model` becomes too close to the `PPTX Package Model`, public APIs inherit file-format complexity too early.

If the `Canonical Presentation IR` is skipped, every exporter, preview path, and parser has to invent its own assumptions about structure and normalization.

If the `PPTX Package Model` is treated as just another view of the authoring model, round-tripping, parser preservation, and compatibility work become much harder to reason about.

## Concrete Flow Example

The intended flow is:

```text
authoring API input
        |
        v
 Authoring Model
        |
   normalize
        |
        v
Canonical Presentation IR
        |
     export
        |
        v
PPTX Package Model
        |
        v
      .pptx
```

For example:

1. A user builds a slide with text, image, and layout helpers through `@pptkit/core`.
2. The system normalizes those structures into canonical elements and references.
3. `@pptkit/layout` resolves placement against the canonical structures.
4. `@pptkit/layout` creates a detached export-ready layout result.
5. `@pptkit/pptx-exporter` translates that result into package parts, relationships, XML, and ZIP output.

## Design Constraint

The three-layer model is a north star for the architecture. It does not require every field and transform to be fully specified yet, but it should guide how future packages and docs evolve.
