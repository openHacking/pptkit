# Presentation Model

The presentation model is PPTKit's format-independent vocabulary for presentation intent.

## What it represents

- presentation metadata, size, and theme
- ordered slides and reusable layouts
- text, images, shapes, connectors, groups, and tables
- assets and stable references
- placeholders, notes, hidden state, sections, tags, and custom data

The authoring model is concise and method-managed. Canonical IR is explicit and detached. Users normally create the former; Layout and exporters consume the latter.

## Why it matters

A stable semantic model keeps application APIs independent from PPTX XML, allows validation before environment-dependent work starts, and gives preview, export, and future parsing work one shared contract.

The model deliberately does not promise that every OOXML structure has a clean authoring equivalent. Future parsing will preserve or degrade unsupported package constructs at the format boundary.

## Continue reading

- [Core API](../api/core.md) for authoring calls and inputs.
- [Core Authoring Model](../architecture/core-authoring-model.md) for identity and state rationale.
- [Canonical Presentation IR v1](../architecture/canonical-ir-v1.md) for normalized invariants.
- [Document Models](../architecture/document-models.md) for the authoring/IR/package separation.
