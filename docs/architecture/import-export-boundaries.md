# Import and Export Boundaries

Exporting a PPTX and parsing a PPTX are related operations, but they are not inverse operations.

This boundary matters for both architecture and user expectations.

## Why They Are Not Inverses

PPTX packages contain behavior and structure that do not map cleanly back into a single normalized document model.

Examples include:

- master and layout inheritance
- theme mappings
- placeholders
- grouped coordinate transforms
- embedded media relationships
- compatibility extensions
- unknown or vendor-specific OOXML

Because of this, parsing is not just “undoing export.”

## Export Responsibilities

`@pptkit/pptx-exporter` should:

- accept normalized document state
- produce valid package structures
- enforce export-time constraints
- fail early when required invariants are not met

The first implementation supports text, images, rectangles, ellipses, and lines. Recoverable image loading failures are represented as warnings and the affected image is omitted; core normalization failures remain hard errors.

It should not assume that every PPTX in the wild could have been produced by the same normalized model.

## Parse Responsibilities

`@pptkit/pptx-parser` should:

- inspect package parts and relationships
- normalize understood structures into PPTKit-readable forms
- retain diagnostics about unsupported or partial results
- preserve unknown structures where possible

It should not assume that every parsed construct can be turned into a clean, fully editable canonical structure.

## Capability Tiers

Parser behavior should be thought of in capability tiers rather than a binary success/failure model.

Useful tiers include:

- fully normalized and editable
- normalized with partial degradation
- preserved but not fully understood
- unsupported and explicitly rejected

This framing helps the parser remain useful before full parity exists.

## Preservation-Oriented Thinking

When the parser encounters unsupported OOXML, silent data loss should be treated as a last resort.

Where practical, the system should preserve unknown structures or package fragments so that:

- round-trip damage is reduced
- diagnostics stay actionable
- future parser improvements can revisit preserved data

## Round-Trip Stability

Round-trip stability is an important long-term goal area, but it should not be documented as a blanket guarantee for every PPTX feature.

Instead, the architecture should assume:

- some structures round-trip cleanly
- some structures normalize with degradation
- some structures are preserved without full semantic understanding

That is a more honest foundation for future implementation work.
