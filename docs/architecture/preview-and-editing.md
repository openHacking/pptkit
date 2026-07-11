# Preview and Editing

Browser preview and future editing flows should operate on normalized document state, not on raw PPTX package state.

## Core Flow

The conceptual split should be:

```text
Canonical Presentation IR
        |
   +----+----+
   |         |
   v         v
SVG renderer PPTX exporter
   |         |
   v         v
preview    delivery
```

In other words:

- `IR -> SVG renderer -> preview`
- `IR -> PPTX exporter -> delivery`

## Why Not Edit Raw PPTX State

Raw package structures are too format-specific to serve as the main editing surface.

Using raw PPTX package state directly for preview or editing would make it harder to support:

- undo and redo
- structured editing
- AI-assisted modifications
- collaboration
- future non-PPTX outputs

It would also leak package-specific details into workflows that should operate on document semantics instead.

## Benefits of Editing IR

Using normalized document state for preview and editing gives the project a better foundation for:

- shared rendering logic
- deterministic updates
- future viewers
- diagnostics and validation
- cleaner separation between authored state and derived output

## Authored vs Derived State

One useful architecture lesson is the difference between authored state and derived state.

For PPTKit, the `Canonical Presentation IR` should be treated as authored or authoritative document state, while preview artifacts and exported packages should be treated as derived outputs.

That distinction helps prevent preview logic, exporter quirks, or package details from becoming the source of truth for the whole system.
