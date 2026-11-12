# ADR0044: Source table theme-matrix references and composable placeholder colors

Status: accepted for an experimental solid/noFill and straight-line import subset, 2026-10-26. Native compatibility and source-style writes remain gated.

## Decision

Parser resolves source `fillRef`/`lnRef` through the target's independently selected format component. Reuse the existing theme relationship/override, MCE, color-map and ordered-transform machinery. Keep source-format identities and original nodes private; materialize ordinary Core paints/strokes and extend existing bounded provenance rather than adding a package, public matrix registry or new report version.

Extract cohesive paint/line interpretation into private `table-style-properties.ts`. The theme resolver owns component selection and lazy positional list indexing; table-style inheritance owns region selection and per-cell application. Environment IO stays in existing import infrastructure.

## Correctness boundaries

- A nearer present format component blocks an older matrix even when the selected list is absent. Color/font selection remains independent.
- Preserve every effective positional slot, including unknown nodes; never compact unsupported entries. Validate unsigned indices and compare to actual list length before selection. Fill 0/1000 mean noFill, fill 1001 starts background slot one, and line indices select only the line list; line zero projects noFill/width zero.
- Bind a reference color only when the selected entry actually requires `phClr`. Fixed entry colors ignore unused reference colors. Unsupported explicit entries block older properties.
- Retain private original RGB plus ordered modifiers for palette/reference/placeholder composition. Evaluate once through the existing color transformation engine, with final quantization only after all modifiers. Preserve alpha set/mod/off order. Never serialize this program into public records.
- Keep reference `/@idx` as selected source and actual theme/list/color/transform paths as dependencies. Audit unsupported theme attributes at their true theme part/path immediately, so general finding caps apply during traversal.
- Charge lazy list traversal and repeated property evaluation to shared work checkpoints; retained records/dependencies retain existing session limits. Cache selection by target theme scope; no cross-master resolved-value or audit contamination.

Neutral flat/single/center line attributes and straight-edge miter limits are supported. Other line semantics, non-solid paints, effects and script/default text appearance are separate admission work. Conditional region previews remain fallback.

## Evidence and consequences

See the [API contract](../../api/table-style-import.md#theme-format-matrix-extension), ten Parser groups, two browser/cross-package groups and installed consumers. Independent source DOM/Python checks verify selected slots and color programs, reject deliberate index mutation, and cover three scalar source-save rounds. A real rounding counterexample distinguishes `080808` from an incorrect intermediate-quantized `090909`.

[Five native controls](../../testing/theme-matrix.md) retain named-producer width, background-index and transform differences; a no-fill source raises an application error. Passing independent source interpretation does not imply native compatibility. Actual PowerPoint/theme/insert-row interaction remains untested for this subset. No-op and admitted scalar saves preserve original style/theme bytes; styles/metrics/structure still block atomically.

Index semantics follow Microsoft's [FillReference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillreference?view=openxml-3.0.1) and [LineReference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linereference?view=openxml-3.0.1) descriptions. The latter page's copied fill-list wording is not used to select line entries; the implementation uses `lnStyleLst`. Placeholder composition follows the DrawingML color program, rather than calibrating output to one producer's observed behavior.
