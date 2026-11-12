# Table cell formatting origins

`resolvePresentationInheritance(document)` reports effective table cell formatting through the same regional merge and normalization path used by ordinary output. No metadata enters Canonical IR. Table targets remain `{ kind: "element", id: tableId }`; groups and shared layouts retain their existing document-global element identity.

Properties use normalized cell positions:

- `rows.<row>.cells.<cell>.style.*` for fill, uniform stroke, sparse side candidates, margins, vertical alignment and explicit text direction.
- `rows.<row>.cells.<cell>.content.<paragraph>.style.*` for paragraph formatting.
- `rows.<row>.cells.<cell>.content.<paragraph>.runs.<run>.style.*` for run formatting and font-clear barriers.

Cell indices describe the compact normalized row, not physical column positions or stable cell IDs. The authoring source path points to the actual input occurrence; regional declarations instead point to `style.wholeTable`, band/edge/corner rules or border roles. Preset sources use quoted registry keys, for example `textStylePresets["cell.body"].frame.margin`. Equal values do not erase ownership. Reusing the same input cell object in multiple positions still records each occurrence separately.

## Cascade and sources

The existing [regional priority](table-style.md#rules-and-exact-priority) is unchanged: built-ins, whole table, border roles, column/row bands, edge/corner regions, the cell preset and direct cell/content overrides. Declaration sidecars travel through those raw merges before defaults are normalized. A materialized regional value is never attributed to the cell just because it now appears in a merged input.

Scalar margins point all four sides to the scalar source. Partial margins retain each selected declaration. Styled tables merge regional/preset/local insets field by field. Unstyled tables retain the existing rule: a cell margin declaration replaces the preset margin input as a whole, and omitted sides then use built-ins; the origin report describes that actual behavior.

Paint and bullet objects replace atomically. A solid paint without opacity selects 1 from a built-in path instead of retaining a lower paint's alpha. A `none` paint has only a type record. A new bullet without character/style/start uses the existing text built-ins. Uniform stroke properties patch all four logical side candidates before sparse side overrides; each resulting field retains the actual uniform/edge/role declaration. A missing edge field can point back to the uniform stroke or its built-in default. Table defaults use `defaults.table.cell.*`; text defaults use `defaults.text.paragraph.*` / `defaults.text.run.*`.

Explicit false, zero, opacity zero and default-valued declarations retain their sources. Optional font and script-map `null` clears keep `state: "cleared"`, including whole-map barriers followed by reopened individual slots. Never-declared optional fonts have no invented records. Effective run opacity 1 is reported even when normalized IR omits that default-valued field.

## Empty cells, edits and physical edges

Styled tables materialize unoccupied slots and empty insertion runs as before. Their formatting points to actual regions, presets or defaults, with no invented local cell declaration. Unstyled empty content retains its existing semantics. Normalized cells carry no source-format IDs. Structural edits recompute regions and their current paths on the next resolution; old reports remain frozen snapshots.

Origins describe **logical cell candidates**. Layout independently resolves conflicts on a shared physical border. For example, one cell's zero-width candidate may lose to a neighbor's nonzero candidate; its exported shared edge then differs from that cell's candidate while both origin and Layout reports remain correct. Use Layout's border grid to inspect the chosen physical edge. This API does not claim Office native table-style precedence.

Fresh PPTX still materializes regional rules and script-font maps with the existing warnings. The separate [quarter-turn cell direction](table-style.md) feature also reports its regional or local authoring origin. No reusable native style definition, source-style writer or physical-font selection capability is added. Parser source-format table-style provenance remains separate. Geometry, row/column structure and text content ownership are outside this formatting report. Full inheritance/template and table parent requirements remain open.

See the [typed example](../examples/table-format-origins.ts), [text origins](text-format-origins.md) and [verification](../testing/table-format-origins.md).
