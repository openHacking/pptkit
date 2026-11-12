# Atomic table structure edits

`editTable(table, operations, { frame, limits? })` is a pure Core transaction for existing compact table grids. It returns a detached, deeply frozen candidate and a complete mapping of original explicit cell anchors. It never changes the input, a document, Layout, source bytes or global identity state. Commit through the existing `slide.replaceElement(table.id, result.table)` only after the entire batch succeeds.

```ts
const result = editTable(table, [
  { type: "insert-row", index: 1, height: 24 },
  { type: "insert-column", index: 2, width: 60 },
  { type: "split", rowIndex: 0, columnIndex: 0 },
], { frame: "fit-grid" });
slide.replaceElement(table.id, result.table);
```

See the [typed example](../examples/table-editing.ts) and [ADR0040](../architecture/decisions/0040-atomic-table-structure-edits.md). Public root/browser exports include `editTable`, `TableEditError`, operations, options, limits, mappings and result types.

## Coordinates and operations

All row/column indices address the **physical grid**, not compact row cell arrays. Insertions use positions `0..length`; deletion and split use existing positions. Operations execute in array order against the preceding candidate state. Merge regions are rectangles `{ rowIndex, columnIndex, rowCount, columnCount }`, with positive counts. No implicit selection widening occurs.

| Operation | Ownership and dimensions |
| --- | --- |
| `insert-row { index, height, cell? }` | Require positive finite height. A merge crossing the insertion point expands its rowSpan; an anchor starting at the point moves after the new row. Uncovered new slots receive independent copies of a unit-cell template, default empty. |
| `insert-column { index, width, cell? }` | Same rule horizontally; require positive finite width. Templates cannot have non-unit spans. |
| `delete-row { index }` | Remove this grid row. Surviving merges shrink; if their original anchor row is removed, content/style stay in the surviving top-left slot. Fully removed cells lose their final target and map to deleted. |
| `delete-column { index }` | Same rule horizontally, including anchor-column removal. |
| `merge { region, content? }` | Every intersected merged cell must fit entirely inside the rectangle. Keep top-left cell style/preset. Default `reject` refuses nonempty or actionable content on other cells. Explicit `append-paragraphs` retains all selected cells' paragraphs in row-major anchor order, including empty paragraphs, run styles and actions. |
| `split { rowIndex, columnIndex }` | Select the owner even from a covered slot. Convert the complete merged rectangle to unit cells. Only its original top-left anchor keeps content; all new cells are empty and independently copy its cell style and text preset. Existing grid widths/heights stay fixed. |

Deleting the final row or column rejects; removing the table element is a separate authoring operation. For row or column deletion, optional `content: "reject"` aborts the entire batch if a fully removed anchor has nonempty text or an actionable run. A merged anchor that survives with a smaller span retains its content and does not trigger this guard. Omitted `content` and explicit `content: "discard"` preserve the existing deletion behavior; the returned mapping identifies removed original anchors. The `reject` guard does not protect formatting on otherwise empty cells. Use `content: "reject-data"` to also reject removal of an anchor with explicit cell style, a text preset, or structured paragraph content, even if its text is empty. Empty string cells without explicit style/preset and implicit blank slots remain removable; surviving merged anchors remain protected by promotion rather than rejected. This conservative policy treats even empty style objects and empty paragraph arrays as explicit data. The original input snapshot remains available for recovery.

Append requires identical cell `textStylePreset` names. This avoids silently applying the top-left preset to text that previously inherited a different preset. Materialize effective text styles before combining different presets. Strings become paragraphs using Core's CRLF/CR/LF rules; rich paragraph and run data are copied without flattening. Appending all empty cells can create multiple empty paragraphs, as explicitly requested. Cell fills, margins and other cell-wide style data use the top-left cell when merging; split cannot reconstruct discarded former cells. The unchanged source snapshot and mapping retain the information needed for an application-level undo operation.

## Grid, geometry and mapping

Input cells use the existing compact representation: each cell starts at the first unoccupied slot of its row; rowSpan/colSpan default to one. Out-of-bounds, overflowing or overlapping partitions reject. Omitted unoccupied grid slots are materialized as empty unit cells, matching existing fresh export behavior. These implicit blanks have no original source mapping.

Row heights omitted in the input are materialized once using the original explicit frame height divided by original row count, matching current Core normalization. Without a usable frame, provide explicit positive row heights. The utility does not resolve external layouts or placeholder geometry. Dimensions must be finite; grid-sum overflow rejects.

Geometry policy is mandatory:

- `frame: "preserve"` leaves the authored frame unchanged, including possible frame/grid mismatch.
- `frame: "fit-grid"` sets frame width/height to the final column/row sums, retaining its x/y (or zero if absent).

Root identity, name, actions, transforms, accessibility and other captured root data remain detached and preserved. Structural/style semantics beyond the existing Core table model are not invented; normal presentation validation still applies to style/action/preset values.

`cellMapping` contains one row per original **explicit** anchor, in original compact traversal order. `source` and `target` have `rowIndex`, compact `cellIndex` and physical `columnIndex`. Final status is retained, merged into another anchor, or deleted; deleted rows have no target. Inserted/implicit cells have no invented source identity. After merge then split, merged origins still map to the surviving top-left content anchor; other split cells are new blanks. Mappings describe this batch only and are not persistent cell IDs or OOXML bindings.

The result is frozen at runtime. Clone it before further direct edits, or pass it as the next transaction input. The table field retains the existing authoring input shape so it can be passed to `replaceElement`; operation/mapping reports expose readonly fields.

## Bounds and failure

Input, operations and templates must be finite plain data with enumerable own data properties. Dense arrays, accessor-free limits and acyclic data are required; getters are rejected without invocation. Unknown operation/option/limit fields reject. The descriptor contract is not a Proxy sandbox.

| Limit | Default | Supported maximum |
| --- | ---: | ---: |
| `maxRows` / `maxColumns` | 1,000 each | 10,000 each |
| `maxGridSlots` | 1,000,000 | 2,000,000 |
| `maxCells` | 100,000 | 200,000 |
| `maxOperations` | 1,000 | 10,000 |
| `maxDataValues` | 1,000,000 | 2,000,000 |
| `maxTextCharacters` | 16,000,000 | 32,000,000 |
| `maxWorkUnits` | 10,000,000 | 100,000,000 |
| `maxDepth` | 128 | 256 |

Rows, columns, slots and cells are bounded at every intermediate stage, even if a later operation would shrink them. Work covers capture paths, grid allocation/traversal, origins and operations; sorting charges a fixed n-log-n cost before execution, independent of the engine's sort algorithm. Captured values/UTF-16 text and key characters accumulate across input, operations, materialized templates/content and final output capture; these are accounting limits, not serialized bytes or a hard peak-memory guarantee. The synchronous transaction owns no timers, listeners, media or cancellation resources.

Failures throw `TableEditError` with a located `path` and code: `invalid-table-edit-input`, `invalid-table-grid`, `invalid-table-operation`, `table-edit-content-conflict` or `table-edit-limit-exceeded`. A failed batch returns no partial result and cannot alter document ownership. Independent retry requires no reset.

## Integration and evidence

Whole-object event targets retain the table's root ID when committed through `replaceElement`; this is not subcell animation/comment remapping. Layout remeasures/paginates the new grid through the existing opt-in sizing/pagination rules. Merge blocks remain indivisible. Whole-slide copy remaps root IDs normally; deleting the source slide does not invalidate its owned copy.

Fresh PPTX output and static SVG consume the edited table normally. This work also fixes an exporter bug: mixed row/column merges wider than two columns previously misplaced continuation records because a changing cursor was reused as the anchor column. The serializer now uses a stable anchor column. The [focused exporter test](../../packages/pptx-exporter/test/table-merge-continuations.test.mjs) checks spans 2/3/4/7, leading columns and every physical continuation slot.

- [13 Core groups](../../packages/core/test/table-edit.test.mjs) cover ownership, rich content/styles, boundaries, promotion, complete mappings, invalid partitions, input descriptors and exact budget boundaries; 200 generated inverse cases and 3,000 mixed edits compare against an independent physical-slot splice oracle.
- [Four cross-package groups](../../scripts/table-edit-parity.test.mjs) check three fresh export/import rounds, Node/ESM/global equality, independent Chromium DOM verification of every native physical merge slot and HTML anchor, Layout remeasurement/pagination, and Parser rejection/retry for source structural edits.
- Installed Node/type consumers exercise the public entry and explicit operation/frame contracts.
- The optional [independent/native runner](../../scripts/inspect-table-edits.mjs) and [manifest](../testing/table-edit-evidence.json) check three edit batches using Python ZIP/ElementTree, including deliberate continuation corruption. The named LibreOffice resaves retain grid/paragraph content; maximum observed width/height differences are below the explicit 0.1-point research tolerance. The last saved output reopens as a two-page PDF containing remaining/new text and controls; deleted text is absent. This is fresh native-table output, **not** source preservation or interactive editing evidence.

AUT-05.2.1/.2/.3 advance within the current direct style model. Full AUT-05.1 border/theme precedence, cell-level comment/animation identity, interactive native editing and source structural saves under IMP-09.3.4 remain open. Parser still blocks structural source mutations; restoring the imported baseline allows no-op or an admitted scalar text retry.

Explicit cell `headerScope` survives row and column edits. Merging cells with different scopes rejects atomically with `table-edit-header-conflict`; splitting a header propagates its scope to the resulting cells. The `reject-data` deletion policy also protects a header scope, including on an empty cell. A Chromium test checks resulting header and cell roles after merge, split and insertion, and compares the editing result between Node and the browser global bundle.

Native table preset and style switches are retained as detached table metadata during structural edits. Invalid preset names or non-Boolean switches reject the transaction before any result is returned. These switches do not imply that the native style definition itself is reusable.
