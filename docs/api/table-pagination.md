# Table pagination and continuation

`LayoutOptions.tablePagination` explicitly splits a top-level slide table into row fragments and inserts continuation slides immediately after its source slide. It is a detached Layout transformation: authoring state and Canonical IR remain unchanged, and no filesystem, fonts or exporter-specific behavior enters the planner.

## Rule

Each `TablePaginationRule` supplies a unique `elementId` and an `area: Box` in slide-local logical points. The area must have positive finite width/height and finite extents. Its x/y replace the first table fragment's position; its height is the available space, not the emitted table height. The actual frame height is the sum of fragment rows. Table columns retain their widths, must match the source frame width and must fit both configured areas. Other source objects do not move automatically.

Optional settings:

| Field | Meaning |
| --- | --- |
| `headerRows` | Number of leading rows repeated on continuations; default 0. Must not cut through a row-spanning cell. |
| `continuation.area` | Available rectangle for subsequent fragments; defaults to the first area. Different width does not reflow or resize columns. |
| `continuation.layoutId` | Existing layout used by generated slides; defaults to the source layout. |
| `overflow` | `preserve` (default) rejects the rule if a header/body block cannot fit an empty page; `overflow` places each oversized block alone with a warning. |
| `nativeRowRegions` | `fragment` (default) copies native row switches to every fragment. `source-boundaries` disables native last-row styling on non-final fragments and disables native first-row styling on continuations without repeated headers; native banding can still restart. |
| `includeHidden` | Opt into hidden source slides/tables; generated pages retain their hidden state. |
| `maxPages` | Positive integer budget including the first fragment; default 1000. Exceeding it retains the entire source table for this rule. |

The initial supported contract allows one active paginated table per source slide. Multiple active rules on the same source reject all those rules, independent of their order. Layout/group-child tables and rotated/flipped tables are diagnosed. They are not silently moved into a different coordinate system. Empty tables and invalid merged grids are diagnosed.

## Rows, merges and overflow

Pagination uses resolved row heights. It computes all boundaries crossed by row spans and treats transitively overlapping spans as an indivisible block. Column spans and cell content/styles remain intact. A merged cell crossing the header/body boundary is rejected because repeating only part of it would be ambiguous.

Pages greedily take complete blocks in source order. A subsequent block moves to a fresh continuation when it no longer fits. A block too tall for an empty page rejects the entire table rule by default, including any already planned fragments. The explicit overflow policy retains the full block, marks the fragment with `pagination-row-overflow`, and never clips or splits its source rows. Header-only tables produce one fragment; zero-header tables do not invent header rows or blank pages.

Explicit cell `headerScope` values are copied with the cells: repeated column headers remain headers on each continuation, and a merged row header remains attached to its entire indivisible block. This does not infer header roles from `headerRows`. SVG exposes the explicit roles; PPTX export warns that those cell-level semantics are not exported.

Core region styles are materialized before pagination and travel with the source cells. Native PPTX table-style switches are different: each fragment is a separate native table and, by default, receives the same switches, so an application can reapply first-row, last-row or banding regions against fragment-local rows. A successful default-policy split with explicit native-style metadata and active native row regions reports `pagination-native-row-style-reapplied`; direct cell formatting can mask the effect. The exporter also enables first-row and banded-row switches when metadata is omitted. For these implicit defaults the diagnostic is emitted when no headers repeat, source/local row parity differs, or a merged row makes native band placement uncertain. Absence of a diagnostic does not establish native visual equivalence.

The opt-in `source-boundaries` policy retains an enabled native last-row switch only on the final fragment. It retains native first-row styling on the first fragment and on continuations that repeat at least one header row; without repeated headers it disables that switch on continuations. Other native flags and the source document remain unchanged. This preserves the logical first/last row boundaries but does not materialize native styles or preserve the phase of banded rows. When native row banding is enabled or defaulted on, the same diagnostic warns about its possible restart. A [PowerPoint control](../testing/native-pagination-row-style.md#opt-in-source-boundary-policy) checks sampled fills after separately removing direct cell formatting. One actual unpatched output was saved and reopened once with page text, flags, geometry and post-save/reopen PDF pixels checked; native editing and broader appearance remain unverified. Source-row materialization does not certify native preset appearance.

In one [PowerPoint observation](../testing/native-pagination-row-style.md), directly formatted PPTKit output masked the preset fill, while a separately source-patched control exposed the native last-row style on every continuation. This is a first-open sample, not an interactive or cross-platform acceptance result.

The first body block must fit the first area under the preserve policy even if a larger continuation area exists: this version does not create a header-only/empty first page to skip it. Increase the first area or explicitly choose overflow. Repeated headers count against every page's height.

## Stage ordering and failure

The pipeline is table sizing → intrinsic sizing → composition → pagination → final connector/image/chart resolution → analysis. This lets measured row heights drive pagination. Explicit pagination areas take precedence over table positions set by earlier composition. Composition rules do not run again on generated pages; use an existing continuation layout for repeated decorations.

Every table rule plans fragments and checks IDs, layout, budgets and finite geometry before committing. Failure leaves the table/slide at their pre-pagination geometry (including earlier successful sizing/composition). Other source slides can paginate independently. Invalid configuration throws before any stage; semantic failures become warnings at `options.tablePagination.<index>` with the source slide/table identity. Exporter/renderer forward them normally.

Successful fragments carry exact `LayoutTableElement.cellLayout` geometry. The SVG renderer uses these cell rectangles rather than allowing an HTML table to expand; PPTX uses the same row/frame dimensions. Fixed row height is not proof that text fits: use measured table sizing and text-overflow diagnostics where required. Native application font reflow remains separate compatibility work.

## Slide identity, templates and links

The first fragment retains the original slide/table IDs. Continuations get deterministic IDs derived from source slide ID, table ID and fragment ordinal. The allocator checks collisions against document, slide, layout, asset and nested element IDs. Conflicts are reported rather than resolved with order-dependent suffixes. Identical source/options yield identical IDs and pages on repeated resolution; Layout results are not Canonical IR and must not be fed back as normalized source documents.

Source sibling elements and notes remain on the source page. Generated slides contain only the table fragment plus the inherited layout's elements; they do not clone source slide decorations or duplicate speaker notes. Hidden state, section, tags and custom data are copied. Without an explicit continuation layout, background and layout provenance follow the source; selecting a layout uses that layout's background. Continuation table fragments keep name, accessibility, actions and styles, but are ordinary tables without the source `placeholderKey` association.

Existing slide actions remain anchored to their original target IDs. A jump to the source slide continues to mean its first fragment; repeated header actions retain their original targets. Export maps these IDs to the final slide order, so insertion does not change the destination. Source-page connectors follow the first fragment's final bounds; they are not cloned onto continuations. Standalone SVG still reports slide actions as metadata-only, as before pagination.

## Source traceability

When pagination is requested, `LayoutResult.pagination.fragments` contains one record for every successful fragment:

- `sourceSlideId`, `sourceElementId`: original source identities.
- `slideId`, `elementId`: resolved fragment identities.
- `pageIndex`: zero-based ordinal within the table's continuation sequence.
- `sourceRowIndices`: source row index for each output row, preserving compact cell order within the row.
- `repeatedHeaderRowCount`: zero on the first fragment, otherwise the number of repeated leading rows.

Every body row appears once; only headers repeat. The map remains available even for a table that fits on one page, and is empty for failed/skipped rules. It is Layout metadata, not hidden OOXML or a promised import/round-trip provenance format. Absolute output order is available from `LayoutResult.slides` (and renderer `index`); [field refresh](field-evaluation.md) can update existing field caches, including repeated headers, after pagination. It does not add fields to continuation pages. Fragment ordinals are stable for identical input, not immutable row identities across edits that change page boundaries.

## Diagnostics

| Code | Meaning |
| --- | --- |
| `pagination-target-unavailable` | Not a top-level table on a source slide. |
| `pagination-source-conflict` | Multiple active paginated tables on one source slide. |
| `pagination-transform-unsupported` | Table rotation/flips cannot use the requested slide-space rectangles. |
| `pagination-table-invalid` | Empty/invalid table, row heights, columns or merges. |
| `pagination-header-invalid` / `pagination-header-merged` | Invalid count or a merge crossing the header boundary. |
| `pagination-width-unavailable` | Column/frame mismatch or insufficient configured width. |
| `pagination-layout-missing` | Unknown continuation layout. |
| `pagination-row-too-tall` | Preserve policy cannot fit an indivisible block or header-only table. |
| `pagination-row-overflow` | Explicit overflow policy placed an oversized block; identifies the emitted fragment. |
| `pagination-native-row-style-reapplied` | Native row regions can be recomputed per fragment; under `source-boundaries`, native row bands may restart. |
| `pagination-page-limit` / `pagination-id-conflict` | Budget or deterministic identity conflict. |
| `pagination-height-invalid` | Nonfinite row sums or output extents. |

## Checked example

<!-- doc-test: docs/examples/table-pagination.ts -->
```ts
import { createPresentation } from "@pptkit/core";
import { resolveLayout, type LayoutOptions } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function createPaginatedReport() {
  const document = createPresentation({ size: { width: 600, height: 400 } });
  const slide = document.addSlide({ id: "report", notes: "Notes remain on the source page." });
  slide.addElement({
    type: "table", id: "data", box: { x: 30, y: 60, width: 540, height: 1148 },
    columns: [180, 180, 180],
    rows: Array.from({ length: 41 }, (_, index) => ({
      height: 28,
      cells: [0, 1, 2].map((column) => ({ content: index === 0 ? `Column ${column + 1}` : `Row ${index}, cell ${column + 1}` })),
    })),
  });
  const layout: LayoutOptions = {
    tablePagination: [{ elementId: "data", area: { x: 30, y: 60, width: 540, height: 300 }, headerRows: 1, maxPages: 50 }],
  };
  const resolved = resolveLayout(document, layout);
  if (resolved.analysis?.diagnostics.length) return { diagnostics: resolved.analysis.diagnostics };
  return { fragments: resolved.pagination?.fragments, generated: await generatePptx(document, { layout }) };
}
```

## Verification

[Layout tests](../../packages/layout/test/pagination.test.mjs) verify row conservation, merged boundaries, source detachment, repeated resolution, continuation metadata, IDs, limits, failures, hidden rules and measured-row integration. [Cross-package tests](../../scripts/pagination-parity.test.mjs) compare full PPTX bytes and exact-cell SVG to manually split decks and verify slide relationship targets. [Chromium tests](../../packages/svg-renderer/test/measurement.spec.mjs) check actual repeated-header text, cell rectangles and exported row heights on every generated page.

[Accessibility-tree tests](../../scripts/svg-accessibility-tree.test.mjs) check repeated column headers and a merged row header across continuation pages in Chromium, together with PPTX semantic-loss warnings. They do not establish PowerPoint screen-reader behavior.


The new [table structural transaction](table-editing.md) can produce a valid edited grid before measurement/pagination. Its mapping covers the authoring edit batch; pagination source-row mappings remain separate. Cell-level native/source references are not inferred.
