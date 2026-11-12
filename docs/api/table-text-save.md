# Experimental table-cell text preservation saves

`importPptx(source, { projection: "content", editSaveProfile: "text-runs-v3" })` adds fixed-topology character changes on existing mapped table anchor runs. V3 also admits the existing [V1 text-box](text-edit-save.md) and [V2 shape/notes](text-owner-save.md) operations. Earlier profiles retain their original owner scope and report shape; structure projection rejects all edit profiles.

This is the first implementation slice of **IMP-09.2.5**. It does not admit cell formatting, row heights, column widths, merge changes, insert/delete/copy of rows/columns, table styles, chart data or arbitrary table reconstruction. These mutations block the entire plan when mixed with supported characters.

## Source cells and Core anchors

DrawingML encodes one physical cell per grid column. Core represents only merge anchors; a vertically covered row can have an empty `cells` array. The existing table reader validates the complete grid/merge occupancy before collecting any run bindings. It associates each mapped anchor with its physical source cell, not with a guessed flattened index or matching text.

V3 adds frozen `tableTextRuns: readonly ImportedTableTextRunSource[]`, including an empty array when no table anchors map:

| Field | Meaning |
| --- | --- |
| `slideId`, `elementId` | Owning Core slide and table element; cells receive no fabricated element IDs |
| `rowIndex` | Zero-based source/Core row |
| `cellIndex` | Zero-based compact Core anchor index within that row |
| `columnIndex` | Zero-based physical source grid column at the anchor |
| `paragraphIndex`, `runIndex` | Zero-based normalized cell run location |
| `source` | Source part, effective anchor-cell path, table source object ID and raw text-element index when available |
| `sourceElementIndex` | One-based raw element preorder including inactive MCE markup |
| `status`, `reason` | Bound character atom or explicitly unmapped run |

For a 2×2 merge at the top-left of a three-column table, the first cell of the second Core row can therefore have `cellIndex: 0` and `columnIndex: 2`. Save paths use the Core indices, such as `/slides/0/elements/4/rows/1/cells/0/content/0/runs/0/text`; raw identities locate the source atom.

Table records are separate from drawing `textRuns` and slide `notesTextRuns`. V3 emits all three arrays. V1/V2 do not gain a table array or table-save permission. Private baseline text, namespace, source nodes and owner discriminators are omitted from public records and Core/IR.

## Admission and preservation

Only existing bound run characters may change. Paragraph/run topology, row/cell counts, spans, grid/frame metrics, styles, actions, identity/order and every other normalized field remain equal to the imported baseline. Direct empty `<a:t/>` atoms can receive characters. Missing text bodies, synthetic empty paragraphs and breaks are unmapped; mixed atoms containing comments/children/processing instructions block editing.

Unmapped/sparse/nonpositive tables and tables whose merge continuations contain text/fields/breaks remain source-only. They acquire no bindings even when supported sibling objects are editable. Unknown table/cell extensions, continuation styling, source table-style IDs, shared themes and field markup remain retained.

Per-part raw guards apply to the table's source drawing identity. A whole-table target can remain unchanged; cell/range/sub-object or opaque targets on an edited table block all edits to that table until range remapping is verified. Inactive targets are scanned too. Equal text in different cells, tables or inactive branches is not treated as identity. Shared raw atoms still require all mapped views to agree.

The [existing whole-plan pipeline](text-edit-save.md#candidate-construction-and-validation) performs lexical XML replacement, ZIP patching, exact candidate package differences and actual-output reprojection. Only approved slide/notes XML payloads may change. Complete change reports classify scalar grid metrics as geometry and spans as structure; they do not authorize those operations. Failures return no candidate. Save never rebases the source.

Drawing, notes and table provenance share `maxReportFindings`; no new unbounded report or budget is introduced. Existing analysis/change/output/XML/validation budgets, snapshot freshness, cancellation/disposal, retry and no-op restoration apply.

## Checked example

The example edits the first run of a selected compact anchor in the first mapped table. Consumers can inspect provenance first; `prepareSave` remains authoritative for source guards and the whole candidate.

<!-- doc-test: docs/examples/table-text-save.ts -->
```ts
import { importPptx, type PptxSaveReport } from "@pptkit/pptx-parser";

export async function editImportedTableRun(
  source: Uint8Array,
  rowIndex: number,
  cellIndex: number,
  text: string,
): Promise<{ bytes: Uint8Array; report: PptxSaveReport }> {
  const session = await importPptx(source, {
    projection: "content",
    editSaveProfile: "text-runs-v3",
  });
  try {
    const slide = session.document.slides.find(item =>
      item.elements.some(element => element.type === "table"));
    const element = slide?.elements.find(item => item.type === "table");
    if (!slide || !element || element.type !== "table") {
      throw new Error("No mapped source table.");
    }
    const updated = structuredClone(element);
    const cell = updated.rows[rowIndex]?.cells[cellIndex];
    if (!cell || !Array.isArray(cell.content) || !cell.content[0]?.runs[0]) {
      throw new Error("No existing first run at that compact anchor.");
    }
    cell.content[0].runs[0].text = text;
    const index = slide.elements.indexOf(element);
    slide.removeElement(element.id);
    slide.insertElement(index, updated);
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    }
    return await session.save(prepared.plan);
  } finally {
    session.dispose();
  }
}
```

## Verification and open gates

- [Parser tests](../../packages/pptx-parser/test/table-save.test.mjs): both dialects, compact/physical merge identity, 24 independently bisected grids, multiple/nested tables, MCE aliases/inactive duplicates, empty/mixed/synthetic atoms, source actions/extensions, range guards, unsupported mutations, shared-source conflicts, limits/retry, undo/freshness, UTF-16 and three rounds.
- [Mixed-owner integration](../../scripts/table-save-parity.test.mjs): two tables, shape/notes text, charts/workbooks/images, independent payload hashes, graph identity, exact two-part differences and full candidate reprojection over three rounds.
- [Browser ESM/global](../../scripts/browser-bundles.test.mjs): Node-identical bytes/reports/documents; DOM independently locates physical row/column anchors; no network requests. Installed [Node](../../tests/consumers/node.mjs) and [declaration](../../tests/consumers/types.ts) consumers verify the actual public entry points.
- [Native runner](../../scripts/inspect-table-save.mjs) and [evidence](../testing/table-save-evidence.json): Python ZIP/XML verifies exact approved atoms, unchanged non-text XML, physical grids/merges and all unselected payload bytes. LibreOfficeDev headless resave retains both DrawingML tables and edited/control text; a two-page PDF contains both edits. The producer changes grid quantization/redundant merge attributes and drops unknown extension/orphan data retained by PPTKit.
- [PowerPoint interactive control](../testing/native-v3-interactive.md): one named mixed sample has actual edits of a merged anchor and a separate physical-column anchor, three native reopen rounds, a post-native PPTKit edit and three further native rounds. PowerPoint drops an orphan and unknown cell attribute and rewrites continuation attributes; the specific merge occupancy remains. First interactive/final pixels agree, while glyph positions differ slightly.
- [Interleaved LibreOffice control](../testing/table-interleaved-native.md): three PPTKit edits interleaved with named headless application resaves retain the named merge/anchor semantics. Each application resave reduces an unrelated image's two dimensions by 360 EMU; the starting native input had already lost its workbook, orphan and opaque extension.

Reproduce the optional native experiment after building with `node scripts/inspect-table-save.mjs [soffice-executable]`. Headless table/text retention and the named PowerPoint control do not prove general interactive editability, repair-dialog absence or complete appearance/implicit-reference fidelity. Those gates, direct style/metric writes and structural operations remain open; IMP-09.2.5 and its parents stay unchecked.
