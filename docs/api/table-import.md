# Experimental table content import

`importPptx(bytes, { projection: "content" })` maps a bounded DrawingML table subset into ordinary Core tables. It uses the existing [content import](content-import.md), source/property reports, budgets and [save session](import-session.md). Structure mode continues to omit all drawing content.

## Geometry and merged cells

The reader recognizes table graphic-data URIs from Transitional and Strict namespaces and uses expanded element names rather than prefixes. A table's drawing identity, name, accessibility description, hidden state and supported direct object action use the same orchestration as other drawings.

Local `p:xfrm` frame, rotation and flips map to points/degrees. Group origins are rebased once. Grid column widths and row heights remain authored values; the reader does not average missing dimensions or resize the frame to their sums. A discrepancy greater than one EMU marks the mapped-fields preview as `fallback` and emits a located finding. This does not implement native table auto-sizing.

Each source row must contain one physical cell per grid column. `gridSpan`/`rowSpan` anchors and `hMerge`/`vMerge` continuations become Core's compact anchor-only rows, including an empty `cells` array when the row is entirely covered by vertical merges. The reader validates all merge rectangles before interpreting anchor text/actions; occupancy memory is proportional to the number of columns. Redundant continuation spans are accepted only when consistent with their anchor.

| Source condition | Result |
| --- | --- |
| Positive authored grid dimensions and valid rectangular merges | Editable grid/anchor subset |
| Missing optional cell text body | Empty Core content; direct cell styles still map |
| Nonpositive row/column sizes, zero-area frame, sparse/excess physical cells | Entire table source-only with reason |
| Continuation text, including whitespace, fields or breaks; multiple paragraphs; explicit empty-text formatting, frame/list properties, actions or unsupported text markup | Entire table source-only; no silent removal or concatenation into the anchor. An otherwise empty continuation with only the writer's `en-US` insertion-language marker remains supported. |
| Missing grid/rows, malformed mapped numbers/booleans, duplicate singleton nodes | Typed invalid-package failure |
| Out-of-grid span, overlapping rectangles, orphan or inconsistent continuation | Typed invalid-package failure; no partial session |
| Unknown graphic-data family | Source-only object; supported sibling drawings still import |

Dimensions use checked integer EMUs before point conversion. Coordinate arithmetic can differ at floating-point precision below one EMU; native evidence compares rounded integer EMUs and 1/60000-degree rotation units. XML and report limits apply to the complete source and interpretation/audit traversal; import cancellation yields no usable partial document/session.

## Text and direct style

Anchor text reuses the existing rich-text reader: paragraphs/runs, lexical Unicode, empty styled paragraphs, breaks, supported lists, direct styles and safe run clicks. Table/slide actions resolve against source relationships without fetching or activating them. Findings retain the owning table's Core/source identity and effective cell/run path; cells do not acquire fabricated Core element IDs.

Direct sRGB/none fills and alpha map to Core paint. Four equal direct edge strokes map to Core's uniform cell stroke, including the exporter's `sysDot` dash encoding. Supported incomplete/unequal direct edges map to sparse [per-edge borders](table-borders.md). Uniform merged outer segments aggregate; heterogeneous or partly absent segments remain source-only with findings. The [source table-style profile](table-style-import.md) now selects inline/reference/default styles and maps supported direct regions/theme colors/Latin defaults with separate provenance. Conditional order remains preview fallback; diagonal lines, style-matrix references and other unsupported properties remain audited. Direct fields do not establish full table appearance.

Standard `tcPr` margins and top/middle/bottom anchoring override compatible legacy PPTKit body insets/anchoring. Nonnegative values map; negative insets receive findings. Other text-frame settings, cell direction, unknown line modifiers and continuation-cell styling remain source-only. Omitted/inherited values use reported Core defaults or neutral paint; unsupported theme/style/default appearance is not inferred; the bounded source-style profile separately records what it actually resolves. Direct per-edge borders and pure structural edits have scoped implementations; full table-style/theme precedence and native interaction remain AUT-05 work.

Fresh rotated/flipped tables now use a native group transform with an editable, locally positioned table child, because Office ignores graphic-frame rotation/flips. Authored root identity/action stay on the group and `table-transform-grouped` reports the hierarchy change. Identity transforms retain flat frame serialization. See the [border/native contract](table-borders.md); native producer transform fidelity remains partial.

## Operations and save boundary

| Operation | Assurance |
| --- | --- |
| Create/import/edit/preview/fresh export | Editable **mapped fields**; property findings still describe omissions; frame/grid mismatch marks preview fallback |
| No-op save | Exact original source bytes, including unsupported table properties and other objects |
| Table edit preservation save | Default/V1/V2: blocked. Explicit [V3 table-character profile](table-text-save.md) admits existing bound anchor runs with fixed topology and candidate validation; broader style/metric/structure changes remain blocked |
| Session rebuild | Unsupported/blocked |

A Core table may be edited or copied and exported as a new projection document. That output contains the supported projection and exporter defaults, rather than untouched source table styles/extensions. This is distinct from source round-trip. Changes invalidate prepared no-op plans; restoring the complete baseline restores eligibility. Copied self-slide links in cell text follow Core's reference-aware copy behavior.

The example returns an explicit fresh output and the separately blocked preservation result; consumers review the source report before adopting the fresh output.

<!-- doc-test: docs/examples/table-import.ts -->
```ts
import { importPptx } from "@pptkit/pptx-parser";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function exportProjectedTableEdit(source: Uint8Array) {
  const session = await importPptx(source, { projection: "content" });
  try {
    const slide = session.document.slides.find(s => s.elements.some(e => e.type === "table"));
    if (!slide) throw new Error("No table in the supported content projection.");
    const index = slide.elements.findIndex(e => e.type === "table");
    const table = slide.elements[index]!;
    if (table.type !== "table") throw new Error("Expected a table.");
    const rows = table.rows.map((row, ri) => ({
      ...row,
      cells: row.cells.map((cell, ci) => ri === 0 && ci === 0
        ? { ...cell, content: "Edited anchor text" }
        : cell),
    }));
    slide.removeElement(table.id);
    slide.insertElement(index, { ...table, rows });
    // This exports the supported Core projection. Unsupported source content
    // is not included, so callers must review the import report first.
    const fresh = await generatePptx(session.document);
    const preservation = await session.prepareSave({ mode: "edit" });
    return { fresh, importReport: session.report, preservation };
  } finally {
    session.dispose();
  }
}
```

## Verification and remaining work

[Independent XML fixtures](../../packages/pptx-parser/test/tables.test.mjs) cover both dialects/URIs, MCE aliases, geometry and nested origins, merge/continuation errors, optional/empty bodies, rich text/actions, style precedence, unsupported properties, budgets/cancellation, stale plans and restoration. Forty independently bisected rectangular partitions check anchor membership and spans against a brute-force geometric oracle.

[Cross-package tests](../../scripts/import-content-parity.test.mjs) verify three fresh export/import rounds, physical-grid/merge/transform/style/action XML, cell edits, copied self-links and stable resource size. [Browser tests](../../scripts/browser-bundles.test.mjs) compare complete normalized documents, reports, output bytes, actual rendered cell pixels, edits and exact no-op restoration with no network requests. [Installed consumers](../../tests/consumers/node.mjs) use published tarballs and public declarations/entries.

The optional [native runner](../../scripts/inspect-table-import.mjs), invoked as `node scripts/inspect-table-import.mjs [soffice-executable]` after building, compares LibreOffice-produced tables against independent Python ZIP/XML grid/anchor/text/transform inventories, checks original no-op bytes and fresh output, and reopens the latter to PDF. The [recorded evidence](../testing/table-import-evidence.json) separates producer changes from Parser projection. The separate [V3 preservation experiment](../testing/table-save-evidence.json) covers existing anchor characters only. Native interactive table editing, PowerPoint behavior, theme/table-style fidelity, visual parity for grid/frame mismatch and broader preservation-aware edits remain unverified or unimplemented. IMP-06.1 and AUT-05 remain open for their full gates.
