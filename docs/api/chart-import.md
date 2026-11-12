# Experimental chart content import

`importPptx(bytes, { projection: "content" })` maps bounded 2D bar, line, pie, doughnut, area, scatter, bubble and column/line combination profiles into public Core charts. Structure mode still omits drawings. This extends the [content projection](content-import.md) and [save-session rules](import-session.md); it is not complete chart appearance or preservation editing. [ADR 0002](../architecture/decisions/0002-chart-data-projection.md) records the source-priority decision.

## Relationships and identity

Parser resolves a graphic frame's chart relationship into a separate, typed `chartSpace` part, using the presentation's Strict/Transitional namespace profile and effective MCE tree. Namespace aliases and selected AlternateContent branches work. Missing/wrong-type/external/fragmented chart targets fail with a typed invalid-package error. Unknown graphic-frame and chart families stay source-only; supported neighbors still import. Single-family bar/line/pie/doughnut/area and the bounded [scatter](chart-scatter.md) and [bubble](chart-bubble.md) subsets are supported along with the bounded [column/line combination](chart-combo.md) profile. Other combinations, date axes, 3D and extended chart families remain source-only.

Frame identity/name/description/hidden state, local geometry/rotation/flips and supported object actions use the drawing profile. Chart properties retain their chart part/effective path, source drawing identity and owning Core element/slide IDs. Auxiliary style, color, user-shape, workbook and external relationships are inventoried as `chart.relationship` findings. They are not followed as rendering/execution instructions. A shared chart source projects independent Core values per frame; a shared workbook is inspected once per canonical source part.

## Data-source policy

Caches describe the stored chart snapshot. Embedded workbook stored cells are a separate source that may have diverged. No formula is recalculated and no external relationship is fetched. Workbook fallback is a data projection: a missing source cache may render differently in a native application. In [one PowerPoint 16.104 observation](../testing/chart-workbook-only-powerpoint-evidence.json), three charts without numeric caches displayed no plots on first open despite populated embedded workbooks and successful data projection. Excel was not installed; refresh and other environments were not tested.

| Source state | Result |
| --- | --- |
| Complete cache/literal, workbook agrees | Map cache; report agreement when a supported local range exists |
| Complete cache, workbook conflicts | Map cache; emit explicit conflict finding; source workbook remains in original bytes |
| Complete cache, unsupported formula/workbook data/external workbook | Map cache; report why workbook comparison is unavailable |
| Reference has no cache; complete compatible stored workbook range exists | Map stored values; report fallback and any reliance on cached formula-cell results |
| Reference has no cache; a supported explicit blank policy and a numeric workbook range matches the already parsed category count | Map absent or plain empty numeric cells as `null`; report fallback and block numeric preservation edits |
| Existing sparse numeric cache with a supported explicit blank policy, declared point count and matching dense category count | Map absent indices to `null`; cache remains authoritative, and numeric preservation edits are blocked |
| Other empty/sparse caches, even with a complete workbook | Whole chart source-only; do not repair holes or substitute zero |
| Missing cache and no complete compatible workbook range | Whole chart source-only, with reason |
| Invalid required relationship, malformed supported XLSX package or cache identity/count | Typed invalid-package failure; no partial session |
| Data/expanded-resource limits or cancellation | Typed limit/abort failure; later independent imports remain usable |

`strRef`/`numRef`, `strLit`/`numLit` and direct series-name `v` use expanded names. Points are sorted by `idx`, and series by explicit `order`, independently of XML order; source series indices/orders must be unique. Declared counts cannot cause unbounded allocation. Empty names/categories, incompatible per-series category arrays, category/value length disagreement, multilevel categories, error/nonfinite values and negative pie/doughnut values remain source-only. Zero and finite negative bar/line values are retained. Stacked bar and area sources whose cumulative values overflow the finite Core range remain source-only; original-byte no-op save retains them. Supported explicit policies are `gap` and `zero` for bar, line, pie and doughnut, and `span` for line. Standard area supports `gap` and `zero`, and stacked area supports `zero`; see [area charts](chart-area.md). Blanks remain `null`, distinct from zero, and arrays retain category positions. Unspecified or unsupported blank modes, missing counts, malformed values and mismatched counts are not converted to blanks; invalid inferred indices can reject the package as malformed. For category charts allocation is bounded by an already parsed category array; scatter sparse coordinates instead charge each position to the aggregate point budget. If a scatter cache is absent, a blank-aware workbook range may be bounded by the independently resolved counterpart; a present sparse or malformed cache is never repaired from the workbook. A complete conflicting workbook range is reported, while the cache remains the snapshot. Numeric lexical parsing accepts decimal/exponent syntax with XML whitespace; it does not use JavaScript's empty-string-to-zero conversion.

Embedded XLSX is inspected through the existing bounded ZIP/XML/OPC reader. Worksheet relationships/content types, inline/shared/rich strings, numeric/string stored results and cell identities are validated. Supported references are local quoted/unquoted A1 scalars or forward one-dimensional ranges, within worksheet bounds. Names, expressions, external references, 2D/reversed ranges, error/incompatible cells and unsupported spreadsheet MCE do not become data. Only a numeric range with a supported explicit blank policy bounded by the matching category count admits absent cells or numeric cells with no value/formula child and only supported attributes as `null`. An empty `<v/>` is not admitted. Formula cells without a stored value, error/boolean/text cells, and unknown cell metadata or content are not blanks. Formula cells use stored results only. The reader does not infer dates from number formats or evaluate hidden-row/filtered-data semantics. Their source properties remain unrepresented in Core.

Source formulas, numeric lexical spelling, number formats, source IDs and unsupported properties remain format-owned in the retained original bytes/inspection and located findings. Core contains only mapped values, never XML, workbook payloads or preservation handles.

## Geometry, styles and limits

Bar direction, supported grouping and expressible spacing map. Single-family bar and line charts use one validated category/value axis pair. Combination charts require two distinct reciprocal pairs and identical ordered categories across both families; the column scale maps to the primary value axis and the line scale to the secondary value axis. The opt-in [existing-values-v3 profile](chart-value-save.md) binds eligible combination values. Direct axis visibility/labels/ticks/strokes/gridlines and text styles map where representable; complete finite manual min/max/major-unit scales map when they produce at most 10,000 intervals. Partial bounds use reported Core auto-scale defaults. Pie has one nonnegative series and no Core axes. Dense opaque point colors are ordered by `dPt/idx`; partial/unsupported point colors stay source-only instead of inventing source colors.

Direct opaque series colors, line strokes/markers, chart/plot fills, supported legend placement and title text/style map. Mixed title formatting is flattened with a finding. Native manual chart layout, inherited styles/fonts/colors, number/date formats, labels, advanced axes/effects, filtering and unsupported properties are audited. Core recomputes geometry; mapped chart preview is always `fallback`. Omitted styling uses Core defaults and does not imply inferred source appearance.

| Content session limit | Default | Accounting |
| --- | --- | --- |
| `maxExpandedChartWorkbookBytes` | 16 MiB | Expanded streams in one distinct embedded workbook, including control streams |
| `maxTotalExpandedChartWorkbookBytes` | 64 MiB | Expanded streams across distinct inspected workbook parts; shared part inspected once |
| `maxChartDataPoints` | 100,000 | Cache/literal points, inspected workbook cells/shared strings and referenced cell reads; repeated ranges are charged to prevent data amplification |

Nested workbooks also obey captured ZIP/XML/OPC options and cancellation. Range traversal yields cooperatively and is bounded by actually inspected cells, rather than declared range size. These are resource-work limits, not a process-memory cap. Report/snapshot/source-retention budgets still apply. Nested package failures identify `outerWorkbookPart!/innerPart` when available.

## Editing and output

Mapped Core categories/series/styles can be edited and freshly exported. Exporter generates a separate new chart and embedded workbook per chart element; it does not retain source workbook sharing, formulas, unsupported styles or auxiliary parts. Both cache and workbook derive from the same current Core arrays. Series/category deletion and reordering regenerate dense indices and remove stale data. Three fresh rounds with duplicate series and category names also check positional data after simultaneous deletion, insertion and reorder; these do not grant source-preserving structural writes. Frame rotation/flips now use the shared PresentationML transform serializer.

The [stable LibreOffice chart edit fixture](../testing/corpus-charts-libreoffice-interactive-evidence.json) is a producer example of complete caches without embedded workbooks. PPTKit imports the edited column value and the line and pie controls, and unchanged sessions preserve the original bytes. Both numeric source-save profiles reject edits without workbook bindings. Fresh export generates new workbooks from the projected values; the integration test verifies their contents over three fresh export/reimport rounds, then confirms a numeric source edit on the fresh output. This does not recover the original workbooks or establish native editable reopen.

Fresh export writes positive and negative zero distinctly in both numeric caches and worksheet cells, so reimport retains the Core value's sign. Subsequent application saves may normalize numeric spelling or zero signs; native application behavior is not established by this serialization guarantee.

Chart series names and categories use OOXML `ST_Xstring` in chart caches and spreadsheet text. Import decodes each source text atom once before comparing cache and workbook, including shared and inline strings. Fresh export protects literal `_xHHHH_` patterns and represents carriage returns and XML-forbidden control characters with Xstring escapes in both sources. Isolated surrogates still reject. Microsoft documents the [chart value type](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numericvalue?view=openxml-3.0.1), [spreadsheet text type](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.text?view=openxml-3.0.1) and [Office escape restrictions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/d34ae755-c53f-4a44-a363-c6dd3ee018a4). In [one LibreOfficeDev Calc headless experiment](../testing/chart-xstring-calc-evidence.json), a literal overlapping `_x005F_x0041_` lost its `_x005F` portion after the first XLSX resave; other tested strings survived three resaves. This is not evidence for PowerPoint, GUI editing or other versions.

SVG cannot represent XML-forbidden characters directly. When a displayed chart title, category or legend label contains one, preview replaces it with U+FFFD and reports `chart-text-xml-replacement` for that chart; the source document is unchanged. Carriage returns use an XML character reference. This warning is scoped to displayed chart text and does not assert coverage of other SVG text owners.

Original-byte no-op remains exact only while the full Core baseline matches. Changes stale prepared plans. Default preservation edits to chart data/styles and session `rebuild` remain **blocked**. Explicit [chart save profiles](chart-value-save.md) admit eligible existing numeric values, category labels and series names; source-preserving series/category structural edits remain blocked. Fresh `generatePptx` output is a new supported-projection document. The unrelated opt-in [text-box run save profile](text-edit-save.md) retains chart/workbook payloads and source sharing without editing them. Consumers must review the import report before choosing fresh output.

<!-- doc-test: docs/examples/chart-import.ts -->
```ts
import { importPptx } from "@pptkit/pptx-parser";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function exportProjectedChartEdit(source: Uint8Array) {
  const session = await importPptx(source, {
    projection: "content",
    sessionLimits: { maxChartDataPoints: 100000, maxExpandedChartWorkbookBytes: 16 * 1024 * 1024 },
  });
  try {
    const slide = session.document.slides.find(s => s.elements.some(e => e.type === "chart"));
    if (!slide) throw new Error("No chart in the supported projection.");
    const index = slide.elements.findIndex(e => e.type === "chart");
    const chart = slide.elements[index]!;
    if (chart.type !== "chart" || chart.chartType === "scatter" || chart.chartType === "bubble" || chart.chartType === "combo") throw new Error("Expected a mapped chart.");
    const changed = { ...chart, series: chart.series.map(s => ({ ...s, values: s.values.map(v => v === null ? null : v * 2) })) };
    slide.removeElement(chart.id);
    slide.insertElement(index, changed);
    // Review importReport before adopting fresh output: unsupported source
    // properties are omitted. This is a new projection document.
    const fresh = await generatePptx(session.document);
    const preservation = await session.prepareSave({ mode: "edit" });
    return { fresh, importReport: session.report, preservation };
  } finally {
    session.dispose();
  }
}
```

## Evidence and remaining gates

[Independent fixtures](../../packages/pptx-parser/test/charts.test.mjs) cover both dialects, effective MCE/aliases/CDATA, indices, malformed/sparse data, cache/workbook agreement/conflict/missing cases, shared strings, stored formulas, unavailable spreadsheet MCE, external targets, nested failures and aggregate work/byte limits. [Integration tests](../../scripts/chart-import-parity.test.mjs) independently inspect series names, category and value caches, range references and every worksheet cell over three fresh rounds, including series deletion/reorder, duplicate names and categories, columns beyond Z, geometry and resource-size stability. Corrupted cache and range controls are rejected. These are fresh projections, not preservation edit-save.

[Chromium tests](../../scripts/browser-bundles.test.mjs) compare normalized documents, full reports and fresh/no-op bytes against Node, edit/reimport charts and verify no network requests. [Installed npm consumers](../../tests/consumers/node.mjs) exercise all three families through public exports. The optional [native runner](../../scripts/inspect-chart-import.mjs) and [recorded evidence](../testing/chart-import-evidence.json) independently compare LibreOffice-produced bar/line/pie cache data with Core, verify exact no-op and fresh cache data, and reopen fresh output to PDF. This is a producer/parser/headless-reopen experiment; interactive Edit Data, repair dialogs, PowerPoint and visual parity remain untested.

IMP-06.2, AUT-06.2 and their parents remain open: inherited dependency gates, number/date/missing-data/label semantics, complete appearance, independent broader comparison and native interactive editing are not complete. Shared-source preservation and chart-aware edit-save belong to IMP-07–09.

## Existing numeric source saves

Default projection and text profiles grant no chart save authority. The separate explicit [existing-values-v1 profile](chart-value-save.md) conditionally binds existing numeric cache/XLSX cells and validates atomic two-layer candidates. It does not admit structural/format changes or certify native Edit Data.


For supported fresh blank charts, [cross-package tests](../../scripts/chart-blank-parity.test.mjs) cover bar/line/pie, all-blank and mixed values, gaps in line geometry, stacked bars, cache/workbook serialization, three fresh rounds, both dialects and Node/browser parity. [PowerPoint 16.104 observation](../testing/chart-blank-powerpoint-evidence.json) records one sample per family on first open and after three sequential save/close/reopen rounds; AX distinguishes blank from zero and the line screenshot shows a gap. Native Edit Data was unavailable on this machine because Excel was not installed. [Additional PowerPoint evidence](../testing/chart-blank-modes-powerpoint-evidence.json) compares line gap, zero and span and bar/pie zero on first open and one save/reopen, including PDF vector geometry and preserved blank cells. Broader workbook blank interpretations, unspecified blank policies, formulas without stored numeric results, native data editing and complete format/label semantics remain open.

Doughnut projection additionally requires an explicit supported hole size; an absent first slice angle or angle value maps to zero; its geometry, multi-series rules, native observation and preservation-edit limits are documented in [doughnut charts](chart-doughnut.md).
