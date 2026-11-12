# Experimental existing chart numeric preservation saves

Explicit `chartEditSaveProfile: "existing-values-v1"` with `projection: "content"` admits existing finite numeric values in mapped single-family 2D bar/line/pie/doughnut charts and the bounded dense scatter subset. Scatter bindings carry `coordinate: x` or `coordinate: y`, and each independently identifies its cache point and workbook cell; the [CLI v2 manifest](cli-save.md) addresses them independently while v1 remains unchanged. For doughnuts, a supported explicit hole size and dense bound numeric sources are required; absent first-slice angles or values project as zero while their source form remains unchanged in source saves; see [doughnut charts](chart-doughnut.md). It is independent of `editSaveProfile` and can combine with `text-runs-v3` for one atomic mixed candidate; `firstSlideNumberSaveProfile: "attribute-v2"` may additionally admit the starting-number attribute in that candidate. `fieldTextProfile` can project supported fields alongside these edits; field identity and cached text remain read-only. Structure projection and invalid profile values reject. Default import, text profiles, no-op authority and fresh export retain their existing contracts.

The opt-in `existing-values-v2` profile additionally admits existing nonblank X/Y points in a sparse scatter chart when the indexed cache and workbook agree at every position, including blanks. It preserves blank positions and edits only existing numeric scalar cache points and workbook cells; changing a blank into a number or a number into a blank is rejected. Missing caches, conflicting data, shared workbook references and unsupported dependencies remain blocked. Three programmatic save/reimport rounds cover both fresh fixtures and a PowerPoint-saved independent source. Four edited fixed-axis samples also have [one PowerPoint save/reopen observation](../testing/chart-scatter-sparse-save-powerpoint-evidence.json), with cache, workbook and retained UI coordinates checked; visual geometry and interactive workbook editing were not measured. The existing v1 profile keeps its original sparse-data restrictions. The [CLI v3 chart manifest](cli-save.md) explicitly selects the v2 profile; older manifest versions retain v1 behavior.

The opt-in `existing-values-v3` additionally binds existing dense X, Y and size values in the supported 2D [bubble profile](chart-bubble.md). Each `coordinate: "x" | "y" | "size"` identifies a cache point and an independent cell. X and Y replacements must be finite; sizes must also be strictly positive. Missing or padded caches, workbook conflicts and cache-only native saves remain unbound. V1 and V2 keep their previous capabilities, and the [CLI v4 manifest](cli-save.md) selects V3. One [LibreOffice native reopen](../testing/bubble-source-save-native-evidence.json) confirms the edited bubble centre and diameter; its subsequent native resave removes the workbook.

The opt-in `existing-values-and-labels-v4` also preserves supported category-label edits. Every affected series cache and its unshared embedded-workbook cell must agree, and the workbook cell must be a direct simple inline string with no unaccounted references. The import report's `chartLabels` lists each projected category's status, reason when unmapped, and observed cache/workbook source pairs; pairs on an unmapped record may be incomplete. A `bound` status is eligibility, not a ready save plan. For a replacement with leading or trailing XML whitespace, the save inserts or updates the worksheet text’s `xml:space="preserve"` when necessary; unsupported bindings are blocked. The save report includes a `chartLabel` before/after value and category index. The Parser API and CLI chart manifest v5 select this profile. See the [source research](../testing/chart-label-source-research.md) for the current evidence and native-validation limits.

The Parser option `existing-values-and-series-names-v5` additionally admits existing series names in categorical, scatter and bubble charts when the direct string cache and unshared workbook cell have proven ownership and agree. `chartSeriesNames` reports source eligibility, and save changes identify `chartSeriesName`. Name, category and numeric edits may share a candidate; aliased or unsupported name sources remain blocked. CLI chart manifest v6 selects this profile. A [PowerPoint reopen observation](../testing/chart-series-name-powerpoint-reopen.md) covers one generated two-series chart across three Parser saves; One PowerPoint resave retained the names and workbook bindings; Excel editing remains unverified.

The same v3 profile supports existing dense numeric values in the [column/line combination profile](chart-combo.md). Bindings use the global sorted series index and category point index; they have no coordinate field. Both families can be changed in one transaction. The native cache-only LibreOffice fixture stays unbound, and v1/v2 do not grant combination write authority.

This is a bounded implementation slice of IMP-09.2.6.1/.2. Categories, series names, ordering/counts, formulas/ranges, styles, labels, geometry and all other normalized fields must remain unchanged. Fixed chart axis bounds are retained rather than recomputed; an edited value outside those bounds can be clipped in the native application, as observed in the [named PowerPoint sample](../testing/chart-value-save.md). Mixed unsupported mutations block the entire plan. Interactive Edit Data and full native appearance/reference fidelity remain untested for this save profile; the parent tasks remain open.

[Chart label source research](../testing/chart-label-source-research.md) records a generated two-series case with two category caches referencing one worksheet range and one-pass escaped strings. It does not enable label writes; shared ownership, string forms and native verification remain open.

## Eligibility and source identity

Opt-in reports add frozen `chartValues: readonly ImportedChartValueSource[]`. Each record names Core `slideId`, `elementId`, zero-based sorted `seriesIndex` and `pointIndex`, chart `source` and optional one-based raw `sourceElementIndex`, optional `workbook`, and `status: "bound" | "unmapped"` with a reason. The chart source path identifies its value container; the raw index identifies the exact scalar `c:v`. Workbook provenance includes outer `part`, inner `worksheetPart`, A1 `cellReference` and raw scalar `sourceElementIndex`. Raw identities include unknown/inactive XML, and equal labels or values never supply identity. Private baseline values, namespaces and lexical guards stay out of public reports/Core.

A bound record is conditional source eligibility. Whole-plan analysis, closure, budgets and actual candidate validation still decide readiness. General chart mapped-field findings do not grant arbitrary chart edit permission; `describeImportCapabilities` remains an observation view, not a save plan.

| Source condition | Numeric write policy |
| --- | --- |
| Complete dense cache and explicit local A1 ranges for series names, categories and values, with exactly agreeing stored XLSX values | Bind existing scalar values |
| Missing cache or workbook, literal-only data, unresolved/ambiguous cells or cache/workbook disagreement | Unmapped; no repair during save. Sparse caches remain read-only in v1; v2 admits existing nonblank scatter points only when the complete blank-aware cache and workbook agree. |
| Formula cells, named calculations, external dependencies, auto-update, signatures/macros or unhandled chart formulas | Block; no evaluation, fetch or dependency rewrite |
| Shared chart, repeated frame use of one relationship, shared workbook or value cell used by another reference including labels | Block pending explicit shared-edit/separation policy |
| Chart MCE projection or opaque/mixed scalar content | Unmapped; retained source cannot be safely patched |
| Unknown adjacent extensions and worksheet numeric formatting | Retain exact lexical surroundings; no general semantic interpretation |

Numeric agreement requires both `Object.is` equality, which distinguishes signed zero, and equality of the source decimal values after accounting for equivalent decimal/exponent spellings. Distinct decimal values that happen to round to the same JavaScript number do not bind. New values serialize with `String(value)` except negative zero uses `"-0"`; original numeric spelling is retained for unmodified atoms. This profile does not perform full XLSX schema validation or prove the meaning of arbitrary opaque extensions.

Nonzero source numbers that underflow to JavaScript zero are not treated as zero. An underflowing chart-cache or axis-bound value leaves the chart source-only; an underflowing workbook value prevents a numeric binding even if the chart cache contains zero. Genuine positive and negative zero and representable subnormal values remain eligible. This avoids silently equating a nonzero source value with zero; it does not promise arbitrary-precision numeric import.

## Candidate transaction

Parser alone owns source binding and writes. It stages lexical chart-cache replacements and inner worksheet replacements, rebuilds only the changed embedded XLSX payloads, then patches the outer package. Each scalar must match its original raw index, namespace and lexical text. Both layers are prepared before a candidate can be published; errors, cancellation and stale snapshots publish no partial bytes and never advance the source baseline.

Independent inner package comparisons allow only approved worksheet payloads. Outer comparison allows only approved text/chart/workbook payloads, retaining relationships, content types and all other payloads. The actual candidate is reprojected and rebound: the proposed number must agree with both cache and workbook at the original identities, followed by full normalized Core equality. A cache-first snapshot alone cannot hide a stale workbook. ZIP envelopes of replaced XLSX packages can change; this is scoped preservation, not whole-archive byte identity after edits.

`PptxSaveChange.chartValue` records `before`, `after`, `seriesIndex`, `pointIndex` on the existing `kind: "other"` change. `maxSaveChanges` counts numeric edits, rather than the two physical outer payloads each edit can affect.

`maxChartBindingWork` defaults to 1,000,000 and bounds opt-in raw-reference scanning and binding work. Existing report/point/workbook expansion, snapshot/analysis/change, output/XML and validation budgets apply. Nested comparison bytes/work share the outer validation budgets; exhausted budgets produce `save-limit-exceeded`. Staged-output accounting conservatively charges changed worksheets, embedded XLSX and other staged payloads. Capture occurs before asynchronous import/file work. Dispose, abort, undo, retry and stale-plan rules match the existing session contract; successful save does not rebase the session.

## Checked example

<!-- doc-test: docs/examples/chart-value-save.ts -->
```ts
import { importPptx } from "@pptkit/pptx-parser";

export async function saveExistingChartValue(
  source: Uint8Array, seriesIndex: number, pointIndex: number, value: number,
) {
  const session = await importPptx(source, {
    projection: "content",
    chartEditSaveProfile: "existing-values-v1",
  });
  try {
    const slide = session.document.slides.find(s => s.elements.some(e => e.type === "chart"));
    const chart = slide?.elements.find(e => e.type === "chart");
    if (!slide || !chart || chart.type !== "chart" || chart.chartType === "scatter" || chart.chartType === "bubble") throw new Error("No mapped chart.");
    const binding = session.report.chartValues?.find(b => b.elementId === chart.id &&
      b.slideId === slide.id && b.seriesIndex === seriesIndex && b.pointIndex === pointIndex);
    if (binding?.status !== "bound") throw new Error(binding?.reason ?? "No numeric source binding.");
    const updated = structuredClone(chart), series = updated.series[seriesIndex];
    if (!Number.isFinite(value) || !series || pointIndex < 0 || pointIndex >= series.values.length) {
      throw new Error("Expected a finite value at an existing point.");
    }
    series.values[pointIndex] = value;
    const index = slide.elements.indexOf(chart);
    slide.removeElement(chart.id);
    slide.insertElement(index, updated);
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") throw new Error(prepared.report.diagnostics.map(d => d.message).join("\n"));
    return await session.save(prepared.plan);
  } finally { session.dispose(); }
}
```

## Evidence and next gates

The experimental `existing-series-order-v6` profile includes the V5 scalar edits and exposes `report.chartSeriesOrders`. For a `bound` chart, each entry's `series` maps the current Core position to the raw `c:idx` source ID. **Series-order save requests are currently blocked:** PowerPoint displayed the original series order after an output changed only the `c:order` attributes. A bound report is source eligibility, not permission to save. The existing scalar edits remain available without an order request.

When source series element order differs from `c:order`, content import reports `chart.series-order`: Core follows `c:order`, while a controlled PowerPoint for Mac column chart displayed the XML element order. Applications may differ; inspect the source and the target application before relying on a preview's series order.

The internal candidate checks package-wide XML extension lists and presentation timing, signature/VBA guards, chart relationships, closure and reimport, but those checks cannot establish native order behavior or prove the absence of arbitrary non-relationship references or opaque binary dependencies. See [order research](../testing/chart-series-order-research.md) and [session tests](../../packages/pptx-parser/test/chart-order-session.test.mjs).

- [Parser tests](../../packages/pptx-parser/test/chart-save.test.mjs) cover both dialects and three families over three rounds, exact unchanged outer/inner payloads, numeric lexical normalization/negative zero, unknown neighbors/cell formatting, missing/conflicting sources, formula/external/auto-update rejection, shared charts/workbooks and label aliases, distinct chart identities, mixed text edits, unsupported topology/style/geometry, limits/undo/staleness and second-layer failure/cancellation/retry. A deliberately wrong equal-valued worksheet target is rejected even when cache/Core alone agree.
- A mixed transaction test combines numeric cache/workbook edits with metadata subject and keyword removal/reinsertion in both dialects for three rounds. It checks exact cache and worksheet changes, unrelated inner and outer payloads, retained metadata siblings and rejection of an unsupported geometry mutation. This is automated source-save evidence, not native Edit Data evidence.
- [Runtime parity](../../scripts/chart-value-save-parity.test.mjs) checks Node/ESM/global bytes and reports over three rounds; browser DOM independently reads chart cache and worksheet cells, with network blocked. Exporter-produced multi-series bar/line/pie sources also bind and save over three rounds; Node file capture and installed Node/declaration consumers exercise public entry points.

- [Independent/native source-save evidence](../testing/chart-value-save.md): separate Python cache/XLSX/payload checks, three-round controlled bar/line/pie saves, untouched chart/series, and named LibreOffice headless resaves/PDF reopen. The native resaves drop referenced XLSX sources and correctly remain unmapped for numeric writes. Optional adversarial oracle tests reject stale/mismatched values and unapproved mutations.

Next: full dependency/reference closure, broader native-produced workbooks, explicit shared-edit policies, structural series/category changes and named PowerPoint/LibreOffice actual Edit Data with save/reopen. No fresh export, parser agreement or headless resave substitutes for interactive evidence. IMP-09.2.6 and its inherited gates remain unchecked.
