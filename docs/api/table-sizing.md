# Measured table row sizing

`LayoutOptions.tableSizing` opts individual tables into measured row geometry. It requires a synchronous [text measurement provider](layout.md#opt-in-measurement-and-diagnostics), including the optional [browser font-backed adapter](browser-measurement.md). No table is resized by default.

## Rule contract

```ts
interface AutoTableRow {
  row: number;
  minHeight?: number;
  maxHeight?: number;
}
interface TableSizingRule {
  elementId: string;
  autoRows: readonly AutoTableRow[];
  overflow?: "preserve" | "clip";
  rotatedTextSizing?: "conservative" | "compact";
  includeHidden?: boolean;
  tolerance?: number;
}
```

`autoRows` names zero-based rows that may shrink or grow. Every unlisted row retains its authored height. Minimum height defaults to 1 logical point; supplied minima/maxima must be finite, positive and ordered. An omitted maximum is unbounded. Duplicate table targets/row indices and malformed numeric options throw before geometry changes.

The table's explicit column widths remain fixed and must sum to its frame width within tolerance. Cell width accounts for column spans and subtracts left/right margins before measuring rich text. Paragraphs, mixed runs, fonts, lists, wrapping and line spacing use the provider contract. Measured text extent plus top/bottom margins determines the required sum of heights across a cell's row span.

For quarter-turn or vertical cell text, an unwrapped inline measurement supplies a conservative row-span height, and a wrapped measurement verifies the resolved area before commit. Quarter turns use horizontal metrics with swapped axes; other vertical modes require a dedicated vertical provider such as the built-in browser adapter. A `table-text-direction-sizing-approximate` warning records the approximation. If this extent exceeds a finite row-span maximum, Layout measures the wrapped text at that maximum; if both axes fit, it probes up to 16 intermediate heights and remeasures the selected result before applying it. `table-text-direction-wrapped-fallback` records that choice. Set `rotatedTextSizing: "compact"` to run the same bounded search even when the unwrapped extent fits or the maximum is unbounded; the default `"conservative"` avoids those extra measurements. This bounded search does not guarantee the smallest fitting height and is not native Office parity.

Single-row constraints are resolved before longer spans; equal-length spans use source row/cell order. Extra required height is shared equally among eligible automatic rows in the span, redistributing when a row reaches its maximum. Increasing a row preserves earlier lower-bound constraints. This is a deterministic balanced-growth policy, not a globally minimum-height optimization. Existing fixed rows never absorb extra height.

The default tolerance is 0.02 logical points, accounting for small browser text-rectangle rounding such as 1/64 pixel. Set `tolerance: 0` for exact comparisons. This tolerance controls this table's width/height comparisons; the separate general Layout diagnostic tolerance still has its own setting.

## Overflow and atomicity

- **`preserve` (default):** if any cell cannot fit fixed/max row heights, has unresolvable width/top overhang, or margins exceed available width, the table retains all authored row/frame geometry and reports the affected cell path.
- **`clip`:** apply bounded row heights despite reported overflow. The SVG preview clips each cell to its resolved bounds. Native PPTX retains editable text and emits those row heights; Office applications may apply their own reflow and clipping rules. This is not a guarantee of native playback/print fidelity.
- Missing/failed measurement, invalid spans, missing rows/targets, width mismatch or nonfinite geometry always retain the table, even in clip mode. Other tables may still resolve.

Codes include `table-sizing-target-missing`, `table-sizing-row-missing`, `table-geometry-mismatch`, `table-span-invalid`, `table-cell-overflow`, `table-row-overflow`, `table-sizing-invalid`, and existing provider/font failure codes. Warnings include object identity and cell paths. Hidden tables and hidden ancestors are skipped unless the rule sets `includeHidden`.

Successful sizing updates the table frame height to the sum of resolved rows. Position, width, cell content, fixed row heights and authoring state are preserved. A successful fixed-only rule (`autoRows: []`) can reconcile frame height with the existing row grid after validating content fits.

## Pipeline and resolved geometry

Sizing runs on detached normalized data before composition, connector anchors, image/chart resolution and final diagnostics. A subsequent stack/alignment rule therefore sees the measured table height. Templates and nested groups retain their normal ownership and coordinate systems. Group frames do not automatically grow to contain taller tables; boundary diagnostics can report that separately.

A successful table becomes a `LayoutTableElement` with `cellLayout`, an array of `{ rowIndex, cellIndex, columnIndex, box }`. Boxes are table-local and include merged spans. Layout's table sizing and diagnostics share one grid resolver to avoid inconsistent cell placement.

The SVG renderer uses these exact boxes for positioned XHTML cells, table/row/cell roles and per-cell clipping. This prevents HTML table minimum-content sizing from silently growing capped rows. Tables without successful explicit sizing retain the previous preview behavior. PPTX uses native table rows/cells and writes the same row dimensions/frame height. The extra Layout geometry is not serialized into Core IR.

Final analysis may remeasure text after sizing; providers must be deterministic for the same requests. Identical diagnostics from sizing and final analysis are deduplicated. The `measure` provider does not read or alter table dimensions itself.

## Checked example

The document supplied here contains a table with ID `report-table` and at least three rows. Row zero remains fixed.

<!-- doc-test: docs/examples/table-sizing.ts -->
```ts
import type { PresentationDocument } from "@pptkit/core";
import { resolveLayout, type LayoutOptions, type TextMeasurer } from "@pptkit/layout";
import { generatePptx } from "@pptkit/pptx-exporter";
import { renderLayoutToSvg } from "@pptkit/svg-renderer";

export async function fitTable(document: PresentationDocument, measure: TextMeasurer) {
  const layout: LayoutOptions = {
    textMeasurement: { measure },
    tableSizing: [{
      elementId: "report-table",
      autoRows: [{ row: 1, minHeight: 24, maxHeight: 120 }, { row: 2, minHeight: 24 }],
      overflow: "preserve",
    }],
  };
  const resolved = resolveLayout(document, layout);
  if (resolved.analysis?.diagnostics.length) return { diagnostics: resolved.analysis.diagnostics };
  return {
    preview: await renderLayoutToSvg(resolved),
    pptx: await generatePptx(document, { layout }),
  };
}
```

## Verification and limits

[Focused tests](../../packages/layout/test/table-sizing.test.mjs) cover rich text/margins, colSpan/rowSpan, overlapping span constraints, min/max growth, fixed rows, preserve/clip, unavailable providers, hidden/nested tables, atomic failure, tolerance, finite geometry and input detachment. [Pinned-font browser tests](../../packages/svg-renderer/test/measurement.spec.mjs) compare actual DOM cell rectangles to `cellLayout` and extract native row/frame dimensions from generated PPTX, for both fitted and capped/clipped rows.

Columns are not automatically sized. Pagination, repeated headers, table style/border precedence and native Office font parity remain separate tasks. Painted border effects are outside text measurement; the positioned preview draws borders inside the cell without changing its content allocation.
