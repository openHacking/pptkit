# Scatter charts

PPTKit can author an XY scatter chart using `chartType: "scatter"`. Each series owns its X and Y arrays, so series may use different coordinates and point counts. Use `null` for an explicit blank coordinate; array holes, `undefined`, unequal X/Y lengths and nonfinite values are invalid. An empty pair of arrays is accepted by the model, but its native application behavior is not yet verified.

```ts
slide.addElement({
  id: "scatter", type: "chart", chartType: "scatter",
  box: { x: 40, y: 60, width: 600, height: 320 },
  blankDisplay: "gap",
  axes: {
    x: { scale: { min: -10, max: 10, majorUnit: 5 } },
    y: { scale: { min: -10, max: 10, majorUnit: 5 } },
  },
  series: [
    { name: "A", xValues: [-8, null, 2], yValues: [4, 3, -4], color: "#FF0000", marker: { shape: "circle", size: 10 } },
    { name: "B", xValues: [7, -1], yValues: [-6, 4], color: "#0000FF", line: { paint: { type: "solid", color: "#0000FF" }, width: 2 } },
  ],
});
```

Both axes use numeric value scales. Omitted scales are computed from the displayed points; fixed scales use the same finite minimum, maximum and positive major-unit contract as other value axes. The default series has markers and no line. A supplied `line` enables straight segments; `false` disables a line or marker. Point order is the input order rather than sorted X order. `gap` omits an incomplete point and breaks the line, `zero` displays each blank coordinate at zero, and `span` omits incomplete points while connecting the surrounding points. The default is `gap` when either coordinate contains a blank, otherwise `zero`. Null stays distinct from zero in both the chart cache and embedded workbook.

A fixed axis that places a finite point so far outside its range that its plot coordinate cannot be represented raises `RangeError` during layout rather than producing nonfinite SVG. The workbook supports at most 8,192 series and 1,048,575 points per series; export rejects larger data. Fresh export writes two native value axes and an embedded workbook with an independent X/Y column pair per series. Content import projects a bounded subset with explicit series names, independent equal-length X/Y coordinates, two position-identified value axes and straight-line/marker styles. Explicit gap, zero and span policies permit sparse caches; missing positions count against the aggregate import budget. Cache snapshots remain authoritative over conflicting workbook data. When an X or Y cache is absent, a workbook range with explicit blanks can project only if its length matches the independently resolved counterpart and an explicit supported blank policy is present; the missing cache remains ineligible for source-bound edits. Legacy signed 32-bit axis IDs are compared as signed values and retained verbatim in source saves; signed and unsigned spellings are not treated as interchangeable. Empty series, smooth styles and other unsupported source forms remain source-only. Three fresh import/export rounds and original-byte no-op preservation are tested. With `existing-values-v1`, dense source-bound X/Y coordinates can be edited independently when their caches and workbook cells satisfy the [numeric preservation guards](chart-value-save.md); the opt-in `existing-values-v2` profile also permits existing nonblank points in matching sparse sources while retaining blanks; conflicting or shared sources remain read-only. One dense two-series sample with three coordinate edits was opened, saved and reopened in PowerPoint; its cache, workbook cells and retained UI observations are checked by the [native evidence](../testing/chart-scatter-coordinate-save-powerpoint-evidence.json). Four sparse fixed-axis samples with eight coordinate edits also have [one PowerPoint save/reopen observation](../testing/chart-scatter-sparse-save-powerpoint-evidence.json), checking caches, workbook cells and retained UI coordinates; visual geometry and interactive workbook editing were not measured. The [CLI v2 chart manifest](cli-save.md) identifies X and Y separately, and v3 opts into sparse-coordinate saves; the existing v1 manifest remains unchanged. The general chart import and preservation rules still apply to neighboring supported objects.

Four fixed-axis PPTKit-authored examples opened in PowerPoint 16.104 and were saved and reopened once. Their first-print marker centers differed from PPTKit layout by at most 2.791 slide points and line endpoints by at most 2.779 points. This is an approximation, not pixel parity; automatic axes, simultaneous X/Y blanks, empty native series, native interactive workbook editing, other producers and a reopened PDF are not established by that observation. See the [authored evidence](../testing/chart-scatter-authored-powerpoint-evidence.json) and [native research](../testing/chart-scatter-research.md).

One attempted interactive test found the Edit Data in Excel control disabled on both a PPTKit sample and an independently generated control in the same PowerPoint installation. Excel was not found by the local application and Spotlight checks; the reason for the disabled controls remains unproven. No interactive workbook edit was performed. See the [availability observation](../testing/chart-scatter-edit-data-availability.json).

For points with a third quantity encoded as area, see [bubble charts](chart-bubble.md).
