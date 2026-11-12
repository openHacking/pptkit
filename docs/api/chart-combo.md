# Column and line combination charts

Use `chartType: "combo"` when clustered columns and lines share categories but measure different quantities:

```ts
deck.slides[0]!.addElement({
  type: "chart",
  chartType: "combo",
  box: { x: 45, y: 30, width: 600, height: 345 },
  categories: ["Jan", "Feb", "Mar"],
  series: [
    { kind: "bar", name: "Volume", values: [2, 5, 8] },
    { kind: "line", name: "Rate", values: [800, 200, 600] },
  ],
  axes: {
    value: { scale: { min: 0, max: 10, majorUnit: 2 } },
    secondaryValue: { scale: { min: 0, max: 1000, majorUnit: 200 } },
  },
});
```

Place one or more bar series before one or more line series. The left axis uses only the column values, and the right axis uses only the line values. Omit either scale for independent automatic ranges. Line points sit at category centers, even when multiple columns are clustered at a center. Both families use the chart's `gap` or `zero` blank policy.

SVG shows the two value axes and clips both series to the shared plot. PPTX uses one editable native chart and one workbook with shared categories, independent series columns and two reciprocal axis pairs. The secondary category axis is hidden. The [LibreOffice run](../testing/combination-chart-probe.md) verifies the two visible value axes and shared category centers after saving and reopening. It also records that LibreOffice's resave removes the embedded workbook.

Content import projects clustered columns on the primary axis and straight lines on the secondary axis when both families use the same ordered categories and two distinct reciprocal axis pairs with shared category centers. It also reads complete caches when the workbook is absent, as in the tested LibreOffice resave. Three fresh export/import rounds retained both scales, series and the line marker. No-op saves preserve the original bytes. The opt-in [existing-values-v3 profile](chart-value-save.md) and [CLI manifest v4](cli-save.md) can update existing finite values in either family when the cache and the embedded workbook agree. Source saves retain the axis graph and scales. Cache-only files remain read-only under this profile; fresh export can create a new editable workbook from the projected chart.

A [LibreOffice native save/reopen](../testing/combination-source-save-native-evidence.json) verified that both edited series plotted at the expected heights. A [PowerPoint 16.104 save/reopen](../testing/combination-powerpoint-resave-evidence.json) retained both scales, aligned categories, edited values and the workbook. PowerPoint moves the hidden secondary category axis to the bottom; content import accepts that position when it stays hidden and correctly paired. PowerPoint interactive data editing remains unverified. Other family combinations, reversed axis assignments and unsupported axis arrangements remain source-only on import.
