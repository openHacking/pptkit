# Bubble charts

Use a 2D bubble chart for paired numeric X/Y observations whose size represents a third quantity:

```ts
deck.slides[0]!.addElement({
  type: "chart",
  chartType: "bubble",
  box: { x: 40, y: 30, width: 640, height: 345 },
  title: "Market size by segment",
  series: [
    { name: "North", xValues: [1, 4], yValues: [2, 4], sizes: [1, 9] },
    { name: "South", xValues: [7], yValues: [6], sizes: [4] },
  ],
});
```

Each series requires three nonempty arrays of equal length. X and Y must be finite numbers, and sizes must be finite and strictly positive. Series lengths may differ; blank, zero-size and negative-size points are outside this profile. Series colors, titles, legends, and the two numeric axes use the existing chart style contracts.

SVG clips the circles to the plot and gives them areas proportional to the authored sizes. The preview sets its largest radius from the plot size, so absolute bubble diameters can differ from those in native applications. PPTX writes a native 2D bubble chart at the default 100% scale, with size representing area. It creates a separate X, Y and size column per series and uses matching references and caches.

The content importer projects this same 2D area profile with two numeric axes. It uses complete caches or compatible stored workbook ranges if a cache is absent. A complete cache stays authoritative on workbook disagreement. A shorter series in the named LibreOffice specimen has the same trailing blank positions in all three caches; those aligned trailing positions are omitted from Core and preserved in original-byte saves. Interior gaps, unequal padding, width scaling, non-default scaling and 3D remain source-only.

Use the [existing-values-v3 profile](chart-value-save.md) to update bound dense source values without rebuilding. If the source lost its workbook or has padding or a cache/workbook conflict, generate a fresh PPTX to edit its projected data; the worksheet is rebuilt from the projected arrays. Combination charts, secondary axes, PowerPoint chart data editing, and native visual parity remain unverified. See the [native probe and evidence](../testing/bubble-export-probe.md).
