# Area charts

The `chartType: "area"` authoring subset supports `grouping: "standard"` (default), `"stacked"`, and `"percentStacked"`. Supply the same category labels and equal-length numeric-or-null series arrays used by other charts. Each series has an editable embedded workbook and chart cache; null remains a blank rather than a numeric zero.

```ts
slide.addElement({
  id: "area", type: "chart", chartType: "area",
  box: { x: 40, y: 60, width: 600, height: 320 },
  categories: ["Q1", "Q2", "Q3"],
  series: [
    { name: "A", values: [4, null, 3], color: "#FF0000" },
    { name: "B", values: [2, -3, 1], color: "#0000FF" },
  ],
  grouping: "stacked", blankDisplay: "zero",
});
```

Standard areas accept `gap` or `zero`: gaps split filled polygons, while zero displays a blank at the baseline without changing its stored value. Stacked variants require `zero`; omitted display defaults to zero for them, and unsupported explicit modes are rejected. Series stack in order using signed cumulative values. Percent stacking divides each signed value by the sum of absolute values at that category; an all-zero category remains zero. Percentage value-axis coordinates use fractions: a manual scale of -1 to 1 means -100% to 100%. SVG preview clips filled polygons to the plot and displays percentage ticks.

Content import maps the supported single-family 2D area subset, including bounded caches or matching workbook data under supported blank policies. Unsupported source variants remain source-only; original-byte no-op save preserves their source. This is not a guarantee of complete source style, data-label, combination-chart or secondary-axis fidelity. Dense source values that satisfy `existing-values-v1` binding rules can be edited through the chart preservation profile; three save/reimport rounds synchronize their caches and workbook cells. Null and conflicting values remain subject to the profile’s binding restrictions. The general [chart import](chart-import.md) and [save-session](import-session.md) constraints continue to apply.

The focused tests cover standard and stacked geometry, signed percentage arithmetic, blanks, workbook/cache, three fresh import/export rounds, three dense numeric preservation rounds, no-op preservation and Node/browser agreement. Four PPTKit-authored fixed-axis samples were opened in PowerPoint 16.104; their first-open and third-reopen PDF polygon vertices were compared with layout and with each other after three save/close/reopen rounds. Source and all three saved chart caches and workbooks were checked. The [recorded evidence](../testing/chart-area-authored-powerpoint-evidence.json) states the exact scope. It does not establish native interactive workbook editing, other producers, or general style parity.
