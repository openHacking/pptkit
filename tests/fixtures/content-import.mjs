// Repository-authored direct-style content import fixture.
export function createContentImportFixture(createPresentation) {
  const p = createPresentation({ metadata: { title: "Content import integration" } }), s = p.addSlide();
  s.addElement({ type: "text", id: "text", accessibility: { description: "text" }, box: { x: 20, y: 30, width: 320, height: 100 }, content: [{ runs: [{ text: "Editable & exact <text>", style: { fontFamily: "Arial", fontSize: 20, color: "123456", bold: true } }], style: { bullet: { type: "number", style: "arabicPeriod", startAt: 2 }, indent: 20, hanging: 10 } }], frame: { margin: 4, autoFit: { mode: "shrink", fontScale: .8 } } });
  s.addElement({ type: "group", id: "group", accessibility: { description: "group" }, box: { x: 60, y: 160, width: 300, height: 180 }, coordinateSize: { width: 150, height: 90 }, transform: { rotation: 20, flipH: true }, children: [
    { type: "shape", id: "shape", accessibility: { description: "shape" }, shape: "ellipse", box: { x: 10, y: 15, width: 80, height: 40 }, style: { fill: { type: "solid", color: "AABBCC", opacity: .7 }, stroke: { width: 2, paint: { type: "solid", color: "112233" }, dash: "dash" } }, text: { content: [{ runs: [{ text: "Inside group", style: { color: "000000", fontFamily: "Arial", fontSize: 16 } }] }] } },
  ] });
  return p;
}

export function createImageImportFixture(createPresentation, pngBytes) {
  const p = createPresentation({ size: { width: 400, height: 240 }, metadata: { title: "Embedded image import" } });
  const asset = p.registerAsset({ kind: "image", id: "shared-raster", mimeType: "image/png", width: 8, height: 4, source: { type: "bytes", value: pngBytes } });
  const s = p.addSlide();
  s.addElement({ type: "image", id: "cropped", name: "Crop", assetId: asset.id, fit: "crop", crop: { left: .125, right: .25, top: .25, bottom: 0 }, opacity: .6, box: { x: 20, y: 20, width: 120, height: 80 }, transform: { rotation: 30, flipH: true }, accessibility: { description: "Cropped raster" } });
  s.addElement({ type: "group", id: "group", name: "Group", accessibility: { description: "Group" }, box: { x: 180, y: 60, width: 160, height: 120 }, coordinateSize: { width: 80, height: 60 }, transform: { rotation: -15, flipV: true }, children: [
    { type: "image", id: "nested", name: "Nested", assetId: asset.id, fit: "stretch", box: { x: 10, y: 5, width: 60, height: 40 }, transform: { flipH: true, flipV: true }, accessibility: { description: "Nested raster" } },
  ] });
  p.addSlide().addElement({ type: "image", id: "shared", name: "Shared", accessibility: { description: "Shared" }, assetId: asset.id, fit: "stretch", box: { x: 0, y: 0, width: 400, height: 240 } });
  return p;
}

export function createReferenceImportFixture(createPresentation) {
  const p = createPresentation({ size: { width: 480, height: 320 } });
  const s = p.addSlide({ id: "source" }), next = p.addSlide({ id: "next" });
  s.addElement({ type: "connector", id: "edge", name: "Edge", accessibility: { description: "Edge" }, start: { elementId: "left", anchor: "right" }, end: { elementId: "right", anchor: "left" }, style: { width: 2, paint: { type: "solid", color: "112233" } }, action: { type: "slide", slideId: s.id, tooltip: "Self" } });
  s.addElement({ type: "shape", id: "left", name: "Left", accessibility: { description: "Left" }, shape: "rect", box: { x: 20, y: 40, width: 40, height: 40 }, style: { fill: { type: "solid", color: "88AAFF" } }, action: { type: "url", url: "https://example.invalid/a?x=1&y=%E4%B8%AD", tooltip: "Open <&>" } });
  s.addElement({ type: "shape", id: "right", name: "Right", accessibility: { description: "Right" }, shape: "ellipse", box: { x: 200, y: 40, width: 60, height: 40 }, transform: { rotation: 90, flipH: true }, style: { fill: { type: "solid", color: "FFCC88" } } });
  s.addElement({ type: "text", id: "linked-text", name: "Linked text", accessibility: { description: "Linked text" }, box: { x: 20, y: 140, width: 200, height: 50 }, content: [
    { runs: [{ text: "Next", action: { type: "slide", slideId: next.id, tooltip: "Next slide" } }, { text: " URL", action: { type: "url", url: "mailto:user@example.invalid?subject=Hello%20world" } }] },
  ] });
  s.addElement({ type: "group", id: "group", name: "Group", accessibility: { description: "Group" }, box: { x: 260, y: 180, width: 160, height: 100 }, coordinateSize: { width: 160, height: 100 }, transform: { rotation: 15, flipV: true }, action: { type: "slide", slideId: next.id }, children: [
    { type: "connector", id: "group-edge", name: "Group edge", accessibility: { description: "Group edge" }, start: { elementId: "group-target", anchor: "top" }, end: { x: 130, y: 80 }, style: { width: 1, paint: { type: "solid", color: "223344" } } },
    { type: "shape", id: "group-target", name: "Group target", accessibility: { description: "Group target" }, shape: "roundRect", box: { x: 10, y: 10, width: 60, height: 30 }, style: { fill: { type: "solid", color: "CCCCCC" } }, action: { type: "slide", slideId: s.id } },
  ] });
  return p;
}

export function createNotesImportFixture(createPresentation) {
  const p = createPresentation({ metadata: { title: "Speaker notes import" } });
  const first = p.addSlide({ id: "notes-source" }), next = p.addSlide({ id: "notes-next" });
  p.addSlide({ id: "notes-none" });
  first.addElement({ type: "text", box: { x: 20, y: 20, width: 300, height: 60 }, content: "Visible slide content" });
  first.setNotes([
    { runs: [{ text: "Speaker α😀 & <literal>", style: { fontFamily: "Arial", fontSize: 18, bold: true, color: "123456", language: "zh-CN" } }, { text: "\nLine two" }], style: { align: "right", indent: 18, hanging: 9, bullet: { type: "number", style: "arabicPeriod", startAt: 3 }, spaceAfter: 4, lineSpacing: 1.2 } },
    { runs: [{ text: "Next", action: { type: "slide", slideId: next.id, tooltip: "Next slide" } }, { text: " Self", action: { type: "slide", slideId: first.id } }, { text: " URL", action: { type: "url", url: "https://example.invalid/notes?a=1&b=%E4%B8%AD", tooltip: "Notes URL" } }] },
    { runs: [{ text: "", style: { fontSize: 12, italic: true } }] },
  ]);
  next.setNotes([{ runs: [{ text: "", style: { fontSize: 14, bold: true, color: "654321" } }], style: { align: "center" } }]);
  return p;
}

export function createTableImportFixture(createPresentation) {
  const p = createPresentation({ size: { width: 640, height: 480 }, metadata: { title: "Merged table import" } });
  const first = p.addSlide({ id: "table-source" }), next = p.addSlide({ id: "table-next" });
  const style = {
    fill: { type: "solid", color: "DDEEFF", opacity: .75 },
    stroke: { width: 1.5, paint: { type: "solid", color: "123456" }, dash: "dot" },
    margin: { left: 3, right: 4, top: 2, bottom: 5 }, verticalAlign: "middle",
  };
  const cell = (text, extra = {}) => ({
    content: [{ runs: [{ text, style: { fontFamily: "Arial", fontSize: 14, color: "223344" } }] }],
    style, ...extra,
  });
  first.addElement({
    type: "table", id: "merged", name: "Merged rich-text table", accessibility: { description: "Merged cells with direct styles" },
    box: { x: 30, y: 40, width: 300, height: 100 }, columns: [90, 100, 110],
    transform: { rotation: 15, flipH: true }, action: { type: "slide", slideId: first.id, tooltip: "Self table" },
    rows: [
      { height: 40, cells: [
        cell("", { rowSpan: 2, colSpan: 2, content: [
          { runs: [{ text: "Merged α😀 & <literal>", style: { fontFamily: "Arial", fontSize: 18, bold: true, color: "112233" }, action: { type: "slide", slideId: first.id, tooltip: "Self cell" } }], style: { bullet: { type: "number", style: "arabicPeriod", startAt: 3 }, indent: 16, hanging: 8, spaceAfter: 2 } },
          { runs: [{ text: "Next", action: { type: "slide", slideId: next.id } }, { text: " URL", action: { type: "url", url: "https://example.invalid/table?a=1&b=%E4%B8%AD", tooltip: "Table URL" } }] },
        ] }),
        cell("Top right"),
      ] },
      { height: 30, cells: [cell("Bottom right")] },
      { height: 30, cells: [cell("A"), cell("B"), cell("", { content: [{ runs: [{ text: "", style: { fontFamily: "Arial", fontSize: 12, bold: true } }] }] })] },
    ],
  });
  first.addElement({
    type: "group", id: "table-group", name: "Table group", accessibility: { description: "Grouped table" },
    box: { x: 360, y: 180, width: 180, height: 160 }, coordinateSize: { width: 180, height: 160 }, transform: { rotation: -10, flipV: true },
    children: [{
      type: "table", id: "all-merged", name: "All merged", accessibility: { description: "All grid positions share one anchor" },
      box: { x: 10, y: 10, width: 150, height: 80 }, columns: [75, 75], transform: { rotation: -20, flipV: true },
      rows: [{ height: 40, cells: [cell("All merged", { rowSpan: 2, colSpan: 2 })] }, { height: 40, cells: [] }],
    }],
  });
  next.addElement({ type: "text", id: "target", name: "Target", accessibility: { description: "Link target" }, box: { x: 20, y: 20, width: 200, height: 60 }, content: "Next table slide" });
  return p;
}

/** Native chart data fixture shared by public Node/browser/installed consumers. */
export function createChartImportFixture(createPresentation) {
  const p = createPresentation({ size: { width: 960, height: 540 } }), slide = p.addSlide();
  const common = { type: "chart", categories: ["α & <Q1>", "Q2", "Q3"], box: { x: 20, y: 30, width: 400, height: 250 }, title: "Chart α😀", legend: { visible: true, position: "bottom" }, transform: { rotation: 15, flipH: true }, titleStyle: { fontSize: 18, color: "112233", bold: true } };
  slide.addElement({ ...common, name: "Bar", chartType: "bar", orientation: "horizontal", grouping: "clustered", seriesGap: 25, categoryGap: 120, series: [{ name: "Keep & <literal>", values: [0, -2.5, 8], color: "112233" }, { name: "DELETE-ME", values: [99, 98, 97], color: "AABBCC" }] });
  slide.addElement({ ...common, name: "Line", box: { x: 460, y: 30, width: 400, height: 250 }, chartType: "line", series: [{ name: "Line", values: [1, 5, 3], line: { width: 2, paint: { type: "solid", color: "ABCDEF" }, dash: "dot" }, marker: { shape: "diamond", size: 8, fill: { type: "solid", color: "FEDCBA" } } }] });
  slide.addElement({ ...common, name: "Pie", box: { x: 30, y: 300, width: 400, height: 220 }, chartType: "pie", transform: { rotation: -10, flipV: true }, series: [{ name: "Pie", values: [1, 0, 4], pointColors: ["112233", "445566", "AABBCC"] }] });
  return p;
}
