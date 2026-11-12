/** Public-authoring fixture shared by runtime, preservation and native checks. */
export function createTextOwnerFixture(createPresentation, imageBytes) {
  const document = createPresentation({ theme: { fonts: { heading: "Liberation Sans", body: "Liberation Sans" } } });
  const first = document.addSlide(), second = document.addSlide();
  first.addElement({
    type: "shape", shape: "roundRect", name: "Text owner",
    box: { x: 30, y: 30, width: 500, height: 80 },
    style: { fill: { type: "solid", color: "E0E7FF" }, stroke: { width: 1, paint: { type: "solid", color: "123456" } } },
    text: { content: [{ runs: [{ text: "Original shape text", style: { fontFamily: "Liberation Sans", fontSize: 24, color: "123456" }, action: { type: "slide", slideId: second.id } }] }] },
  });
  first.addElement({
    type: "group", box: { x: 550, y: 30, width: 180, height: 100 }, coordinateSize: { width: 90, height: 50 }, transform: { rotation: 15, flipH: true },
    children: [{ type: "shape", shape: "ellipse", box: { x: 5, y: 5, width: 80, height: 40 }, style: { fill: { type: "solid", color: "FFF1CC" } }, text: { content: "Nested shape control" } }],
  });
  first.addElement({ type: "chart", chartType: "bar", categories: ["A", "B"], series: [{ name: "Unchanged data", values: [3, 5] }], box: { x: 30, y: 150, width: 300, height: 180 } });
  first.addElement({ type: "table", columns: [130, 130], rows: [{ height: 40, cells: [{ content: "Cell control A" }, { content: "Cell control B" }] }], box: { x: 370, y: 180, width: 260, height: 40 } });
  first.setNotes([
    { runs: [{ text: "Original speaker notes", style: { fontFamily: "Liberation Sans", fontSize: 16, color: "654321" }, action: { type: "url", url: "https://example.invalid/notes?x=1&y=2", tooltip: "Notes link" } }] },
    { runs: [{ text: "Notes control é🙂", style: { bold: true } }] },
  ]);
  second.addElement({ type: "text", box: { x: 30, y: 30, width: 600, height: 80 }, content: "Second slide control" });
  second.setNotes([{ runs: [{ text: "Second notes control", style: { italic: true } }] }]);
  if (imageBytes) {
    document.registerAsset({ id: "shared", kind: "image", source: { type: "bytes", value: imageBytes } });
    for (const slide of [first, second]) slide.addElement({ type: "image", assetId: "shared", box: { x: 650, y: 350, width: 100, height: 60 } });
  }
  return document;
}
