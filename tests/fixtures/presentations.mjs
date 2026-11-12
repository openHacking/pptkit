// Repository-authored deterministic fixtures, distributed under the root MIT license.
// Inject the public factory so package, browser, and consumer tests share the data.
export const fixtureFontFamily = "PPTKit Fixture";
export const pixelPng = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC";

const box = (x, y, width, height) => ({ x, y, width, height });

export const fixtureCatalog = [
  { id: "text", capabilities: ["rich-text", "lists", "actions", "notes"], snapshot: "corpus-text.png", warnings: { export: [], preview: ["slide-action-not-linked"] } },
  { id: "theme-layout", capabilities: ["theme", "layout", "placeholder", "hidden-slide", "metadata"], snapshot: "corpus-theme-layout.png", warnings: { export: [], preview: [] } },
  { id: "geometry", capabilities: ["shapes", "connector", "nested-group", "transforms"], snapshot: "corpus-geometry.png", warnings: { export: [], preview: [] } },
  { id: "images", capabilities: ["image", "contain", "cover", "crop", "accessibility"], snapshot: "corpus-images.png", warnings: { export: [], preview: [] } },
  { id: "table", capabilities: ["table", "column-span", "row-span"], snapshot: "corpus-table.png", warnings: { export: [], preview: [] } },
  { id: "charts", capabilities: ["bar", "line", "pie", "workbook", "image"], snapshot: "corpus-charts.png", warnings: { export: [], preview: [] } },
  { id: "missing-asset", capabilities: ["asset-failure"], snapshot: undefined, warnings: { export: ["asset-read-failed", "image-omitted"], preview: ["asset-path-unsupported"] } },
];

function addImage(presentation, slide, id, fit, imageBox, extra = {}) {
  if (!presentation.getAsset("pixel")) presentation.registerAsset({
    id: "pixel", kind: "image", source: { type: "url", value: pixelPng }, mimeType: "image/png", width: 8, height: 4,
    accessibility: { description: "Repository fixture pixel" },
  });
  slide.addElement({ type: "image", id, assetId: "pixel", fit, box: imageBox, ...extra });
}

export function createFixture(id, createPresentation) {
  if (!fixtureCatalog.some((fixture) => fixture.id === id)) throw new Error(`Unknown fixture: ${id}`);
  const presentation = createPresentation({
    id: `fixture-${id}`, metadata: { title: `PPTKit ${id}`, author: "PPTKit", language: "en-US" },
    theme: { fonts: { heading: fixtureFontFamily, body: fixtureFontFamily }, colors: { accent1: "2457D6", accent2: "D94A66" } },
  });
  if (id === "theme-layout") presentation.defineSlideLayout({
    id: "fixture-layout", name: "Fixture layout", background: { type: "solid", color: "E0E7FF" },
    elements: [{ type: "shape", id: "layout-stripe", shape: "rect", box: box(0, 0, 16, 540), style: { fill: { type: "solid", color: { theme: "accent1" } } } }],
    placeholders: [{ key: "title", kind: "title", box: box(40, 30, 800, 100), textStyle: { run: { fontFamily: { theme: "heading" }, fontSize: 32 } } }],
  });
  const slide = presentation.addSlide({ id: `${id}-slide`, ...(id === "theme-layout" ? { layoutId: "fixture-layout", section: "Opening", tags: ["fixture"], customData: { fixture: true } } : {}) });
  if (id === "text") {
    slide.addElement({ type: "text", id: "rich-text", box: box(40, 30, 800, 150), content: [
      { runs: [{ text: "Rich & <editable> ", style: { bold: true, color: { theme: "accent1" } } }, { text: "text", style: { italic: true, underline: true }, action: { type: "url", url: "https://example.com/?a=1&b=2" } }] },
      { style: { bullet: { type: "number", startAt: 3 } }, runs: [{ text: "Numbered paragraph" }] },
      { style: { bullet: { type: "bullet", character: "*" } }, runs: [{ text: "Bulleted paragraph" }] },
    ] });
    presentation.addSlide({ id: "notes-slide", notes: "Speaker & reviewer notes", hidden: true });
    slide.addElement({ type: "text", id: "slide-link", box: box(40, 220, 400, 60), content: [{ runs: [{ text: "Go to notes", action: { type: "slide", slideId: "notes-slide" } }] }] });
  } else if (id === "theme-layout") {
    slide.addElement({ type: "text", id: "bound-title", placeholderKey: "title", content: "Inherited title" });
    presentation.addSlide({ id: "override-slide", layoutId: "fixture-layout", hidden: true, background: { type: "solid", color: "FFFFFF" } });
  } else if (id === "geometry") {
    slide.addElement({ type: "shape", id: "source", shape: "roundRect", box: box(40, 40, 180, 90), text: { content: "Source" }, style: { fill: { type: "solid", color: "DBEAFE" } } });
    slide.addElement({ type: "shape", id: "target", shape: "ellipse", box: box(400, 40, 180, 90), style: { fill: { type: "solid", color: "FBCFE8" } } });
    slide.addElement({ type: "connector", id: "anchored", start: { elementId: "source", anchor: "right" }, end: { elementId: "target", anchor: "left" }, style: { endArrow: "triangle", width: 2, paint: { type: "solid", color: "2457D6" } } });
    slide.addElement({ type: "group", id: "outer", box: box(100, 240, 360, 180), coordinateSize: { width: 180, height: 90 }, transform: { rotation: 10 }, children: [
      { type: "group", id: "inner", box: box(10, 10, 140, 60), coordinateSize: { width: 280, height: 120 }, transform: { flipH: true }, children: [
        { type: "shape", id: "nested-shape", shape: "triangle", box: box(20, 10, 90, 80), style: { fill: { type: "solid", color: { theme: "accent1" } } } },
        { type: "text", id: "nested-text", content: "Nested", box: box(120, 20, 150, 60) },
      ] },
    ] });
  } else if (id === "images") {
    for (const [index, fit] of ["stretch", "contain", "cover", "crop"].entries()) {
      const x = 40 + index * 220;
      slide.addElement({ type: "shape", id: `background-${fit}`, shape: "rect", box: box(x, 80, 180, 110), style: { fill: { type: "solid", color: "2457D6" } } });
      addImage(presentation, slide, `image-${fit}`, fit, box(x, 80, 180, 110), fit === "crop" ? { crop: { left: 0.1, right: 0.2, top: 0.15, bottom: 0.05 } } : {});
      slide.addElement({ type: "text", id: `label-${fit}`, content: fit, box: box(x, 220, 180, 50) });
    }
  } else if (id === "table") {
    slide.addElement({ type: "table", id: "spans", box: box(40, 40, 720, 240), columns: [240, 240, 240], rows: [
      { height: 80, cells: [{ content: "Merged heading", colSpan: 3, style: { fill: { type: "solid", color: "DBEAFE" } } }] },
      { height: 80, cells: [{ content: "Two rows", rowSpan: 2 }, { content: "A" }, { content: "B" }] },
      { height: 80, cells: [{ content: "C" }, { content: "D" }] },
    ] });
  } else if (id === "charts") {
    // An actual asynchronously loaded asset exposes cross-export state leaks.
    addImage(presentation, slide, "chart-pixel", "stretch", box(10, 10, 8, 8));
    for (const [index, chartType] of ["bar", "line", "pie"].entries()) slide.addElement({
      type: "chart", id: `chart-${chartType}`, chartType, categories: ["A", "B", "C"],
      series: [{ name: "Series", values: [2, 4, 3] }], title: chartType, box: box(30 + index * 310, 80, 290, 300),
    });
  } else {
    presentation.registerAsset({ id: "missing", kind: "image", source: { type: "path", value: "/pptkit-fixture-does-not-exist.png" } });
    slide.addElement({ type: "image", id: "omitted", assetId: "missing", box: box(40, 40, 100, 100) });
  }
  return presentation;
}
