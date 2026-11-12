import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg, renderLayoutToSvg } from "../packages/svg-renderer/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { pixelPng } from "../tests/fixtures/presentations.mjs";

function fixture(positioned = false) {
  const document = createPresentation({ size: { width: 600, height: 250 }, metadata: { title: "Composition parity" } });
  const image = document.registerAsset({ id: "asset", kind: "image", source: { type: "url", value: pixelPng }, mimeType: "image/png", width: 8, height: 4 });
  const sizes = [[40, 20], [60, 40], [80, 50], [100, 70], [50, 60]];
  const xs = [20, 80, 160, 260, 380];
  const frame = (index) => ({ x: positioned ? xs[index] : 1, y: positioned ? (250 - sizes[index][1]) / 2 : 2, width: sizes[index][0], height: sizes[index][1] });
  const slide = document.addSlide({ id: "slide" });
  slide.addElement({ type: "shape", id: "a", shape: "rect", box: frame(0) });
  slide.addElement({ type: "image", id: "b", assetId: image.id, fit: "contain", box: frame(1) });
  slide.addElement({ type: "table", id: "c", box: frame(2), columns: [40, 40], rows: [{ height: 50, cells: [{ content: "A" }, { content: "B" }] }] });
  slide.addElement({ type: "chart", id: "d", box: frame(3), chartType: "bar", categories: ["One"], series: [{ name: "Series", values: [1] }] });
  slide.addElement({ type: "group", id: "e", box: frame(4), coordinateSize: { width: 100, height: 120 }, children: [{ type: "shape", id: "child", shape: "ellipse", box: { x: 10, y: 10, width: 80, height: 100 } }] });
  slide.addElement({ type: "connector", id: "link", start: { elementId: "a", anchor: "right" }, end: { elementId: "e", anchor: "left" } });
  return document;
}
const rule = { type: "stack", container: { kind: "slide", id: "slide" }, elementIds: ["a", "b", "c", "d", "e"], direction: "horizontal", gap: 20, padding: 20, align: "center" };

test('long connector survives the authoring-to-SVG and PPTX pipeline', async () => {
  const document = createPresentation({ size: { width: 500, height: 300 } });
  const route = Array.from({ length: 150000 }, (_, index) => ({ x: index % 101, y: index % 51 }));
  document.addSlide({ elements: [{ id: 'long-path', type: 'connector', start: { x: 0, y: 0 }, end: { x: 100, y: 50 }, route }] });
  const [pptx, svg] = await Promise.all([generatePptx(document), renderPresentationToSvg(document)]);
  const slide = readGeneratedZip(pptx.bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(slide, /<a:moveTo><a:pt x="0" y="0"\/><\/a:moveTo>/);
  let index = 0;
  for (const match of slide.matchAll(/<a:lnTo><a:pt x="(-?\d+)" y="(-?\d+)"\/><\/a:lnTo>/g)) {
    const expected = index < route.length ? route[index] : { x: 100, y: 50 };
    assert.equal(Number(match[1]), expected.x * 12700); assert.equal(Number(match[2]), expected.y * 12700);
    index++;
  }
  assert.equal(index, route.length + 1);
  const points = /<polyline points="([^"]+)"/.exec(svg.slides[0].svg)?.[1];
  assert.ok(points);
  const parsed = points.split(' '); assert.equal(parsed.length, route.length + 2);
  assert.equal(parsed[0], '0,0'); assert.equal(parsed.at(-1), '100,50');
  for (let i = 0; i < route.length; i++) assert.equal(parsed[i + 1], `${route[i].x},${route[i].y}`);
});

test("composition produces the same native package and SVG as independently positioned authoring", async () => {
  const document = fixture();
  const before = normalizePresentation(document);
  const options = { composition: [rule] };
  const resolved = resolveLayout(document, options);
  const expected = resolveLayout(fixture(true));
  const { analysis, ...geometry } = resolved;
  assert.deepEqual(analysis.diagnostics, []);
  assert.deepEqual(geometry, expected);
  assert.equal(resolved.slides[0].elements[1].box.y, 110, "image contain runs after allocation");
  assert.deepEqual(resolved.slides[0].elements[5].end, { x: 380, y: 125 }, "connector targets arranged group");
  const [actualPptx, expectedPptx, actualSvg, expectedSvg] = await Promise.all([
    generatePptx(document, { layout: options }), generatePptx(fixture(true)),
    renderPresentationToSvg(document, { layout: options }), renderPresentationToSvg(fixture(true)),
  ]);
  assert.deepEqual(actualPptx.bytes, expectedPptx.bytes);
  assert.deepEqual(actualPptx.warnings, []);
  assert.deepEqual(actualSvg, expectedSvg);
  assert.deepEqual(await renderLayoutToSvg(resolved), actualSvg);
  const entries = readGeneratedZip(actualPptx.bytes);
  assert.ok(entries.has("ppt/charts/chart1.xml"));
  assert.ok(entries.get("ppt/slides/slide1.xml").toString().includes('<a:off x="254000" y="1460500"/>'));
  assert.deepEqual(normalizePresentation(document), before);
});

test("failed composition preserves valid output and forwards actionable warnings to both formats", async () => {
  const document = fixture();
  const layout = { composition: [{ ...rule, area: { x: 0, y: 0, width: 1, height: 1 } }] };
  const plain = await generatePptx(document);
  const result = await generatePptx(document, { layout });
  assert.deepEqual(result.bytes, plain.bytes);
  assert.equal(result.warnings[0].code, "composition-space-unavailable");
  assert.equal(result.warnings[0].path, "options.composition.0");
  const svg = await renderPresentationToSvg(document, { layout });
  assert.deepEqual(svg.slides, (await renderPresentationToSvg(document)).slides);
  assert.deepEqual(svg.warnings, result.warnings);
});


test("mixed grid tracks produce identical export/preview to independent placements", async () => {
  function gridDeck(positioned) {
    const document = createPresentation({ size: { width: 300, height: 200 } });
    const slide = document.addSlide({ id: "s" });
    for (const [id, x, y, width, height] of [["a", 5, 30, 40, 20], ["b", 60, 0, 60, 30], ["c", 220, 160, 80, 40]]) slide.addElement({ type: "shape", id, shape: "rect", box: { x: positioned ? x : 1, y: positioned ? y : 2, width, height } });
    slide.addElement({ type: "connector", id: "edge", start: { elementId: "a", anchor: "right" }, end: { elementId: "c", anchor: "top" } });
    return document;
  }
  const layout = { composition: [{ type: "grid", container: { kind: "slide", id: "s" }, columns: [{ type: "fixed", size: 50 }, { type: "content" }, { type: "fraction" }], rows: [{ type: "fixed", size: 50 }, { type: "fraction" }], columnGap: 10, rowGap: 10, placements: [{ elementId: "a", row: 0, column: 0, horizontal: "center", vertical: "end" }, { elementId: "b", row: 0, column: 1 }, { elementId: "c", row: 1, column: 2, horizontal: "end", vertical: "end" }] }] };
  const [actual, expected, svg, expectedSvg] = await Promise.all([generatePptx(gridDeck(false), { layout }), generatePptx(gridDeck(true)), renderPresentationToSvg(gridDeck(false), { layout }), renderPresentationToSvg(gridDeck(true))]);
  assert.deepEqual(actual.warnings, []);
  assert.deepEqual(actual.bytes, expected.bytes);
  assert.deepEqual(svg, expectedSvg);
});
