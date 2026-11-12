import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation } from "@pptkit/core";
import { resolveLayout } from "@pptkit/layout";
import { generatePptx } from "../dist/index.js";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
const style = { paint: { type: "solid", color: "224466" }, width: 2, endArrow: "triangle" };
const shape = (id, x, y, extra = {}) => ({ type: "shape", shape: "rect", id, name: id, box: { x, y, width: 40, height: 20 }, ...extra });
const edge = (extra = {}) => ({ type: "connector", id: "edge", name: "edge", start: { elementId: "a", anchor: "right" }, end: { elementId: "b", anchor: "left" }, style, ...extra });
function deck(a = {}, b = {}, c = {}) { const p = createPresentation(); const s = p.addSlide({ id: "slide" }); s.addElement(edge(c)); s.addElement(shape("a", 10, 20, a)); s.addElement(shape("b", 100, 20, b)); return p; }
const xml = (result, part = "ppt/slides/slide1.xml") => readGeneratedZip(result.bytes).get(part).toString();
const native = p => generatePptx(p, { connectorMode: "native" });
const namedId = (value, name) => new RegExp(`<p:cNvPr id="(\\d+)" name="${name}"`).exec(value)?.[1];

test("native straight connections use preassigned target IDs even when targets follow the connector", async () => {
  const p = deck(), before = normalizePresentation(p), result = await native(p), value = xml(result);
  assert.deepEqual(result.warnings, []); assert.match(value, /<p:cxnSp>/); assert.match(value, /prst="straightConnector1"/);
  assert.ok(value.includes(`<a:stCxn id="${namedId(value, "a")}" idx="3"/>`));
  assert.ok(value.includes(`<a:endCxn id="${namedId(value, "b")}" idx="1"/>`));
  assert.match(value, /<a:ext cx="635000" cy="0"/); assert.match(value, /<a:tailEnd type="triangle"/);
  assert.deepEqual(normalizePresentation(p), before);
});

test("rectangle, rounded rectangle, ellipse and text have explicit site tables", async () => {
  for (const target of ["rect", "roundRect", "ellipse", "text"]) {
    for (const [anchor, site] of [["top", 0], ["left", 1], ["bottom", 2], ["right", 3]]) {
      const p = createPresentation(), s = p.addSlide();
      s.addElement(target === "text" ? { type: "text", id: "a", name: "a", content: "Target", box: { x: 0, y: 0, width: 40, height: 20 } } : shape("a", 0, 0, { shape: target }));
      s.addElement(edge({ start: { elementId: "a", anchor }, end: { x: 100, y: 100 } }));
      const result = await native(p), value = xml(result); assert.deepEqual(result.warnings, []);
      assert.ok(value.includes(`<a:stCxn id="${namedId(value, "a")}" idx="${site * (target === "ellipse" ? 2 : 1)}"/>`));
      assert.ok(!value.includes("<a:endCxn"));
    }
  }
});

test("free native straight endpoints preserve every direction and zero extents", async () => {
  for (const end of [{ x: 100, y: 80 }, { x: -100, y: 80 }, { x: 100, y: -80 }, { x: -100, y: -80 }, { x: 0, y: 80 }, { x: 100, y: 0 }, { x: 0, y: 0 }]) {
    const p = createPresentation(); p.addSlide().addElement(edge({ start: { x: 0, y: 0 }, end }));
    const result = await native(p), value = xml(result); assert.deepEqual(result.warnings, []);
    assert.equal(value.includes('flipH="1"'), end.x < 0); assert.equal(value.includes('flipV="1"'), end.y < 0);
    assert.ok(value.includes(`<a:ext cx="${Math.abs(end.x) * 12700}" cy="${Math.abs(end.y) * 12700}"/>`));
    assert.ok(!value.includes("<a:stCxn")); assert.ok(!value.includes("<a:endCxn"));
  }
});

test("native horizontal-first elbows preserve all four directions and authored routes", async () => {
  for (const end of [{ x: 100, y: 80 }, { x: -100, y: 80 }, { x: 100, y: -80 }, { x: -100, y: -80 }]) {
    const p = createPresentation(), route = [{ x: end.x, y: 0 }];
    p.addSlide().addElement(edge({ start: { x: 0, y: 0 }, end, route }));
    const before = normalizePresentation(p), resolved = resolveLayout(p).slides[0].elements[0], result = await native(p), value = xml(result);
    assert.deepEqual(result.warnings, []); assert.deepEqual(normalizePresentation(p), before); assert.deepEqual(resolved.route, route);
    const connector = /<p:cxnSp>.*?<\/p:cxnSp>/.exec(value)[0];
    assert.match(connector, /<a:prstGeom prst="bentConnector2"><a:avLst\/><\/a:prstGeom>/);
    assert.equal(connector.includes('flipH="1"'), end.x < 0); assert.equal(connector.includes('flipV="1"'), end.y < 0);
    assert.ok(connector.includes(`<a:off x="${Math.min(0, end.x) * 12700}" y="${Math.min(0, end.y) * 12700}"/>`));
    assert.ok(connector.includes(`<a:ext cx="${Math.abs(end.x) * 12700}" cy="${Math.abs(end.y) * 12700}"/>`));
    assert.match(connector, /<a:tailEnd type="triangle"/);
    const legacy = await generatePptx(p); assert.match(xml(legacy), /<a:custGeom>/); assert.ok(!xml(legacy).includes('<p:cxnSp>'));
  }
});

test("native elbows retain target bindings and fall back for unrepresentable routes", async () => {
  const p = deck({}, { box: { x: 100, y: 80, width: 40, height: 20 } }, { route: [{ x: 100, y: 30 }] });
  const result = await native(p), value = xml(result); assert.deepEqual(result.warnings, []);
  assert.match(value, /prst="bentConnector2"/);
  assert.ok(value.includes(`<a:stCxn id="${namedId(value, "a")}" idx="3"/>`)); assert.ok(value.includes(`<a:endCxn id="${namedId(value, "b")}" idx="1"/>`));
  for (const route of [[{ x: 0, y: 80 }], [{ x: 100, y: 1 }], [{ x: 100, y: 0 }, { x: 100, y: 40 }], [{ x: 0, y: 0 }]]) {
    const candidate = createPresentation(); candidate.addSlide({ id: 'slide' }).addElement(edge({ start: { x: 0, y: 0 }, end: { x: 100, y: 80 }, route }));
    const fallback = await native(candidate); assert.equal(fallback.warnings[0].code, 'native-connector-fallback'); assert.match(xml(fallback), /<a:custGeom>/);
  }
});

test('native double elbows retain bounded adjustments and reject quantization loss', async () => {
  for (const end of [{ x: 100, y: 80 }, { x: -100, y: 80 }, { x: 100, y: -80 }, { x: -100, y: -80 }]) {
    for (const fraction of [0.25, 0.5, 0.75]) {
      const p = createPresentation(), middle = end.x * fraction;
      p.addSlide().addElement(edge({ start: { x: 0, y: 0 }, end, route: [{ x: middle, y: 0 }, { x: middle, y: end.y }] }));
      const result = await native(p), value = xml(result); assert.deepEqual(result.warnings, []);
      assert.match(value, new RegExp(`<a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val ${fraction * 100000}"/>`));
      assert.equal(value.includes('flipH="1"'), end.x < 0); assert.equal(value.includes('flipV="1"'), end.y < 0);
    }
  }
  for (const route of [[{ x: 100 / 3, y: 0 }, { x: 100 / 3, y: 80 }], [{ x: -10, y: 0 }, { x: -10, y: 80 }], [{ x: 0, y: 0 }, { x: 0, y: 80 }], [{ x: 30, y: 0 }, { x: 31, y: 80 }]]) {
    const p = createPresentation(); p.addSlide().addElement(edge({ start: { x: 0, y: 0 }, end: { x: 100, y: 80 }, route }));
    const result = await native(p); assert.ok(result.warnings.some(warning => warning.code === 'native-connector-fallback')); assert.match(xml(result), /<a:custGeom>/);
  }
});

test("nested same-group attachments retain scope and IDs under group transforms", async () => {
  const p = createPresentation();
  p.addSlide().addElement({ type: "group", id: "group", box: { x: 30, y: 40, width: 200, height: 100 }, coordinateSize: { width: 200, height: 100 }, transform: { rotation: 20, flipH: true }, children: [edge(), shape("a", 10, 20), shape("b", 100, 20)] });
  const result = await native(p), value = xml(result); assert.deepEqual(result.warnings, []);
  assert.ok(value.includes(`<a:stCxn id="${namedId(value, "a")}" idx="3"/>`)); assert.equal((value.match(/<p:cxnSp>/g) ?? []).length, 1);
  const ids = [...value.matchAll(/<p:cNvPr id="(\d+)"/g)].map(m => m[1]); assert.equal(new Set(ids).size, ids.length);
});

test("unsupported targets, transforms, center anchors and routes retain paths with located diagnostics", async () => {
  const cases = [deck({ hidden: true }), deck({ shape: "triangle" }), deck({}, {}, { start: { elementId: "a" } }), deck({}, {}, { route: [{ x: 70, y: 80 }] }), deck({}, {}, { transform: { rotation: 10 } })];
  for (const p of cases) {
    const result = await native(p), value = xml(result);
    assert.equal(result.warnings.filter(w => w.code === "native-connector-fallback").length, 1);
    assert.equal(result.warnings[0].elementId, "edge"); assert.equal(result.warnings[0].slideId, "slide");
    assert.ok(!value.includes("<p:cxnSp>")); assert.match(value, /<a:custGeom>/);
  }
});

test("layout-owned native references stay part-local while cross-part slide references fall back", async () => {
  const p = createPresentation();
  p.defineSlideLayout({ id: "layout", name: "Layout", elements: [edge(), shape("a", 10, 20), shape("b", 100, 20)] });
  p.addSlide({ id: "slide", layoutId: "layout" }).addElement(edge({ id: "inherited", start: { elementId: "a", anchor: "right" }, end: { x: 100, y: 100 } }));
  const result = await native(p);
  const layoutNumber = resolveLayout(p).layouts.findIndex(item => item.id === "layout") + 1;
  assert.match(xml(result, `ppt/slideLayouts/slideLayout${layoutNumber}.xml`), /<p:cxnSp>/);
  assert.ok(!xml(result).includes("<p:cxnSp>")); assert.equal(result.warnings[0].elementId, "inherited");
});

test("cross-group descendants are rejected rather than exported at a fabricated origin", async () => {
  for (const hidden of [false, true]) {
    const p = createPresentation(), s = p.addSlide({ id: "slide" });
    s.addElement({ type: "group", id: "group", hidden, box: { x: 0, y: 0, width: 200, height: 100 }, coordinateSize: { width: 200, height: 100 }, children: [shape("a", 10, 20)] });
    s.addElement(edge({ end: { x: 100, y: 100 } }));
    await assert.rejects(native(p), error => error.name === "PresentationValidationError" && error.diagnostics.some(item => item.code === "missing-connector-target"));
  }
});

test("native layout fallback reports the owning layout and skips hidden connectors", async () => {
  const p = createPresentation(); p.defineSlideLayout({ id: "layout", name: "Layout", elements: [edge({ start: { elementId: "a" }, end: { x: 0, y: 0 } }), shape("a", 10, 20)] }); p.addSlide({ layoutId: "layout" });
  const result = await native(p); assert.equal(result.warnings[0].layoutId, "layout"); assert.equal(result.warnings[0].elementId, "edge");
  const hidden = await native(deck({}, {}, { hidden: true })); assert.deepEqual(hidden.warnings, []); assert.ok(!xml(hidden).includes("<p:cxnSp>"));
});

test("Layout retains detached authored connection identities with resolved coordinates", () => {
  const p = deck(), a = resolveLayout(p), b = resolveLayout(p), edge = a.slides[0].elements[0];
  assert.deepEqual(edge.start, { x: 50, y: 30 }); assert.deepEqual(edge.connections, { start: { elementId: "a", anchor: "right" }, end: { elementId: "b", anchor: "left" } });
  edge.connections.start.elementId = "changed"; assert.equal(b.slides[0].elements[0].connections.start.elementId, "a"); assert.equal(p.slides[0].elements[0].start.elementId, "a");
});

test("copy then source deletion remaps native attachment IDs to the copied targets", async () => {
  const p = deck(), original = p.slides[0]; p.duplicateSlide(original.id); p.removeSlide(original.id);
  const result = await native(p), value = xml(result); assert.deepEqual(result.warnings, []);
  assert.ok(value.includes(`<a:stCxn id="${namedId(value, "a")}" idx="3"/>`));
});

test("target replacement recomputes geometry and identities without mutating previous output", async () => {
  const p = deck(), slide = p.slides[0], before = await native(p), originalXml = xml(before);
  slide.removeElement("a");
  slide.addElement(shape("a", 160, 60, { box: { x: 160, y: 60, width: 80, height: 40 } }));
  const result = await native(p), value = xml(result);
  assert.deepEqual(result.warnings, []);
  const resolved = resolveLayout(p).slides[0].elements.find(item => item.id === "edge");
  assert.deepEqual(resolved.start, { x: 240, y: 80 });
  assert.deepEqual(resolved.end, { x: 100, y: 30 });
  const connector = /<p:cxnSp>.*?<\/p:cxnSp>/.exec(value)[0];
  assert.ok(connector.includes(`<a:stCxn id="${namedId(value, "a")}" idx="3"/>`));
  assert.ok(connector.includes(`<a:endCxn id="${namedId(value, "b")}" idx="1"/>`));
  assert.match(connector, /<a:xfrm flipH="1" flipV="1">/);
  assert.match(connector, /<a:off x="1270000" y="381000"\/>/);
  assert.match(connector, /<a:ext cx="1778000" cy="635000"\/>/);
  assert.notEqual(namedId(value, "a"), namedId(originalXml, "a"));
  assert.equal(xml(before), originalXml);
});

test("legacy default path output stays identical and invalid modes reject", async () => {
  const p = deck(), a = await generatePptx(p), b = await generatePptx(p, { connectorMode: "path" });
  assert.deepEqual(a.bytes, b.bytes); assert.deepEqual(a.warnings, []); assert.ok(!xml(a).includes("<p:cxnSp>"));
  await assert.rejects(generatePptx(p, { connectorMode: "unknown" }), TypeError);
});

test("native part identity allocation remains local across concurrent exports", async () => {
  const p = deck(), expected = (await native(p)).bytes;
  for (const result of await Promise.all(Array.from({ length: 8 }, () => native(p)))) assert.deepEqual(result.bytes, expected);
});


test("native target IDs do not collide with generated layout placeholders or omitted images", async () => {
  const p = createPresentation();
  p.registerAsset({ id: "missing", kind: "image", source: { type: "path", value: "unavailable.png" } });
  p.defineSlideLayout({ id: "layout", name: "With placeholders", placeholders: [{ key: "title", kind: "title", box: { x: 0, y: 0, width: 100, height: 20 } }], elements: [
    { type: "image", id: "missing-image", name: "missing-image", assetId: "missing", box: { x: 0, y: 0, width: 10, height: 10 } }, edge(), shape("a", 10, 20), shape("b", 100, 20),
  ] }); p.addSlide({ layoutId: "layout" });
  const result = await native(p), layoutNumber = resolveLayout(p).layouts.findIndex(item => item.id === "layout") + 1, value = xml(result, `ppt/slideLayouts/slideLayout${layoutNumber}.xml`);
  const ids = [...value.matchAll(/<p:cNvPr id="(\d+)"/g)].map(match => match[1]);
  assert.equal(new Set(ids).size, ids.length); assert.ok(value.includes(`<a:stCxn id="${namedId(value, "a")}" idx="3"/>`));
  assert.equal(result.warnings.filter(w => w.code === "native-connector-fallback").length, 0);
  assert.equal(result.warnings.filter(w => w.code === "image-omitted").length, 1);
});


test("transformed targets keep local site identity and export resolved attachment coordinates", async () => {
  const cases = [
    [{ rotation: 90 }, 30, 50],
    [{ rotation: -90 }, 30, 10],
    [{ rotation: 180 }, 10, 30],
    [{ flipH: true }, 10, 30],
    [{ flipV: true }, 50, 30],
    [{ rotation: 90, flipH: true, flipV: true }, 30, 10],
    [{ rotation: 45 }, 30 + Math.sqrt(200), 30 + Math.sqrt(200)],
  ];
  for (const [transform, x, y] of cases) {
    const p = deck({ transform }), before = normalizePresentation(p);
    const resolved = resolveLayout(p).slides[0].elements[0];
    assert.ok(Math.abs(resolved.start.x - x) < 1e-10);
    assert.ok(Math.abs(resolved.start.y - y) < 1e-10);
    const result = await native(p), value = xml(result);
    assert.deepEqual(result.warnings, []);
    const connector = /<p:cxnSp>.*?<\/p:cxnSp>/.exec(value)[0];
    assert.ok(connector.includes(`<a:stCxn id="${namedId(value, "a")}" idx="3"/>`));
    assert.ok(connector.includes(`<a:off x="${Math.round(Math.min(x, 100) * 12700)}" y="${Math.round(Math.min(y, 30) * 12700)}"/>`));
    assert.ok(connector.includes(`<a:ext cx="${Math.round(Math.abs(100 - x) * 12700)}" cy="${Math.round(Math.abs(30 - y) * 12700)}"/>`));
    assert.deepEqual(normalizePresentation(p), before);
  }
});
