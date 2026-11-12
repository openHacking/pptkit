import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, validatePresentation } from "../dist/index.js";
const box = { x: 10, y: 20, width: 40, height: 30 };
const shape = (id, extra = {}) => ({ type: "shape", shape: "rect", id, box: { ...box }, ...extra });
const connector = (id, start, end) => ({ type: "connector", id, start, end });
const group = (id, children, extra = {}) => ({ type: "group", id, box: { x: 0, y: 0, width: 200, height: 100 }, coordinateSize: { width: 200, height: 100 }, children, ...extra });
const link = (slideId) => ({ type: "slide", slideId, tooltip: "Navigate" });
const rich = (slideId) => [{ runs: [{ text: slideId, action: link(slideId) }, { text: "URL", action: { type: "url", url: "https://example.com/?slide=source" } }] }];
function nodes(elements) { const result = []; for (const element of elements) { result.push(element); if (element.type === "group") result.push(...nodes(element.children)); } return result; }

test("atomic drawing order preserves connectors and navigation references", () => {
  const p = createPresentation(), s = p.addSlide({ id: 'source' }); p.addSlide({ id: 'destination' });
  s.addElement(connector('edge', { elementId: 'left' }, { elementId: 'right' }));
  s.addElement(shape('left', { action: link('destination') })); s.addElement(shape('right'));
  const edge = s.elements[0], left = s.elements[1];
  s.reorderElements(['right', 'edge', 'left']);
  assert.equal(s.elements[1], edge); assert.equal(s.elements[2], left);
  assert.equal(edge.start.elementId, 'left'); assert.equal(edge.end.elementId, 'right'); assert.equal(left.action.slideId, 'destination');
  assert.deepEqual(validatePresentation(p), []);
});

test("group child reorder preserves scoped connectors and navigation actions", () => {
  const p = createPresentation(), s = p.addSlide({ id: 'source' }); p.addSlide({ id: 'destination' });
  s.addElement(group('outer', [group('inner', [connector('edge', { elementId: 'left' }, { elementId: 'right' }), shape('left', { action: link('destination') }), shape('right')])]));
  s.reorderGroupElements('inner', ['right', 'edge', 'left']);
  const [right, edge, left] = s.elements[0].children[0].children;
  assert.equal(right.id, 'right'); assert.equal(edge.start.elementId, left.id); assert.equal(edge.end.elementId, right.id); assert.equal(left.action.slideId, 'destination');
  assert.deepEqual(validatePresentation(p), []);
});

test("slide copying remaps forward/backward connector targets before validation", () => {
  const p = createPresentation(), s = p.addSlide({ id: "source" });
  s.addElement(connector("before", { elementId: "right", anchor: "left" }, { elementId: "left", anchor: "right" }));
  s.addElement(shape("left")); s.addElement(shape("right", { box: { ...box, x: 100 } }));
  s.addElement(connector("after", { elementId: "left" }, { x: 150, y: 50 }));
  const copy = p.duplicateSlide(s.id), [before, left, right, after] = copy.elements;
  assert.equal(before.start.elementId, right.id); assert.equal(before.end.elementId, left.id);
  assert.equal(before.start.anchor, "left"); assert.equal(after.start.elementId, left.id); assert.deepEqual(after.end, { x: 150, y: 50 });
  assert.deepEqual(validatePresentation(p), []); assert.equal(s.elements[0].start.elementId, "right");
  p.removeSlide(s.id); assert.deepEqual(validatePresentation(p), []);
});

test("group duplication maps every nested identity and preserves local connection topology", () => {
  const p = createPresentation(), s = p.addSlide();
  const original = s.addElement(group("outer", [connector("outer-link", { elementId: "inner" }, { elementId: "outer-shape" }), shape("outer-shape"), group("inner", [connector("inner-link", { elementId: "inner-shape" }, { x: 10, y: 20 }), shape("inner-shape")])]));
  const copy = s.duplicateElement(original.id), copied = nodes([copy]), ids = new Map(nodes([original]).map((item, index) => [item.id, copied[index].id]));
  assert.equal(copy.children[0].start.elementId, ids.get("inner")); assert.equal(copy.children[0].end.elementId, ids.get("outer-shape"));
  assert.equal(copy.children[2].children[0].start.elementId, ids.get("inner-shape"));
  assert.equal(new Set([...nodes([original]), ...copied].map(item => item.id)).size, 12);
  s.removeElement(original.id); assert.deepEqual(validatePresentation(p), []);
});

test("copying only a connector keeps references to objects outside the copied selection", () => {
  const p = createPresentation(), s = p.addSlide(); s.addElement(shape("target"));
  const original = s.addElement(connector("edge", { elementId: "target", anchor: "bottom" }, { x: 0, y: 0 }));
  const copy = s.duplicateElement(original.id);
  assert.deepEqual(copy.start, original.start); assert.notEqual(copy.start, original.start); assert.deepEqual(validatePresentation(p), []);
});

test("copying only a shape does not retarget incoming connectors on the source", () => {
  const p = createPresentation(), s = p.addSlide(); s.addElement(shape("target"));
  s.addElement(connector("incoming", { elementId: "target" }, { x: 0, y: 0 }));
  const copy = s.duplicateElement("target");
  assert.notEqual(copy.id, "target"); assert.equal(s.elements.find(e => e.id === "incoming").start.elementId, "target");
});

test("slide self links remap in element, text, shape, table, nested group and notes actions", () => {
  const p = createPresentation(), other = p.addSlide({ id: "other" });
  const s = p.addSlide({ id: "source", notes: rich("source") });
  s.addElement({ type: "text", id: "text", box, content: rich("source"), action: link("source") });
  s.addElement(shape("shape", { text: { content: rich("source") }, action: link("other") }));
  s.addElement({ type: "table", id: "table", box, columns: [40], rows: [{ cells: [{ content: rich("source") }] }] });
  s.addElement(group("group", [shape("nested", { action: link("source") })], { action: link("source") }));
  const copy = p.duplicateSlide("source");
  assert.equal(copy.elements[0].action.slideId, copy.id); assert.equal(copy.elements[0].content[0].runs[0].action.slideId, copy.id);
  assert.equal(copy.elements[1].text.content[0].runs[0].action.slideId, copy.id); assert.equal(copy.elements[1].action.slideId, other.id);
  assert.equal(copy.elements[2].rows[0].cells[0].content[0].runs[0].action.slideId, copy.id);
  assert.equal(copy.elements[3].action.slideId, copy.id); assert.equal(copy.elements[3].children[0].action.slideId, copy.id);
  assert.equal(copy.notes[0].runs[0].action.slideId, copy.id);
  assert.equal(copy.elements[0].content[0].runs[0].text, "source");
  assert.equal(copy.notes[0].runs[1].action.url, "https://example.com/?slide=source");
  assert.equal(s.notes[0].runs[0].action.slideId, s.id); assert.deepEqual(validatePresentation(p), []);
});

test("element duplication leaves its slide actions unchanged", () => {
  const p = createPresentation(), s = p.addSlide({ id: "source" });
  s.addElement({ type: "text", id: "text", box, content: rich(s.id), action: link(s.id) });
  const copy = s.duplicateElement("text");
  assert.equal(copy.action.slideId, s.id); assert.equal(copy.content[0].runs[0].action.slideId, s.id);
  assert.deepEqual(validatePresentation(p), []);
});

test("shared assets/layout references and inherited connector targets retain their identities", () => {
  const p = createPresentation();
  p.registerAsset({ id: "image", kind: "image", source: { type: "bytes", value: Uint8Array.of(1, 2, 3) } });
  const layout = p.defineSlideLayout({ id: "shared-layout", name: "Shared", elements: [shape("layout-target")], placeholders: [{ key: "title", kind: "title", box }] });
  const s = p.addSlide({ id: "source", layoutId: layout.id });
  s.addElement({ type: "image", id: "picture", assetId: "image", box });
  s.addElement({ type: "text", id: "title", placeholderKey: "title", content: "Title" });
  s.addElement(connector("edge", { elementId: "layout-target" }, { elementId: "picture" }));
  const copy = p.duplicateSlide(s.id);
  assert.equal(copy.layoutId, layout.id); assert.equal(copy.elements[0].assetId, "image"); assert.equal(copy.elements[1].placeholderKey, "title");
  assert.equal(copy.elements[2].start.elementId, "layout-target"); assert.equal(copy.elements[2].end.elementId, copy.elements[0].id);
  assert.equal(p.assets.length, 1); assert.equal(p.layouts.length, 1); assert.equal(p.layouts[0], layout);
  assert.equal(p.getAssetReferences("image").length, 2); assert.deepEqual(validatePresentation(p), []);
});

test("copy does not rewrite arbitrary source IDs in names, plain text, tags or custom data", () => {
  const p = createPresentation(), s = p.addSlide({ id: "source", section: "source", tags: ["source"], customData: { slideId: "source", elementId: "target" } });
  s.addElement(shape("target", { name: "target", text: { content: "target source" } }));
  const copy = p.duplicateSlide(s.id);
  assert.equal(copy.elements[0].name, "target"); assert.equal(copy.elements[0].text.content, "target source");
  assert.deepEqual(copy.customData, { slideId: "source", elementId: "target" }); assert.deepEqual(copy.tags, ["source"]); assert.equal(copy.section, "source");
  assert.notEqual(copy.customData, s.customData); assert.ok(Object.isFrozen(copy.customData));
});

test("copying copies repeatedly follows their new self identity without leaking original targets", () => {
  const p = createPresentation(), s = p.addSlide({ id: "source", notes: rich("source") });
  s.addElement(shape("target", { action: link(s.id) })); s.addElement(connector("edge", { elementId: "target" }, { x: 1, y: 2 }));
  let source = s;
  for (let index = 0; index < 20; index++) {
    const copy = p.duplicateSlide(source.id, 0);
    assert.equal(copy.elements[1].start.elementId, copy.elements[0].id); assert.equal(copy.elements[0].action.slideId, copy.id);
    assert.equal(copy.notes[0].runs[0].action.slideId, copy.id); source = copy;
  }
  assert.equal(new Set(nodes(p.slides.flatMap(slide => slide.elements)).map(e => e.id)).size, 42);
  assert.deepEqual(validatePresentation(p), []);
});

test("generated copy IDs avoid occupied slide/layout/nested element identities", () => {
  const p = createPresentation(); p.defineSlideLayout({ id: "layout", name: "L", elements: [shape("element-1")] });
  const s = p.addSlide({ id: "slide-1", elements: [group("element-2", [shape("element-3")])] });
  const copy = p.duplicateSlide(s.id);
  assert.equal(copy.id, "slide-2"); assert.equal(copy.elements[0].id, "element-4"); assert.equal(copy.elements[0].children[0].id, "element-5");
  assert.deepEqual(validatePresentation(p), []);
});

test("invalid copy destinations leave state and allocator sequence unchanged", () => {
  const p = createPresentation(), s = p.addSlide({ id: "source" }); s.addElement(shape("original"));
  const before = normalizePresentation(p);
  for (const index of [-1, 3, 0.5, NaN, Infinity]) {
    assert.throws(() => p.duplicateSlide(s.id, index), RangeError);
    assert.throws(() => s.duplicateElement("original", index), RangeError);
    assert.deepEqual(normalizePresentation(p), before);
  }
  assert.throws(() => p.duplicateSlide("missing"), /Unknown slide/); assert.throws(() => s.duplicateElement("missing"), /Unknown element/);
  assert.equal(s.duplicateElement("original").id, "element-1"); assert.equal(p.duplicateSlide(s.id).id, "slide-1");
});

test("failed reference remapping does not consume element identities", () => {
  const p = createPresentation(), s = p.addSlide({ id: "slide" });
  s.addElement({ id: "bad", type: "text", box: { x: 0, y: 0, width: 10, height: 10 }, content: [{ runs: null }] });
  s.addElement({ id: "good", type: "shape", shape: "rect", box: { x: 0, y: 0, width: 10, height: 10 } });
  assert.throws(() => s.duplicateElement("bad"));
  assert.deepEqual(s.elements.map(element => element.id), ["bad", "good"]);
  assert.equal(s.duplicateElement("good").id, "element-1");
});

test("failed whole-slide copying leaves both generated identity sequences unchanged", () => {
  const p = createPresentation();
  p.addSlide({ id: "bad", elements: [shape("bad-shape")], notes: [{ runs: null }] });
  p.addSlide({ id: "good", elements: [shape("good-shape")] });
  assert.throws(() => p.duplicateSlide("bad"));
  assert.deepEqual(p.slides.map(slide => slide.id), ["bad", "good"]);
  const copy = p.duplicateSlide("good");
  assert.equal(copy.id, "slide-1");
  assert.equal(copy.elements[0].id, "element-1");
});

test("slide insertion stages identities and rejects reentrant destination changes", () => {
  const p = createPresentation(), first = p.addSlide({ id: "first" });
  assert.throws(() => p.addSlide({ elements: [shape(undefined), shape("duplicate"), shape("duplicate")] }), /Duplicate element/);
  assert.deepEqual(p.slides.map(slide => slide.id), ["first"]);
  const added = p.addSlide({ elements: [shape(undefined)] });
  assert.equal(added.id, "slide-1");
  assert.equal(added.elements[0].id, "element-1");
  const input = { get elements() { first.addElement(shape("during")); return [shape("candidate")]; } };
  assert.throws(() => p.addSlide(input), /changed during append/);
  assert.deepEqual(p.slides.map(slide => slide.id), ["first", "slide-1"]);
  assert.deepEqual(first.elements.map(element => element.id), ["during"]);
  const retry = p.addSlide({ elements: [shape(undefined)] });
  assert.equal(retry.id, "slide-2");
  assert.equal(retry.elements[0].id, "element-2");
  retry.removeElement("element-2");
  assert.equal(first.addElement(shape("element-2")).id, "element-2");
  const reordered = { get elements() { p.moveSlide("first", 1); p.moveSlide("first", 0); return []; } };
  assert.throws(() => p.addSlide(reordered), /Slides changed during append/);
  assert.deepEqual(p.slides.map(slide => slide.id), ["first", "slide-1", "slide-2"]);
  assert.equal(p.addSlide().id, "slide-3");
});

test("existing invalid references remain diagnosable instead of being silently repaired", () => {
  const p = createPresentation(), s = p.addSlide(); s.addElement(connector("bad", { elementId: "missing" }, { x: 1, y: 2 }));
  const copy = p.duplicateSlide(s.id);
  assert.equal(copy.elements[0].start.elementId, "missing");
  assert.equal(validatePresentation(p).filter(d => d.code === "missing-connector-target").length, 2);
});

test("copying an empty hidden slide preserves metadata and remaps notes without inventing content", () => {
  const p = createPresentation(), s = p.addSlide({ id: "source", hidden: true, notes: rich("source"), tags: ["tag"], customData: { value: 7 } });
  const copy = p.duplicateSlide(s.id, 0);
  assert.equal(copy.elements.length, 0); assert.equal(copy.hidden, true); assert.equal(copy.notes[0].runs[0].action.slideId, copy.id);
  assert.deepEqual(copy.tags, ["tag"]); assert.deepEqual(copy.customData, { value: 7 }); assert.equal(p.slides[0], copy);
});


test("wide forward-reference graphs copy without collisions or dependence on traversal order", () => {
  const p = createPresentation(), s = p.addSlide({ id: "wide" });
  for (let i = 0; i < 300; i++) s.addElement(connector(`edge-${i}`, { elementId: `shape-${i}` }, { elementId: `shape-${(i + 1) % 300}` }));
  for (let i = 0; i < 300; i++) s.addElement(shape(`shape-${i}`));
  const copy = p.duplicateSlide(s.id);
  for (let i = 0; i < 300; i++) {
    assert.equal(copy.elements[i].start.elementId, copy.elements[300 + i].id);
    assert.equal(copy.elements[i].end.elementId, copy.elements[300 + (i + 1) % 300].id);
  }
  assert.equal(new Set(copy.elements.map(e => e.id)).size, 600);
  p.removeSlide(s.id); assert.deepEqual(validatePresentation(p), []);
});
