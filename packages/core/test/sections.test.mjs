import test from "node:test";
import assert from "node:assert/strict";
import { createPresentation, normalizePresentation } from "../dist/index.js";

test("section identities and membership survive slide edits independently of labels", () => {
  const deck = createPresentation();
  deck.addSlide({ id: "a", section: "legacy" }); deck.addSlide({ id: "b" }); deck.addSlide({ id: "c" });
  const input = [{ id: "one", name: "same", slideIds: ["b", "a"] }, { id: "empty", slideIds: [] }, { id: "two", name: "same", slideIds: ["c"] }];
  deck.setSections(input); input[0].slideIds.length = 0;
  assert.deepEqual(deck.sections.map(s => s.slideIds), [["a", "b"], [], ["c"]]);
  assert.ok(Object.isFrozen(deck.sections[0].slideIds));
  assert.equal(deck.slides[0].section, "legacy");
  const duplicate = deck.duplicateSlide("a", 1);
  assert.deepEqual(deck.sections[0].slideIds, ["a", duplicate.id, "b"]);
  deck.moveSlide("c", 0);
  assert.deepEqual(deck.sections[2].slideIds, ["c"]);
  deck.removeSlide("c");
  assert.deepEqual(deck.sections[2].slideIds, []);
  assert.deepEqual(normalizePresentation(deck).sections, deck.sections);
  deck.setSections(undefined); assert.equal(deck.sections, undefined);
});

test("invalid replacement is atomic and never guesses membership by name", () => {
  const deck = createPresentation(); deck.addSlide({ id: "a" }); deck.addSlide({ id: "b" });
  deck.setSections([{ id: "one", slideIds: ["a"] }, { id: "two", name: "", slideIds: [] }]);
  const before = deck.sections;
  for (const invalid of [
    [{ id: "one", slideIds: ["missing"] }],
    [{ id: "one", slideIds: ["a"] }, { id: "two", slideIds: ["a"] }],
    [{ id: "same", slideIds: [] }, { id: "same", slideIds: [] }],
  ]) {
    assert.throws(() => deck.setSections(invalid)); assert.deepEqual(deck.sections, before);
  }
  const sparse = new Array(1); sparse.extra = 1;
  assert.throws(() => deck.setSections(sparse)); assert.deepEqual(deck.sections, before);
  const reentrant = { get id() { deck.addSlide({ id: "side-effect" }); return "new"; }, slideIds: [] };
  assert.throws(() => deck.setSections([reentrant]), /changed during section capture/);
  assert.deepEqual(deck.sections, before);
  deck.addSlide({ id: "c" }); assert.deepEqual(deck.sections, before);
});

test("section capture reads bounded array entries rather than inherited iterators", () => {
  const deck = createPresentation(); deck.addSlide({ id: "a" });
  const maliciousPrototype = Object.assign(Object.create(Array.prototype), { [Symbol.iterator]: function* () { yield "not-an-entry"; } });
  const members = ["a"]; Object.setPrototypeOf(members, maliciousPrototype);
  const sections = [{ id: "wanted", slideIds: members }]; Object.setPrototypeOf(sections, maliciousPrototype);
  deck.setSections(sections);
  assert.deepEqual(deck.sections, [{ id: "wanted", slideIds: ["a"] }]);
  const before = deck.sections;
  const changing = [{ get id() { changing.push({ id: "extra", slideIds: [] }); return "new"; }, slideIds: [] }];
  assert.throws(() => deck.setSections(changing), /changed during capture/);
  assert.deepEqual(deck.sections, before);
});

test("section creation, rename, boundary moves, merge and deletion use explicit identities", () => {
  const deck = createPresentation({ id: "deck" });
  for (const id of ["a", "b", "c", "d"]) deck.addSlide({ id });
  deck.setSections([{ id: "left", name: "same", slideIds: ["a", "b"] }, { id: "right", name: "same", slideIds: ["c", "d"] }]);
  deck.setSections([{ id: "left", name: "renamed", slideIds: ["a"] }, { id: "new", slideIds: ["b"] }, { id: "right", name: "same", slideIds: ["c", "d"] }]);
  // Move c across the boundary and d to the front in one transaction.
  deck.setSectionStructure(["d", "a", "b", "c"], [{ id: "right", name: "same", slideIds: ["d"] }, { id: "left", name: "renamed", slideIds: ["a", "b", "c"] }]);
  assert.deepEqual(deck.slides.map(s => s.id), ["d", "a", "b", "c"]);
  assert.deepEqual(deck.sections.map(s => [s.id, s.slideIds]), [["right", ["d"]], ["left", ["a", "b", "c"]]]);
  // Deleting the left section unassigns its pages rather than deleting them.
  deck.setSections([{ id: "right", name: "same", slideIds: ["d"] }]);
  assert.deepEqual(deck.slides.map(s => s.id), ["d", "a", "b", "c"]);
  deck.removeSlide("d"); assert.deepEqual(deck.sections[0].slideIds, []);
  deck.applyEditOperation({ schemaVersion: "pptkit-edit-v1", documentId: "deck", changes: [{ type: "slide-order", expected: ["a", "b", "c"], value: ["c", "b", "a"] }] });
  assert.deepEqual(deck.slides.map(s => s.id), ["c", "b", "a"]);
});

test("section structure replacement rejects invalid order or membership without partial mutation", () => {
  const deck = createPresentation(); for (const id of ["a", "b", "c"]) deck.addSlide({ id });
  deck.setSections([{ id: "one", slideIds: ["a", "b"] }, { id: "two", slideIds: ["c"] }]);
  const before = () => ({ order: deck.slides.map(s => s.id), sections: deck.sections });
  const original = before();
  for (const [order, sections] of [
    [["a", "b"], []], [["a", "b", "b"], []], [["a", "b", "missing"], []],
    [["c", "b", "a"], [{ id: "one", slideIds: ["a", "a"] }]],
    [["c", "b", "a"], [{ id: "one", slideIds: ["missing"] }]],
  ]) {
    assert.throws(() => deck.setSectionStructure(order, sections)); assert.deepEqual(before(), original);
  }
  const order = ["c", "b", "a"];
  const reentrant = { get id() { deck.setSections([]); return "changed"; }, slideIds: [] };
  assert.throws(() => deck.setSectionStructure(order, [reentrant]), /changed during section structure capture/);
  assert.deepEqual(deck.slides.map(s => s.id), original.order);
  assert.deepEqual(deck.sections, []);
});

test("copies do not guess section identities from equal names", () => {
  const source = createPresentation(); source.addSlide({ id: "source", section: "label" });
  source.setSections([{ id: "source-section", name: "same", slideIds: ["source"] }]);
  const dest = createPresentation(); dest.addSlide({ id: "target" });
  dest.setSections([{ id: "target-section", name: "same", slideIds: ["target"] }]);
  const result = dest.copySlideFrom(source, "source", { layoutId: null, theme: "destination", geometry: "preserve", customData: "preserve", assets: { deduplicate: "content", external: "reject" } });
  assert.equal(result.slide.section, "label");
  assert.deepEqual(dest.sections.map(s => s.slideIds), [["target"]]);
  dest.setSections([{ id: "target-section", name: "same", slideIds: ["target", result.slide.id] }]);
  assert.deepEqual(dest.sections[0].slideIds, ["target", result.slide.id]);
});
