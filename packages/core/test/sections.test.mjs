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
