import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, validatePresentation, replayEventTimeline, EventTimelineError } from "../dist/index.js";
const shape = (id, text = id) => ({ id, type: "shape", shape: "rect", box: { x: 10, y: 10, width: 100, height: 50 }, text: { content: text } });
const timeline = (pageId, elementId, button = elementId) => ({ stages: [{ id: "stage", target: { pageId, elementId }, trigger: { type: "object-click", elementId: button }, timing: { kind: "span", id: "time", durationMs: 10, fill: "hold" } }] });
function fixture() {
  const document = createPresentation(), slide = document.addSlide({ id: "p", elements: [shape("a"), shape("button"), { id: "g", type: "group", coordinateSize: { width: 100, height: 100 }, box: { x: 100, y: 100, width: 100, height: 100 }, children: [shape("child")] }] });
  return { document, slide };
}

test("authored timeline snapshots bind owned root/group objects and normalize to detached versioned IR plans", () => {
  const { document, slide } = fixture(), source = timeline("p", "child", "button");
  slide.setEventTimeline(source); source.stages[0].timing.durationMs = 999;
  assert.equal(slide.eventTimeline.stages[0].timing.durationMs, 10); assert.ok(Object.isFrozen(slide.eventTimeline.stages[0].timing));
  const plan = slide.getEventTimelinePlan(); assert.deepEqual(plan.elementIds, ["a", "button", "child", "g"]);
  const normalized = normalizePresentation(document); assert.equal(normalized.irVersion, 2); assert.deepEqual(normalized.slides[0].eventTimeline.plan, plan); assert.deepEqual(normalized.slides[0].eventTimeline.source, slide.eventTimeline);
  normalized.slides[0].eventTimeline.plan.stages[0].timing.leaves[0].durationMs = 88;
  assert.equal(slide.getEventTimelinePlan().stages[0].timing.leaves[0].durationMs, 10); assert.deepEqual(validatePresentation(document), []);
  assert.equal(replayEventTimeline(plan, { visitId: "v", timeMs: 5, elementIds: plan.elementIds, events: [{ id: "e", pageId: "p", visitId: "v", timeMs: 0, type: "object-click", elementId: "button" }] }).stages[0].samples[0].progress, .5);
});

test("atomic top-level reordering retains identities, timelines and rejects stale or incomplete permutations", () => {
  const { document, slide } = fixture(); slide.setEventTimeline(timeline('p', 'child', 'button'));
  const original = [...slide.elements], event = slide.eventTimeline;
  for (const ids of [['a', 'a', 'g'], ['a', 'button'], ['a', 'button', 'child'], ['a', 'button', 'g', 'extra'], [null, 'button', 'g']]) {
    assert.throws(() => slide.reorderElements(ids)); assert.deepEqual(slide.elements, original);
  }
  assert.throws(() => slide.reorderElements('g,button,a'), TypeError);
  slide.reorderElements(['g', 'button', 'a']);
  assert.deepEqual(slide.elements.map(item => item.id), ['g', 'button', 'a']);
  assert.equal(slide.elements[0], original[2]); assert.equal(slide.elements[1], original[1]); assert.equal(slide.eventTimeline, event);
  assert.equal(slide.getEventTimelinePlan().stages[0].target.elementId, 'child'); assert.deepEqual(validatePresentation(document), []);
  const malicious = ['g', 'button', 'a']; Object.defineProperty(malicious, 0, { get() { slide.moveElement('a', 0); return 'g'; } });
  assert.throws(() => slide.reorderElements(malicious), /changed during reorder/);
  assert.deepEqual(slide.elements.map(item => item.id), ['a', 'g', 'button']);
  slide.reorderElements(['a', 'g', 'button']);
  assert.deepEqual(slide.elements.map(item => item.id), ['a', 'g', 'button']);
  const empty = createPresentation().addSlide(); assert.doesNotThrow(() => empty.reorderElements([])); assert.throws(() => empty.reorderElements(['missing']));
});

test("nested group reordering keeps subtree identities and timing references with atomic rejection", () => {
  const document = createPresentation();
  const slide = document.addSlide({ id: 'p', elements: [shape('button'), { id: 'outer', type: 'group', box: { x: 0, y: 0, width: 200, height: 100 }, coordinateSize: { width: 200, height: 100 }, children: [shape('sibling'), { id: 'inner', type: 'group', box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [shape('first'), shape('second')] }] }] });
  slide.setEventTimeline(timeline('p', 'first', 'button'));
  const before = normalizePresentation(document), oldOuter = slide.elements[1], oldEvent = slide.eventTimeline;
  for (const [group, ids] of [['missing', ['second', 'first']], ['button', ['second', 'first']], ['inner', ['first', 'first']], ['inner', ['sibling', 'first']], ['outer', ['inner', 'first']]]) {
    assert.throws(() => slide.reorderGroupElements(group, ids)); assert.deepEqual(normalizePresentation(document), before);
  }
  slide.reorderGroupElements('inner', ['second', 'first']);
  const outer = slide.elements[1]; assert.notEqual(outer, oldOuter);
  assert.deepEqual(outer.children.map(child => child.id), ['sibling', 'inner']);
  assert.deepEqual(outer.children[1].children.map(child => child.id), ['second', 'first']);
  assert.deepEqual(oldOuter.children[1].children.map(child => child.id), ['first', 'second']);
  assert.equal(slide.eventTimeline, oldEvent); assert.equal(slide.getEventTimelinePlan().stages[0].target.elementId, 'first');
  assert.deepEqual(validatePresentation(document), []);
  const malicious = ['second', 'first']; Object.defineProperty(malicious, 0, { get() { slide.moveElement('button', 1); return 'second'; } });
  assert.throws(() => slide.reorderGroupElements('inner', malicious), /changed during reorder/);
  assert.deepEqual(slide.elements.map(item => item.id), ['outer', 'button']);
  assert.deepEqual(slide.elements[0].children[1].children.map(child => child.id), ['second', 'first']);
  slide.reorderGroupElements('outer', ['inner', 'sibling']);
  assert.deepEqual(slide.elements[0].children.map(child => child.id), ['inner', 'sibling']);
  const empty = createPresentation().addSlide({ elements: [{ id: 'empty', type: 'group', box: { x: 0, y: 0, width: 10, height: 10 }, coordinateSize: { width: 10, height: 10 }, children: [] }] });
  assert.doesNotThrow(() => empty.reorderGroupElements('empty', []));
  assert.throws(() => empty.reorderGroupElements('empty', ['missing']));
});

test("setter/init failures leave prior state and document identity registry unchanged", () => {
  const { document, slide } = fixture(); slide.setEventTimeline(timeline("p", "a")); const old = slide.eventTimeline;
  for (const bad of [timeline("wrong", "a"), timeline("p", "missing"), { stages: [{ ...timeline("p", "a").stages[0], target: { pageId: "p", elementId: "a", paragraphIndex: 0 } }] }]) assert.throws(() => slide.setEventTimeline(bad), EventTimelineError);
  assert.equal(slide.eventTimeline, old);
  assert.throws(() => document.addSlide({ id: "bad", elements: [shape("released")], eventTimeline: timeline("bad", "missing") }), EventTimelineError);
  assert.equal(document.slides.length, 1); assert.doesNotThrow(() => document.addSlide({ id: "bad", elements: [shape("released")] }));
  const overlong = "x".repeat(257); assert.throws(() => slide.addElement(shape(overlong)), EventTimelineError);
  assert.ok(!slide.elements.some(e => e.id === overlong));
  slide.setEventTimeline(); assert.doesNotThrow(() => slide.addElement(shape(overlong)), "failed timed insertion did not reserve the rejected ID");
  let calls = 0; assert.throws(() => slide.setEventTimeline({ get stages() { calls++; return []; } }), EventTimelineError); assert.equal(calls, 0);
});

test("deleting targets/trigger sources or groups containing them fails atomically until the timeline is explicitly cleared", () => {
  const { document, slide } = fixture(); slide.setEventTimeline(timeline("p", "child", "button")); const before = normalizePresentation(document);
  for (const id of ["g", "button"]) assert.throws(() => slide.removeElement(id), EventTimelineError);
  assert.deepEqual(normalizePresentation(document), before);
  const removed = slide.removeElement("a"); assert.equal(removed.id, "a"); assert.deepEqual(slide.getEventTimelinePlan().elementIds, ["button", "child", "g"]);
  slide.setEventTimeline(); assert.equal(slide.eventTimeline, undefined); slide.removeElement("g"); slide.removeElement("button"); assert.deepEqual(validatePresentation(document), []);
});

test("atomic identity-preserving replacement updates targeted content without remove/reinsert or reference churn", () => {
  const { document, slide } = fixture(); slide.setEventTimeline(timeline("p", "a", "button")); const prior = slide.getEventTimelinePlan();
  const updated = slide.replaceElement("a", { ...shape("a", "new"), id: undefined, box: { x: 40, y: 50, width: 120, height: 80 } });
  assert.equal(updated.id, "a"); assert.equal(updated.text.content, "new"); assert.deepEqual(slide.getEventTimelinePlan(), prior);
  assert.throws(() => slide.replaceElement("a", shape("different")), /retain/); assert.throws(() => slide.replaceElement("missing", shape("missing")), /Unknown/);
  assert.deepEqual(validatePresentation(document), []);
});

test("group replacement preserves retained child identities and rejects dropped targets or external ID collisions before registry commit", () => {
  const { document, slide } = fixture(); slide.setEventTimeline(timeline("p", "child", "button")); const original = slide.elements.find(e => e.id === "g");
  assert.throws(() => slide.replaceElement("g", { ...structuredClone(original), children: [shape("candidate")] }), EventTimelineError);
  assert.equal(slide.elements.find(e => e.id === "g"), original);
  assert.doesNotThrow(() => slide.addElement(shape("candidate")), "failed replacement did not reserve candidate IDs");
  assert.throws(() => slide.replaceElement("g", { ...structuredClone(original), children: [shape("child"), shape("a")] }), /Duplicate/);
  assert.equal(slide.elements.find(e => e.id === "g"), original);
  slide.replaceElement("g", { ...structuredClone(original), children: [shape("child", "changed"), shape("new-child")] });
  assert.equal(slide.eventTimeline.stages[0].target.elementId, "child"); assert.deepEqual(validatePresentation(document), []);
  assert.throws(() => slide.addElement(shape("new-child")), /Duplicate/);
});

test("whole-slide copy remaps page, targets and trigger sources using the allocator's complete descendant map", () => {
  const { document, slide } = fixture(); slide.setEventTimeline(timeline("p", "child", "button")); const source = normalizePresentation(document).slides[0];
  for (let i = 0; i < 3; i++) {
    const copy = document.duplicateSlide("p"), target = copy.elements.find(e => e.type === "group").children[0], button = copy.elements.find(e => e.type === "shape" && e.text.content === "button");
    assert.equal(copy.eventTimeline.stages[0].target.pageId, copy.id); assert.equal(copy.eventTimeline.stages[0].target.elementId, target.id); assert.equal(copy.eventTimeline.stages[0].trigger.elementId, button.id);
    assert.equal(copy.eventTimeline.stages[0].id, "stage"); assert.equal(copy.getEventTimelinePlan().stages[0].timing.leaves[0].id, "time");
  }
  assert.deepEqual(normalizePresentation(document).slides.find(s => s.id === "p"), source); document.removeSlide("p"); assert.deepEqual(validatePresentation(document), []);
  assert.doesNotThrow(() => document.slides[0].addElement(shape("child", "new unrelated object")), "released nested source IDs can be registered by a new owner");
  for (const copy of document.slides) assert.ok(copy.getEventTimelinePlan().elementIds.includes(copy.eventTimeline.stages[0].target.elementId));
});

test("object-only copies do not clone page-scoped gate stages; reorder does not retarget timing", () => {
  const { slide } = fixture(); slide.setEventTimeline(timeline("p", "a", "button")); const prior = slide.getEventTimelinePlan();
  const clone = slide.duplicateElement("a"); assert.notEqual(clone.id, "a"); assert.equal(slide.eventTimeline.stages.length, 1); assert.equal(slide.eventTimeline.stages[0].target.elementId, "a");
  slide.moveElement("a", slide.elements.length - 1); assert.deepEqual(slide.getEventTimelinePlan().stages, prior.stages);
  slide.removeElement(clone.id); assert.deepEqual(slide.getEventTimelinePlan(), prior);
});

test("timeline targets cannot silently bind inherited layout objects and diagnostics validate structural document consumers", () => {
  const { document, slide } = fixture(); document.defineSlideLayout({ id: "layout", name: "Layout", elements: [shape("layout-shape")] });
  const owned = document.addSlide({ id: "owned", layoutId: "layout", elements: [shape("owned-object")] });
  assert.throws(() => owned.setEventTimeline(timeline("owned", "layout-shape")), EventTimelineError);
  const fake = { ...document, slides: [{ ...slide, id: slide.id, elements: slide.elements, eventTimeline: timeline("p", "missing") }], assets: document.assets, layouts: document.layouts, theme: document.theme, textStylePresets: document.textStylePresets };
  const diagnostics = validatePresentation(fake); assert.ok(diagnostics.some(d => d.code === "invalid-event-timeline-input" && d.slideId === "p" && d.path.includes("eventTimeline")));
});

test("released slide snapshots cannot replace or unregister identities reused by live objects", () => {
  const document = createPresentation(), old = document.addSlide({ id: "old", elements: [shape("reused")] });
  document.removeSlide(old.id);
  const live = document.addSlide({ id: "live", elements: [shape("reused", "live value")] });
  assert.throws(() => old.replaceElement("reused", shape("reused", "stale overwrite")), /released/);
  old.removeElement("reused");
  assert.throws(() => live.addElement(shape("reused")), /Duplicate/);
  assert.equal(live.elements[0].text.content, "live value");
  live.replaceElement("reused", shape("reused", "live update")); assert.deepEqual(validatePresentation(document), []);
});
