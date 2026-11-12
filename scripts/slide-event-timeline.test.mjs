import assert from "node:assert/strict";
import test from "node:test";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createPresentation, normalizePresentation, replayEventTimeline } from "../packages/core/dist/index.js";
import { resolveLayout, resolveNormalizedLayout } from "../packages/layout/dist/index.js";
import { generatePptx, PptxExportError } from "../packages/pptx-exporter/dist/index.js";
import { writePptx } from "../packages/pptx-exporter/dist/node.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
function fixture() {
  const document = createPresentation({ size: { width: 400, height: 300 } }), slide = document.addSlide({ id: "slide" });
  slide.addElement({ id: "text", type: "text", box: { x: 10, y: 10, width: 120, height: 40 }, content: "Control text" });
  slide.addElement({ id: "table", type: "table", box: { x: 20, y: 70, width: 160, height: 120 }, columns: [80, 80], rows: Array.from({ length: 4 }, (_, i) => ({ height: 30, cells: [{ content: `A${i}` }, { content: `B${i}` }] })) });
  const timeline = { stages: [{ id: "stage", target: { pageId: slide.id, elementId: "text" }, trigger: { type: "object-click", elementId: "table" }, timing: { kind: "span", id: "time", durationMs: 20 } }] };
  return { document, slide, timeline };
}

test("IR/JSON/Layout retain exact event targets and detached schedules; static SVG reports rather than applies timing", async () => {
  const { document, slide, timeline } = fixture(), before = await renderPresentationToSvg(document);
  slide.setEventTimeline(timeline);
  const normalized = normalizePresentation(document), transported = JSON.parse(JSON.stringify(normalized)), layout = resolveNormalizedLayout(transported);
  assert.deepEqual(layout.slides[0].eventTimeline.plan, slide.getEventTimelinePlan());
  assert.deepEqual(layout.slides[0].eventTimeline.source, slide.eventTimeline);
  layout.slides[0].eventTimeline.plan.stages[0].target.elementId = "changed"; assert.equal(transported.slides[0].eventTimeline.plan.stages[0].target.elementId, "text");
  assert.equal(slide.eventTimeline.stages[0].target.elementId, "text");
  const after = await renderPresentationToSvg(document); assert.equal(after.slides[0].svg, before.slides[0].svg);
  assert.deepEqual(after.warnings.map(w => [w.code, w.slideId]), [["event-timeline-static", "slide"]]);
  const valid = resolveLayout(document).slides[0].eventTimeline.plan;
  const result = replayEventTimeline(valid, { visitId: "v", timeMs: 5, elementIds: ["text", "table"], events: [{ id: "click", pageId: "slide", visitId: "v", timeMs: 0, type: "object-click", elementId: "table" }] });
  assert.equal(result.stages[0].samples[0].progress, .25);
});

test("direct normalized layout rejects wrong owner page, stale targets/sources and corrupt future timing", () => {
  const { document, slide, timeline } = fixture(); slide.setEventTimeline(timeline); const normalized = normalizePresentation(document);
  for (const change of [n => { n.slides[0].eventTimeline.plan.pageId = "wrong"; }, n => { n.slides[0].elements.pop(); }, n => { n.slides[0].eventTimeline.plan.stages[0].timing.boundaries.pop(); }, n => { n.slides[0].eventTimeline.source.stages[0].timing.durationMs = 99; }, n => { n.slides[0].eventTimeline.schemaVersion = "future"; }, n => { delete n.slides[0].eventTimeline.source; }]) {
    const malformed = structuredClone(normalized); change(malformed); assert.throws(() => resolveNormalizedLayout(malformed));
  }
  assert.deepEqual(resolveNormalizedLayout(normalized).slides[0].eventTimeline, normalized.slides[0].eventTimeline);
});

test("pagination refuses timed source slides atomically instead of cloning stale references onto continuations", () => {
  const { document, slide, timeline } = fixture(); slide.setEventTimeline(timeline); const normalized = normalizePresentation(document), before = structuredClone(normalized);
  const rule = { elementId: "table", area: { x: 20, y: 70, width: 160, height: 60 } };
  const timed = resolveNormalizedLayout(normalized, { tablePagination: [rule] }); assert.equal(timed.slideCount, 1); assert.equal(timed.slides[0].elements[1].rows.length, 4);
  assert.ok(timed.analysis.diagnostics.some(d => d.code === "pagination-event-timeline-unsupported")); assert.deepEqual(normalized, before);
  const fitting = resolveLayout(document, { tablePagination: [{ ...rule, area: { ...rule.area, height: 120 } }] });
  assert.equal(fitting.slideCount, 1); assert.ok(!fitting.analysis.diagnostics.some(d => d.code === "pagination-event-timeline-unsupported"));
  assert.deepEqual(fitting.slides[0].eventTimeline, normalized.slides[0].eventTimeline);
  slide.setEventTimeline(); assert.equal(resolveLayout(document, { tablePagination: [rule] }).slideCount, 2);
});

test("PPTX defaults to located immutable refusal, explicit omit reports each stage and retains static package payloads", async () => {
  const { document, slide, timeline } = fixture(), control = await generatePptx(document); slide.setEventTimeline(timeline);
  await assert.rejects(generatePptx(document), error => error instanceof PptxExportError && error.code === "event-timeline-export-unsupported" && error.diagnostics[0].slideId === "slide" && error.diagnostics[0].elementId === "text" && Object.isFrozen(error.diagnostics[0]) && Object.isFrozen(error.diagnostics));
  const saved = await generatePptx(document, { eventTimelineMode: "omit" }); assert.equal(saved.status, "generated-with-warnings");
  assert.deepEqual(saved.warnings.map(w => [w.code, w.slideId, w.stageId, w.elementId]), [["event-timeline-omitted", "slide", "stage", "text"]]);
  assert.deepEqual(readGeneratedZip(saved.bytes), readGeneratedZip(control.bytes)); assert.equal(slide.eventTimeline.stages[0].id, "stage");
  for (const bad of [true, "native", null]) await assert.rejects(generatePptx(document, { eventTimelineMode: bad }), TypeError);
  const reopened = await importPptx(saved.bytes, { projection: "content" }); try { assert.equal(reopened.document.slides[0].eventTimeline, undefined); } finally { reopened.dispose(); }
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptkit-timed-export-")), output = path.join(dir, "control.pptx");
  try { await writeFile(output, "keep"); await assert.rejects(writePptx(document, { output }), PptxExportError); assert.equal(await readFile(output, "utf8"), "keep"); } finally { await rm(dir, { recursive: true, force: true }); }
});

test("explicit omission happens before pagination and asset resolution; default refusal performs no asset work", async () => {
  const { document, slide, timeline } = fixture(); slide.setEventTimeline(timeline);
  const asset = document.registerAsset({ kind: "image", mimeType: "image/png", source: { type: "url", value: "https://example.invalid/no-network.png" } });
  slide.addElement({ type: "image", assetId: asset.id, box: { x: 250, y: 40, width: 20, height: 20 } });
  let calls = 0; await assert.rejects(generatePptx(document, { resolveAsset: () => { calls++; throw Error("must not resolve"); } }), PptxExportError); assert.equal(calls, 0);
  // A separate asset-free fixture verifies explicit omission enables the requested structural transform.
  const clean = fixture(); clean.slide.setEventTimeline(clean.timeline);
  const output = await generatePptx(clean.document, { eventTimelineMode: "omit", layout: { tablePagination: [{ elementId: "table", area: { x: 20, y: 70, width: 160, height: 60 } }] } });
  assert.equal(output.slideCount, 2); assert.ok(output.warnings.some(w => w.code === "event-timeline-omitted")); assert.ok(!output.warnings.some(w => w.code === "pagination-event-timeline-unsupported"));
});

test("Parser source saves block authored timeline changes atomically and recover after explicit clear", async () => {
  const { document } = fixture(), source = await generatePptx(document), imported = await importPptx(source.bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
  try {
    const slide = imported.document.slides[0], text = slide.elements.find(e => e.type === "text");
    slide.setEventTimeline({ stages: [{ id: "stage", target: { pageId: slide.id, elementId: text.id }, trigger: { type: "enter" }, timing: { kind: "span", id: "t", durationMs: 1 } }] });
    const changed = structuredClone(text); changed.content[0].runs[0].text = "Changed"; slide.replaceElement(text.id, changed);
    const blocked = await imported.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.ok(blocked.report.diagnostics.some(d => d.code === "save-edit-unsupported"));
    slide.setEventTimeline(); const ready = await imported.prepareSave({ mode: "edit" }); assert.equal(ready.status, "ready");
    const saved = await imported.save(ready.plan); assert.notDeepEqual(saved.bytes, source.bytes);
  } finally { imported.dispose(); }
});

test("bound IR retains authored timing group identities/topology and rejects plan-only semantic drift", () => {
  const { document, slide, timeline } = fixture();
  timeline.stages[0].timing = { id: "semantic-root", kind: "sequence", children: [{ id: "empty-parent", kind: "parallel", delayMs: 2, children: [] }, { id: "leaf", kind: "span", durationMs: 5 }] };
  slide.setEventTimeline(timeline); const normalized = normalizePresentation(document), bound = normalized.slides[0].eventTimeline;
  assert.deepEqual(bound.source, slide.eventTimeline); assert.equal(bound.source.stages[0].timing.children[0].id, "empty-parent");
  const restored = resolveNormalizedLayout(JSON.parse(JSON.stringify(normalized))); assert.deepEqual(restored.slides[0].eventTimeline.source, timeline);
  const drift = structuredClone(normalized), plan = drift.slides[0].eventTimeline.plan.stages[0].timing;
  plan.durationMs++; plan.leaves[0].durationMs++; plan.leaves[0].endMs++; plan.boundaries.find(b => b.type === "end").timeMs++;
  // This flat schedule remains valid on its own; the retained source disproves it.
  assert.doesNotThrow(() => replayEventTimeline(drift.slides[0].eventTimeline.plan, { visitId: "v", timeMs: 0, elementIds: ["text", "table"], events: [] }));
  assert.throws(() => resolveNormalizedLayout(drift), /disagree/);
  bound.source.stages[0].timing.id = "transport-only-change";
  assert.equal(slide.eventTimeline.stages[0].timing.id, "semantic-root");
});
