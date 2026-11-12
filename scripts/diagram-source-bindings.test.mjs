import assert from "node:assert/strict";
import test from "node:test";
import { bindDiagramSourceContent } from "./research/diagram-source-bindings.mjs";
import { projectDiagramContent, projectDiagramContentBudgeted } from "./research/diagram-content-model.mjs";
const layout = family => `urn:microsoft.com/office/officeart/2005/8/layout/${family}`;
const namespace = "12345678-1234-4ABC-8DEF", scope = "document-A/diagram-A";
const pt = (elementIndex, modelId, kind = "node", label = modelId) => ({ elementIndex, modelId, kind, identityStatus: "unique", text: [{ runs: [{ kind: "text", text: kind === "doc" ? "" : label }] }] });
const edge = (i, from, to, sourceOrderValue = 0) => ({ elementIndex: i, kind: "parOf", identityStatus: "unique", source: { status: "resolved", targetElementIndex: from }, destination: { status: "resolved", targetElementIndex: to }, sourceOrderValue });
const fixture = () => ({ part: "/ppt/diagrams/data1.xml", status: "inspected", findings: [], points: [pt(1, "0", "doc"), pt(2, "2"), pt(3, "3"), pt(4, "4")], connections: [edge(10, 1, 2, 0), edge(11, 2, 3, 0), edge(12, 1, 4, 1)] });
const options = extra => ({ scope, namespace, ...extra });
const run = (model = fixture(), extra, name = "process3") => bindDiagramSourceContent(model, layout(name), options(extra));
const map = r => new Map(r.bindings.map(b => [b.canonicalModelId, b.id]));
function reindex(model) {
  for (const p of model.points) p.elementIndex += 100;
  for (const e of model.connections) { e.elementIndex += 1000; e.source.targetElementIndex += 100; e.destination.targetElementIndex += 100; }
}

test("explicit source scope binds raw model identities independently of XML indices, text and traversal order", () => {
  const source = fixture(), before = structuredClone(source), a = run(source);
  assert.deepEqual(source, before); assert.equal(new Set(a.candidate.nodes.map(n => n.id)).size, 3);
  assert.ok(a.candidate.nodes.every(n => /^\{12345678-1234-4ABC-8DEF-[0-9A-F]{12}\}$/.test(n.id)));
  const originalState = structuredClone(a.state), bSource = fixture(); reindex(bSource); bSource.points.reverse(); bSource.connections.reverse();
  for (const e of bSource.connections) if (e.source.targetElementIndex === 101) e.sourceOrderValue = 1 - e.sourceOrderValue;
  bSource.points.find(p => p.modelId === "2").text = [{ runs: [{ kind: "text", text: "Updated 😀" }, { kind: "break" }, { kind: "text", text: "正文" }] }, { runs: [] }];
  const b = run(bSource, { state: a.state }); assert.deepEqual(a.state, originalState); assert.deepEqual(map(b), map(a));
  assert.deepEqual(b.candidate.roots, [map(a).get("int:4"), map(a).get("int:2")]);
  assert.deepEqual(b.candidate.nodes.find(n => n.id === map(a).get("int:2")).paragraphs, ["Updated 😀\n正文", ""]);
  assert.equal(b.bindings.find(n => n.canonicalModelId === "int:2").elementIndex, 102);
  assert.ok(!JSON.stringify(b.candidate).includes(source.part), "format provenance is outside the candidate");
  b.bindings[0].sourceModelId = "mutated"; b.state.identities.entries[0].key = "mutated"; b.candidate.nodes[0].paragraphs[0] = "mutated";
  assert.deepEqual(a.state, originalState); assert.deepEqual(run(fixture()).state, originalState);
  const budgeted = projectDiagramContentBudgeted(source, layout("process3")); assert.deepEqual(budgeted.projection, projectDiagramContent(source, layout("process3"))); assert.ok(budgeted.work > 0);
});

test("recognized numeric spellings canonicalize but GUID lexical profile and original source spelling stay explicit", () => {
  const a = run(), source = fixture(); source.points[0].modelId = " +0000\r\n"; source.points[1].modelId = "\t+0002 ";
  const b = run(source, { state: a.state }); assert.deepEqual(map(a), map(b));
  assert.equal(b.bindings.find(b => b.canonicalModelId === "int:2").sourceModelId, "\t+0002 ");
  for (const raw of ["2147483648", "-2147483649", "{abcdef00-0000-0000-0000-000000000001}", "1.0", "NaN", "0".repeat(1025)]) {
    const m = fixture(); m.points[1].modelId = raw; assert.throws(() => run(m), TypeError);
  }
  const duplicate = fixture(); duplicate.points[2].modelId = "+02"; assert.throws(() => run(duplicate), /Duplicate canonical/);
  const guid = fixture(); guid.points[1].modelId = "{ABCDEF00-0000-0000-0000-000000000001}";
  const g = run(guid); assert.equal(g.bindings[0].canonicalModelId, "guid:{ABCDEF00-0000-0000-0000-000000000001}");
});

test("source additions allocate fresh identities; retired source IDs require exact explicit restoration", () => {
  const a = run(), removed = fixture(); removed.points = removed.points.filter(p => p.modelId !== "3"); removed.connections = removed.connections.filter(e => e.destination.targetElementIndex !== 3);
  const b = run(removed, { state: JSON.parse(JSON.stringify(a.state)) });
  const returning = fixture(); assert.throws(() => run(returning, { state: b.state }), /Retired source identity/);
  const restored = run(returning, { state: b.state, restore: ["+003"] }); assert.deepEqual(map(restored), map(a));
  for (const restore of [["3", "+03"], ["2"], ["99"], ["0"], "3"]) assert.throws(() => run(returning, { state: b.state, restore }), TypeError);
  const inserted = structuredClone(removed); inserted.points.push(pt(5, "99")); inserted.connections.push(edge(13, 1, 5, 2));
  const c = run(inserted, { state: b.state }); assert.equal(map(c).get("int:2"), map(a).get("int:2"));
  assert.notEqual(map(c).get("int:99"), map(a).get("int:3")); assert.equal(c.state.identities.entries.length, a.state.identities.entries.length + 1);
  assert.equal(b.state.identities.entries.length, a.state.identities.entries.length);
});

test("contexts, roots, malformed reservation keys and state collections cannot silently rebind", () => {
  const a = run(), mutations = [s => { s.version = 2; }, s => { s.scope = "other"; }, s => { s.part = "/other"; },
    s => { s.layoutIdentifier = layout("orgChart1"); }, s => { s.rootIdentity = "int:22"; },
    s => { s.activeSourceIds.push(s.activeSourceIds[0]); }, s => { s.activeSourceIds.push("int:99"); },
    s => { s.activeSourceIds.push("int:0"); }, s => { s.identities.entries[1].key = '["foreign","int:2"]'; },
    s => { s.identities.entries[1].key = '["source-point","int:+02"]'; }, s => { s.identities.entries[0].key = '["source-point","int:100"]'; },
    s => { s.identities.entries[1].modelId = s.identities.entries[0].modelId; }, s => { s.extra = true; }];
  for (const mutate of mutations) { const state = structuredClone(a.state); mutate(state); const before = structuredClone(state); assert.throws(() => run(fixture(), { state })); assert.deepEqual(state, before); }
  const newRoot = fixture(); newRoot.points[0].modelId = "10"; assert.throws(() => run(newRoot, { state: a.state }), /context\/root/);
  assert.throws(() => run(fixture(), { state: a.state, namespace: "FFFFFFFF-FFFF-4FFF-8FFF" }), /namespace/);
  for (const scope of [undefined, "", "a".repeat(257)]) assert.throws(() => run(fixture(), { scope }));
  for (const part of ["relative.xml", "a".repeat(2050)]) { const m = fixture(); m.part = part; assert.throws(() => run(m)); }
  assert.throws(() => run(fixture(), { extra: true }));
});

test("source projection and binding share exact work/key budgets and preserve baseline on failure", () => {
  const a = run(), baseline = structuredClone(a.state);
  for (const limits of [{ maxWork: 1 }, { maxKeyBytes: 1 }, { maxEntries: 3 }, { maxNodes: 3 }, { maxDepth: 1 }]) {
    assert.throws(() => run(fixture(), { state: a.state, limits })); assert.deepEqual(a.state, baseline);
  }
  for (const limits of [{ unknown: 2 }, { maxWork: NaN }, { maxKeyBytes: 0 }]) assert.throws(() => run(fixture(), { limits }));
  const result = run(fixture(), { state: a.state });
  assert.deepEqual(run(fixture(), { state: a.state, limits: { maxWork: result.work } }), result);
  assert.throws(() => run(fixture(), { state: a.state, limits: { maxWork: result.work - 1 } }), /maxWork/);
  let lo = 1, hi = 100000;
  while (lo < hi) { const mid = Math.floor((lo + hi) / 2); try { run(fixture(), { state: a.state, limits: { maxKeyBytes: mid } }); hi = mid; } catch (e) { assert.match(e.message, /maxKeyBytes/); lo = mid + 1; } }
  assert.deepEqual(run(fixture(), { state: a.state, limits: { maxKeyBytes: lo } }), result);
  assert.throws(() => run(fixture(), { state: a.state, limits: { maxKeyBytes: lo - 1 } }), /maxKeyBytes/);
  assert.deepEqual(a.state, baseline);
});

test("hierarchy roles and Cycle layout closure survive managed binding without source-style or write authority", () => {
  const hierarchy = fixture(); hierarchy.points[2].kind = "asst";
  const a = run(hierarchy, undefined, "orgChart1"); assert.equal(a.candidate.nodes[1].role, "assistant");
  assert.deepEqual(a.candidate.nodes[0].children, [a.candidate.nodes[1].id]);
  const flat = fixture(); flat.connections[1].source.targetElementIndex = 1; flat.connections[1].sourceOrderValue = 2;
  const cycle = run(flat, undefined, "cycle2"); assert.equal(cycle.candidate.closure, "layout-generated"); assert.equal(cycle.candidate.roots.length, 3);
  const empty = fixture(); empty.points = [empty.points[0]]; empty.connections = [];
  const e = run(empty, undefined, "cycle2"); assert.deepEqual(e.candidate.nodes, []); assert.equal(e.state.identities.entries.length, 1);
  for (const mutation of [m => { m.status = "source-only"; }, m => { m.points[1].text[0].runs[0].kind = "field"; }, m => { m.points[1].identityStatus = "duplicate"; }, m => { m.findings = [{ code: "uninterpreted-structure" }]; }, m => { delete m.findings; }]) {
    const m = fixture(); mutation(m); assert.throws(() => run(m));
  }
});

test("seeded reinspection binds identical labels by source identity with explicit retirement restoration", () => {
  let seed = 18473, state, nextSerial = 2, previousActive = new Set();
  const known = new Map(), random = () => (seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0);
  for (let round = 0; round < 100; round++) {
    const chosen = new Set(); for (let i = 0, size = random() % 25; i < size; i++) chosen.add(1 + random() % 35);
    const order = [...chosen], base = round * 100 + 1, model = { part: "/ppt/diagrams/data1.xml", status: "inspected", findings: [], points: [pt(base, "0", "doc")], connections: [] };
    const restore = order.filter(key => known.has(key) && !previousActive.has(key)).map(String);
    for (const [index, key] of order.entries()) {
      model.points.push(pt(base + index + 1, round % 2 ? ` +00${key} ` : String(key), "node", "identical label")); model.connections.push(edge(100000 + index, base, base + index + 1, index));
      if (!known.has(key)) known.set(key, `{${namespace}-${(nextSerial++).toString(16).toUpperCase().padStart(12, "0")}}`);
    }
    if (round % 2) { model.points.reverse(); model.connections.reverse(); }
    const result = run(model, { state, restore }, "cycle2");
    assert.deepEqual(result.candidate.roots, order.map(key => known.get(key)));
    assert.equal(result.state.identities.nextSerial, nextSerial); assert.equal(result.state.identities.entries.length, known.size + 1);
    assert.deepEqual(result.bindings.map(b => b.canonicalModelId), order.map(key => `int:${key}`));
    previousActive = chosen; state = JSON.parse(JSON.stringify(result.state));
  }
});
