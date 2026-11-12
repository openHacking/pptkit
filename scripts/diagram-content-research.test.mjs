import assert from "node:assert/strict";
import test from "node:test";
import { projectDiagramContent } from "./research/diagram-content-model.mjs";
const uri = family => `urn:microsoft.com/office/officeart/2005/8/layout/${family}`;
const point = (elementIndex, kind = "node", text = kind === "doc" ? "" : `${elementIndex}`) => ({ elementIndex, kind, modelId: String(elementIndex), identityStatus: "unique", text: [{ runs: [{ kind: "text", text }] }] });
const edge = (elementIndex, from, to, order = 0) => ({ elementIndex, kind: "parOf", identityStatus: "unique", source: { status: "resolved", targetElementIndex: from }, destination: { status: "resolved", targetElementIndex: to }, sourceOrderValue: order });
const fixture = () => ({ part: "/source/data.xml", status: "inspected", points: [point(1, "doc", ""), point(2), point(3), point(4), point(5)], connections: [edge(8, 1, 4, 20), edge(9, 2, 3), edge(10, 1, 2, 10), edge(11, 4, 5)] });
const run = (m = fixture(), layout = "process3", limits) => projectDiagramContent(m, uri(layout), limits);

test("process3 retains child content; numeric source order wins over XML/point ordering", () => {
  const m = fixture(), before = structuredClone(m), r = run(m);
  assert.deepEqual(r.candidate.roots, ["n1", "n3"]);
  assert.deepEqual(r.candidate.nodes.map(n => [n.id, n.children, n.paragraphs]), [["n1", ["n2"], ["2"]], ["n2", [], ["3"]], ["n3", ["n4"], ["4"]], ["n4", [], ["5"]]]);
  assert.deepEqual(r.bindings.map(b => b.elementIndex), [2, 3, 4, 5]); assert.deepEqual(m, before);
  m.points.reverse(); m.connections.reverse(); assert.deepEqual(run(m), r);
  r.candidate.nodes[0].paragraphs[0] = "changed"; assert.equal(run(m).candidate.nodes[0].paragraphs[0], "2");
});

test("hierarchy keeps multiple roots, assistant role and paragraph boundaries", () => {
  const m = fixture(); m.points[2].kind = "asst"; m.points[1].text = [{ runs: [{ kind: "text", text: "Manager 😀" }, { kind: "break" }, { kind: "text", text: "continued" }] }, { runs: [] }];
  const r = run(m, "orgChart1"); assert.equal(r.candidate.nodes[1].role, "assistant"); assert.deepEqual(r.candidate.nodes[0].paragraphs, ["Manager 😀\ncontinued", ""]);
  assert.equal(r.candidate.roots.length, 2);
  assert.throws(() => run(m), { code: "unsupported-assistant" });
});

test("cycle2 is an ordered flat content forest with layout-generated closure, never a parent graph cycle", () => {
  const m = fixture(); m.connections = [edge(8, 1, 4, 2), edge(9, 1, 2, 0), edge(10, 1, 5, 3), edge(11, 1, 3, 1)];
  const r = run(m, "cycle2"); assert.equal(r.candidate.closure, "layout-generated"); assert.equal(r.candidate.roots.length, 4); assert.ok(r.candidate.nodes.every(n => n.children.length === 0));
  assert.deepEqual(r.bindings.map(b => b.elementIndex), [2, 3, 4, 5]);
  assert.throws(() => run(fixture(), "cycle2"), { code: "nested-cycle" });
  m.points = [point(1, "doc", "")]; m.connections = []; assert.deepEqual(run(m, "cycle2").candidate.roots, []);
  m.points.push(point(2)); m.connections.push(edge(3, 1, 2)); assert.equal(run(m, "cycle2").candidate.nodes.length, 1);
});

test("presentation/transition records remain in separate source coverage and never become visible authoring nodes", () => {
  const m = fixture(); m.points.push(point(6, "pres"), point(7, "sibTrans")); m.connections.push({ ...edge(12, 2, 6), kind: "presOf" });
  const r = run(m); assert.equal(r.source.ignoredPoints, 2); assert.equal(r.source.ignoredConnections, 1);
  assert.equal(r.candidate.nodes.length, 4); assert.ok(!JSON.stringify(r.candidate).includes("/source/"));
  m.connections.push(edge(13, 7, 6)); assert.throws(() => run(m), { code: "unsupported-edge" });
});

test("unknown/partial sources, fields, duplicate identities/orders and unresolved topology reject without repair", () => {
  const mutations = [
    [m => { m.status = "partial"; }, "incomplete-source"],
    [m => { m.points[1].kind = "future"; }, "unknown-point"],
    [m => { m.points[1].identityStatus = "duplicate"; }, "source-identity"],
    [m => { m.connections[0].destination.status = "ambiguous"; }, "source-reference"],
    [m => { m.connections[0].sourceOrderValue = 10; }, "sibling-order"],
    [m => { m.connections.push(edge(12, 1, 2, 30)); }, "multiple-parents"],
    [m => { m.points.push(point(99)); }, "disconnected"],
    [m => { m.points[0].kind = "node"; }, "document-root"],
    [m => { m.points[0].text[0].runs[0].text = "Meaningful root text"; }, "document-text"],
    [m => { m.points.push(point(99, "doc")); }, "document-root"],
    [m => { m.connections[0].kind = "unknownRelationship"; }, "unknown-connection"],
    [m => { m.points[1].text[0].runs[0].kind = "field"; }, "unsupported-field"],
  ];
  for (const [mutate, code] of mutations) { const m = fixture(); mutate(m); assert.throws(() => run(m), { code }); }
  assert.throws(() => projectDiagramContent(fixture(), uri("process3-custom")), { code: "unknown-layout" });
  const cycle = fixture(); cycle.connections = [edge(8, 2, 3), edge(9, 3, 2), edge(10, 1, 4), edge(11, 4, 5)]; assert.throws(() => run(cycle), { code: "disconnected" });
});

test("iterative traversal bounds deep/wide forests and preserves failed-input/retry state", () => {
  const m = fixture(), before = structuredClone(m);
  for (const limits of [{ maxNodes: 4 }, { maxEdges: 3 }, { maxDepth: 1 }, { maxWork: 2 }]) assert.throws(() => run(m, "process3", limits));
  assert.deepEqual(m, before); assert.equal(run(m).candidate.nodes.length, 4);
  for (const limits of [{ maxWork: NaN }, { unknown: 1 }, { maxDepth: 0 }]) assert.throws(() => run(m, "process3", limits), { code: "invalid-limit" });
  const deep = { part: "/deep", status: "inspected", points: Array.from({ length: 1500 }, (_, i) => point(i + 1, i ? "node" : "doc")), connections: Array.from({ length: 1499 }, (_, i) => edge(2000 + i, i + 1, i + 2)) };
  assert.throws(() => run(deep), { code: "depth-limit" }); assert.equal(run(deep, "orgChart1", { maxDepth: 1500 }).candidate.nodes.length, 1499);
});

test("seeded forests preserve every node/parent/order under shuffled storage", () => {
  let seed = 1981; const random = n => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % n; };
  for (let round = 0; round < 100; round++) {
    const size = 2 + random(40), m = { part: "/seeded", status: "inspected", points: [point(1, "doc")], connections: [] }, parent = new Map(), counts = new Map();
    for (let i = 2; i <= size; i++) {
      const from = 1 + random(i - 1), order = counts.get(from) ?? 0; counts.set(from, order + 1); parent.set(i, from);
      m.points.push(point(i)); m.connections.push(edge(100 + i, from, i, order));
    }
    const expected = run(m, "orgChart1"); m.points.reverse(); m.connections.reverse(); const actual = run(m, "orgChart1"); assert.deepEqual(actual, expected);
    const source = new Map(actual.bindings.map(b => [b.id, b.elementIndex]));
    assert.equal(actual.candidate.nodes.length, size - 1);
    for (const node of actual.candidate.nodes) for (const child of node.children) assert.equal(parent.get(source.get(child)), source.get(node.id));
    for (const root of actual.candidate.roots) assert.equal(parent.get(source.get(root)), 1);
  }
});
