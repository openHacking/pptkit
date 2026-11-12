import assert from "node:assert/strict";
import test from "node:test";
import { resolveCycleLayout } from "./research/cycle-layout.mjs";
const input = n => ({ family: "cycle", closure: "layout-generated", roots: Array.from({ length: n }, (_, i) => `n${i}`), nodes: Array.from({ length: n }, (_, i) => ({ id: `n${i}`, role: "content", children: [], paragraphs: [`Node ${i}`] })) });
const frame = { x: 10, y: 20, width: 200, height: 200 };
const near = (a, b, tolerance = 1e-8) => assert.ok(Math.abs(a - b) <= tolerance, `${a} != ${b}`);
const pointNear = (a, b) => { near(a.x, b.x); near(a.y, b.y); };
const distance = (a, b) => Math.hypot(a.x - b.x, a.y - b.y);
const segmentDistance = (p, a, b) => {
  const dx = b.x - a.x, dy = b.y - a.y, t = Math.max(0, Math.min(1, ((p.x - a.x) * dx + (p.y - a.y) * dy) / (dx * dx + dy * dy)));
  return distance(p, { x: a.x + t * dx, y: a.y + t * dy });
};

test("hand-calculated square and triangle cycles fix node centers, packing, arrow direction and closure", () => {
  const square = resolveCycleLayout(input(4), { frame, nodeDiameter: 40 });
  const r = 30 * Math.SQRT2;
  const centers = [{ x: 110, y: 120 - r }, { x: 110 + r, y: 120 }, { x: 110, y: 120 + r }, { x: 110 - r, y: 120 }];
  square.nodes.forEach((node, i) => { pointNear(node.center, centers[i]); near(node.diameter, 40); });
  pointNear(square.orbit.center, { x: 110, y: 120 }); near(square.orbit.radius, r);
  assert.deepEqual(square.edges.map(e => [e.sourceId, e.targetId]), [["n0", "n1"], ["n1", "n2"], ["n2", "n3"], ["n3", "n0"]]);
  pointNear(square.edges[0].polygon[3], { x: 110 + 35 / Math.SQRT2, y: 120 - 25 / Math.SQRT2 });
  const tri = resolveCycleLayout(input(3), { frame: { x: 0, y: 0, width: 200, height: 200 }, nodeDiameter: 40 });
  const triangle = [{ x: 100, y: 100 - 15 * Math.sqrt(3) }, { x: 130, y: 100 + 15 * Math.sqrt(3) }, { x: 70, y: 100 + 15 * Math.sqrt(3) }];
  tri.nodes.forEach((node, i) => pointNear(node.center, triangle[i]));
  near(tri.orbit.radius, 20 * Math.sqrt(3));
  const auto = resolveCycleLayout(input(4), { frame }); near(auto.diameter, 200 / (1 + 3 / Math.SQRT2));
  assert.ok(auto.findings.every(f => f.code === "text-unmeasured")); assert.equal(auto.nativeAdmission, "untested");
});

test("empty and singleton policies are explicit; two nodes have disjoint clipped directed arcs", () => {
  const empty = resolveCycleLayout(input(0), { frame }); assert.deepEqual(empty.nodes, []); assert.deepEqual(empty.edges, []); assert.equal(empty.diameter, 0);
  const one = resolveCycleLayout(input(1), { frame }); near(one.diameter, 200); pointNear(one.nodes[0].center, { x: 110, y: 120 }); assert.deepEqual(one.edges, []);
  const two = resolveCycleLayout(input(2), { frame }); assert.equal(two.edges.length, 2); assert.ok(two.edges.every(e => e.kind === "arc"));
  assert.notDeepEqual(two.edges[0].head, two.edges[1].head);
  const nodes = new Map(two.nodes.map(n => [n.id, n]));
  for (const edge of two.edges) {
    near(distance(edge.start, nodes.get(edge.sourceId).center), two.diameter / 2);
    near(distance(edge.end, nodes.get(edge.targetId).center), two.diameter / 2);
    const a = Math.atan2(edge.start.y - edge.center.y, edge.start.x - edge.center.x), b = Math.atan2(edge.end.y - edge.center.y, edge.end.x - edge.center.x);
    const delta = (b - a + Math.PI * 2) % (Math.PI * 2);
    for (let i = 0; i <= 40; i++) {
      const angle = a + delta * i / 40, p = { x: edge.center.x + edge.radius * Math.cos(angle), y: edge.center.y + edge.radius * Math.sin(angle) };
      for (const node of two.nodes) assert.ok(distance(p, node.center) >= two.diameter / 2 - 1e-8);
    }
    for (const node of two.nodes) for (let i = 0; i < edge.head.length; i++) assert.ok(segmentDistance(node.center, edge.head[i], edge.head[(i + 1) % edge.head.length]) >= two.diameter / 2 - 1e-8);
  }
});

test("root ordering, start angle, direction, rotation and mirrors preserve semantic identity", () => {
  const c = input(4), before = structuredClone(c), baseline = resolveCycleLayout(c, { frame, nodeDiameter: 40 });
  const shuffled = structuredClone(c); shuffled.nodes.reverse(); assert.deepEqual(resolveCycleLayout(shuffled, { frame, nodeDiameter: 40 }), baseline);
  const ccw = resolveCycleLayout(c, { frame, nodeDiameter: 40, direction: "counterclockwise" }); pointNear(ccw.nodes[1].center, baseline.nodes[3].center);
  const rotated = resolveCycleLayout(c, { frame, nodeDiameter: 40, transform: { rotation: 90 } }); pointNear(rotated.nodes[0].center, baseline.nodes[1].center);
  const mirror = resolveCycleLayout(c, { frame, nodeDiameter: 40, transform: { flipH: true } }); pointNear(mirror.nodes[1].center, baseline.nodes[3].center);
  const angle = resolveCycleLayout(c, { frame, nodeDiameter: 40, startAngle: 270 + 360 * 1000 }); assert.deepEqual(angle, baseline);
  const both = resolveCycleLayout(c, { frame, nodeDiameter: 40, transform: { flipH: true, flipV: true } }); pointNear(both.nodes[0].center, baseline.nodes[2].center);
  assert.equal(resolveCycleLayout(input(2), { frame, transform: { flipH: true } }).edges[0].clockwise, false);
  assert.deepEqual(c, before);
});

test("measured bounds report overflow explicitly and strict policy rejects without mutation", () => {
  const c = input(3), before = structuredClone(c);
  const fit = resolveCycleLayout(c, { frame, measureText: request => ({ x: 0, y: 0, width: request.availableWidth, height: request.availableHeight }) });
  assert.ok(fit.nodes.every(n => n.textFit === "fits")); assert.deepEqual(fit.findings, []);
  const bearings = resolveCycleLayout(c, { frame, measureText: request => ({ x: -2, y: -1, width: request.availableWidth, height: request.availableHeight }) });
  for (const n of bearings.nodes) { assert.equal(n.textFit, "fits"); near(n.textOrigin.x, n.textBox.x + 2); near(n.textOrigin.y, n.textBox.y + 1); assert.deepEqual(n.placedTextBounds, n.textBox); }
  const bad = resolveCycleLayout(c, { frame, measureText: request => ({ x: -1, y: 0, width: request.availableWidth + 1, height: request.availableHeight + 1 }) });
  assert.deepEqual(bad.findings, c.roots.map(nodeId => ({ nodeId, code: "text-overflow" })));
  assert.throws(() => resolveCycleLayout(c, { frame, overflow: "error", measureText: () => ({ x: 0, y: 0, width: 1000000, height: 1000000 }) }), /text overflows/);
  for (const measureText of [() => ({ width: 0, height: 0 }), () => ({ x: 0, y: 0, width: Infinity, height: 0 }), async () => ({ x: 0, y: 0, width: 0, height: 0 })]) assert.throws(() => resolveCycleLayout(c, { frame, measureText }), TypeError);
  assert.deepEqual(c, before);
  const capturedFrame = { ...frame }, captured = input(3);
  const output = resolveCycleLayout(captured, { frame: capturedFrame, measureText: request => { capturedFrame.width = 1; captured.roots.reverse(); request.paragraphs[0] = "request mutation"; return { x: 0, y: 0, width: 0, height: 0 }; } });
  assert.equal(output.frame.width, frame.width); assert.deepEqual(output.nodes.map(n => n.paragraphs), c.nodes.map(n => n.paragraphs));
});

test("invalid graphs and options reject, geometry/text/work budgets are exact and failures retry", () => {
  for (const mutate of [c => { c.family = "process"; }, c => { c.closure = "explicit-edge"; }, c => { c.roots[1] = c.roots[0]; }, c => { c.nodes[0].id = c.nodes[1].id; }, c => { c.nodes[0].children.push("n1"); }, c => { c.nodes[0].role = "assistant"; }, c => { c.nodes[0].paragraphs = []; }]) {
    const c = input(3); mutate(c); assert.throws(() => resolveCycleLayout(c, { frame }), TypeError);
  }
  for (const options of [{ direction: "backwards" }, { minDiameter: 0 }, { gapRatio: 0 }, { frame: { ...frame, width: 0 } }, { transform: { flipH: 1 } }, { layoutIdentifier: "cycle-custom" }, { unknown: true }, { limits: { maxWork: NaN } }]) assert.throws(() => resolveCycleLayout(input(3), { frame, ...options }), TypeError);
  const c = input(4), before = structuredClone(c);
  for (const options of [{ minDiameter: 100 }, { nodeDiameter: 100 }, { limits: { maxWork: 10 } }, { limits: { maxTextBytes: 1 } }]) { assert.throws(() => resolveCycleLayout(c, { frame, ...options }), RangeError); assert.deepEqual(c, before); }
  assert.throws(() => resolveCycleLayout(c, { frame, limits: { maxNodes: 3 } }));
  const r = resolveCycleLayout(c, { frame }); assert.deepEqual(resolveCycleLayout(c, { frame, limits: { maxWork: r.work } }), r);
  assert.throws(() => resolveCycleLayout(c, { frame, limits: { maxWork: r.work - 1 } }), /maxWork/);
  r.nodes[0].center.x = 999; r.nodes[0].paragraphs[0] = "changed"; assert.deepEqual(c, before);
});

test("seeded diagrams preserve separation, closed topology and transformed frame containment", () => {
  let seed = 4071; const rand = () => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed >>> 8; };
  for (let round = 0; round < 120; round++) {
    const count = 2 + rand() % 30, f = { x: rand() % 40 - 20, y: rand() % 40 - 20, width: 100 + rand() % 900, height: 100 + rand() % 700 };
    const transform = { rotation: rand() % 360, flipH: !!(rand() % 2), flipV: !!(rand() % 2) }, direction = rand() % 2 ? "clockwise" : "counterclockwise";
    const gapRatio = 0.1 + (rand() % 390) / 100;
    const r = resolveCycleLayout(input(count), { frame: f, minDiameter: 0.001, gapRatio, startAngle: rand() % 360, direction, transform });
    const angle = -transform.rotation * Math.PI / 180, origin = { x: f.x + f.width / 2, y: f.y + f.height / 2 };
    const inverse = p => { const x = p.x - origin.x, y = p.y - origin.y; return { x: origin.x + (x * Math.cos(angle) - y * Math.sin(angle)) * (transform.flipH ? -1 : 1), y: origin.y + (x * Math.sin(angle) + y * Math.cos(angle)) * (transform.flipV ? -1 : 1) }; };
    const inside = (p, radius = 0) => { const q = inverse(p); assert.ok(q.x - radius >= f.x - 1e-7 && q.x + radius <= f.x + f.width + 1e-7); assert.ok(q.y - radius >= f.y - 1e-7 && q.y + radius <= f.y + f.height + 1e-7); };
    for (const n of r.nodes) inside(n.center, n.diameter / 2);
    for (let i = 0; i < count; i++) for (let j = i + 1; j < count; j++) assert.ok(distance(r.nodes[i].center, r.nodes[j].center) >= r.diameter * (1 + gapRatio) - 1e-7);
    assert.deepEqual(r.edges.map(e => [e.sourceId, e.targetId]), r.nodes.map((n, i) => [n.id, r.nodes[(i + 1) % count].id]));
    for (const edge of r.edges) {
      for (const p of edge.polygon ?? edge.head) inside(p);
      if (edge.kind === "arrow") for (const node of r.nodes) for (let i = 0; i < edge.polygon.length; i++) assert.ok(segmentDistance(node.center, edge.polygon[i], edge.polygon[(i + 1) % edge.polygon.length]) >= r.diameter / 2 - 1e-7);
      else {
        inside(edge.start); inside(edge.end);
        const a = Math.atan2(edge.start.y - edge.center.y, edge.start.x - edge.center.x), b = Math.atan2(edge.end.y - edge.center.y, edge.end.x - edge.center.x), sign = edge.clockwise ? 1 : -1;
        const delta = (sign * (b - a) + Math.PI * 2) % (Math.PI * 2); assert.ok(delta < Math.PI);
        for (let i = 0; i <= 40; i++) { const angle = a + sign * delta * i / 40, p = { x: edge.center.x + edge.radius * Math.cos(angle), y: edge.center.y + edge.radius * Math.sin(angle) }; inside(p); for (const n of r.nodes) assert.ok(distance(p, n.center) >= r.diameter / 2 - 1e-7); }
        for (const n of r.nodes) for (let i = 0; i < edge.head.length; i++) assert.ok(segmentDistance(n.center, edge.head[i], edge.head[(i + 1) % edge.head.length]) >= r.diameter / 2 - 1e-7);
      }
    }
  }
});
