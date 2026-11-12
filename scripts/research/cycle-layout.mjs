// Pure feasibility experiment for equal-circle Cycle content. Not a DiagramML
// evaluator, native cache generator or admitted public Layout API.
const encoder = new TextEncoder();
const LAYOUT = "urn:microsoft.com/office/officeart/2005/8/layout/cycle2";
const TAU = Math.PI * 2, EPS = 1e-10;
const fail = message => { throw new TypeError(message); };
function plain(value, keys) {
  if (!value || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) fail("Expected a plain Cycle layout record.");
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) if (!keys.includes(key) || !("value" in descriptor)) fail(`Unknown Cycle layout field/accessor ${key}.`);
}
const finite = (value, min, max, name) => { if (!Number.isFinite(value) || value < min || value > max) fail(`Invalid ${name}.`); return value; };
const point = (angle, radius) => ({ x: radius * Math.cos(angle), y: radius * Math.sin(angle) });
const clean = value => Math.abs(value) < 1e-12 ? 0 : value;
const degrees = angle => clean(((angle % 360) + 360) % 360);

export function resolveCycleLayout(candidate, options = {}) {
  plain(candidate, ["family", "roots", "nodes", "closure"]);
  plain(options, ["layoutIdentifier", "frame", "startAngle", "direction", "gapRatio", "minDiameter", "maxDiameter", "nodeDiameter", "transform", "measureText", "overflow", "limits"]);
  if (candidate.family !== "cycle" || candidate.closure !== "layout-generated" || (options.layoutIdentifier ?? LAYOUT) !== LAYOUT) fail("Only the exact Cycle research profile is admitted.");
  const frame = options.frame; plain(frame, ["x", "y", "width", "height"]);
  for (const key of ["x", "y", "width", "height"]) finite(frame[key], key === "x" || key === "y" ? -1e6 : 1e-6, 1e6, `frame ${key}`);
  const capturedFrame = { ...frame };
  const transform = { rotation: 0, flipH: false, flipV: false, ...options.transform };
  if (options.transform !== undefined) plain(options.transform, ["rotation", "flipH", "flipV"]);
  finite(transform.rotation, -1e9, 1e9, "rotation");
  if (typeof transform.flipH !== "boolean" || typeof transform.flipV !== "boolean") fail("Invalid frame flip.");
  const start = degrees(finite(options.startAngle ?? -90, -1e9, 1e9, "startAngle")) * Math.PI / 180;
  const direction = options.direction ?? "clockwise"; if (!["clockwise", "counterclockwise"].includes(direction)) fail("Invalid cycle direction.");
  const sign = direction === "clockwise" ? 1 : -1;
  const gap = finite(options.gapRatio ?? 0.5, 0.1, 4, "gapRatio");
  const minDiameter = finite(options.minDiameter ?? 8, 1e-6, 1e6, "minDiameter"), maxDiameter = finite(options.maxDiameter ?? 1e6, minDiameter, 1e6, "maxDiameter");
  if (options.nodeDiameter !== undefined) finite(options.nodeDiameter, minDiameter, maxDiameter, "nodeDiameter");
  const measureText = options.measureText, overflow = options.overflow ?? "report";
  if (measureText !== undefined && typeof measureText !== "function") fail("measureText must be an explicit synchronous callback.");
  if (!["report", "error"].includes(overflow)) fail("Invalid text overflow policy.");
  const limits = { maxNodes: 10000, maxWork: 300000, maxTextBytes: 1000000, ...options.limits };
  if (options.limits !== undefined) plain(options.limits, ["maxNodes", "maxWork", "maxTextBytes"]);
  for (const [key, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value < 1) fail(`Invalid ${key}.`);
  if (!Array.isArray(candidate.nodes) || !Array.isArray(candidate.roots) || candidate.nodes.length > limits.maxNodes || candidate.roots.length !== candidate.nodes.length) fail("Cycle needs a bounded flat ordered node collection.");
  let work = 0, textBytes = 0;
  const step = () => { if (++work > limits.maxWork) throw new RangeError("Cycle layout exceeds maxWork."); };
  const byId = new Map();
  for (const node of candidate.nodes) {
    step(); plain(node, ["id", "role", "children", "paragraphs"]);
    if (typeof node.id !== "string" || !node.id.length || node.id.length > 256 || byId.has(node.id)) fail("Invalid/duplicate Cycle node identity.");
    if (node.role !== "content" || !Array.isArray(node.children) || node.children.length || !Array.isArray(node.paragraphs) || !node.paragraphs.length) fail("Cycle layout requires plain content leaves with paragraphs.");
    const paragraphs = node.paragraphs.map(text => { step(); if (typeof text !== "string") fail("Cycle paragraphs must be strings."); textBytes += encoder.encode(text).byteLength; if (textBytes > limits.maxTextBytes) throw new RangeError("Cycle layout exceeds maxTextBytes."); return text; });
    byId.set(node.id, { id: node.id, paragraphs });
  }
  const seen = new Set(), ordered = candidate.roots.map(id => { step(); if (!byId.has(id) || seen.has(id)) fail("Missing/duplicate Cycle root."); seen.add(id); return byId.get(id); });
  const n = ordered.length, radius = n > 1 ? (1 + gap) / (2 * Math.sin(Math.PI / n)) : 0;
  const unitNodes = ordered.map((node, i) => { step(); return { ...node, angle: start + sign * TAU * i / n, ...point(start + sign * TAU * i / n, radius) }; });
  const unitEdges = [];
  const extent = { minX: Infinity, minY: Infinity, maxX: -Infinity, maxY: -Infinity };
  const include = p => { step(); extent.minX = Math.min(extent.minX, p.x); extent.minY = Math.min(extent.minY, p.y); extent.maxX = Math.max(extent.maxX, p.x); extent.maxY = Math.max(extent.maxY, p.y); };
  for (const node of unitNodes) { include({ x: node.x - 0.5, y: node.y - 0.5 }); include({ x: node.x + 0.5, y: node.y + 0.5 }); }
  if (n > 1) for (let i = 0; i < n; i++) {
    step(); const source = unitNodes[i], target = unitNodes[(i + 1) % n];
    if (n === 2) {
      // Distinct directed half-orbits avoid the coincident reversed chords of a
      // two-node cycle. Nodes clip the circular track analytically.
      const offset = 2 * Math.asin(0.5 / (2 * radius)), a = source.angle + sign * offset, b = source.angle + sign * Math.PI - sign * offset;
      const tip = point(b, radius), tangent = { x: -sign * Math.sin(b), y: sign * Math.cos(b) }, normal = { x: Math.cos(b), y: Math.sin(b) };
      const head = [tip, { x: tip.x - 0.15 * tangent.x + 0.09 * normal.x, y: tip.y - 0.15 * tangent.y + 0.09 * normal.y }, { x: tip.x - 0.15 * tangent.x - 0.09 * normal.x, y: tip.y - 0.15 * tangent.y - 0.09 * normal.y }];
      const samples = [point(a, radius), point(b, radius)];
      for (let k = -8; k <= 8; k++) { step(); const angle = k * Math.PI / 2; if (sign * (angle - a) >= -EPS && sign * (b - angle) >= -EPS) samples.push(point(angle, radius)); }
      for (const p of [...head, ...samples]) include(p);
      unitEdges.push({ sourceId: source.id, targetId: target.id, kind: "arc", start: a, end: b, head });
    } else {
      const angle = Math.atan2(target.y - source.y, target.x - source.x), c = Math.cos(angle), s = Math.sin(angle), cx = (source.x + target.x) / 2, cy = (source.y + target.y) / 2;
      const length = gap / 2, halfHeight = 0.3375 / 2;
      const polygon = [[-length / 2, -halfHeight / 2], [0, -halfHeight / 2], [0, -halfHeight], [length / 2, 0], [0, halfHeight], [0, halfHeight / 2], [-length / 2, halfHeight / 2]].map(([x, y]) => { step(); return { x: cx + x * c - y * s, y: cy + x * s + y * c }; });
      polygon.forEach(include); unitEdges.push({ sourceId: source.id, targetId: target.id, kind: "arrow", polygon });
    }
  }
  const diameter = n ? options.nodeDiameter ?? Math.min(maxDiameter, capturedFrame.width / (extent.maxX - extent.minX), capturedFrame.height / (extent.maxY - extent.minY)) : 0;
  if (n && (diameter < minDiameter - EPS || diameter * (extent.maxX - extent.minX) > capturedFrame.width + EPS || diameter * (extent.maxY - extent.minY) > capturedFrame.height + EPS)) throw new RangeError("Cycle layout cannot fit the requested minimum/explicit node diameter.");
  const origin = { x: capturedFrame.x + capturedFrame.width / 2, y: capturedFrame.y + capturedFrame.height / 2 };
  const unitCenter = n ? { x: (extent.minX + extent.maxX) / 2, y: (extent.minY + extent.maxY) / 2 } : { x: 0, y: 0 };
  const rotation = degrees(transform.rotation) * Math.PI / 180, cos = Math.cos(rotation), sin = Math.sin(rotation);
  const project = p => {
    step(); const x = (p.x - unitCenter.x) * diameter * (transform.flipH ? -1 : 1), y = (p.y - unitCenter.y) * diameter * (transform.flipV ? -1 : 1);
    return { x: clean(origin.x + x * cos - y * sin), y: clean(origin.y + x * sin + y * cos) };
  };
  const orbitCenter = project({ x: 0, y: 0 }), findings = [];
  const nodes = unitNodes.map(node => {
    step(); const center = project(node), side = diameter / Math.SQRT2;
    // Rectangle corners are inside the node circle; no glyph measurement is
    // inferred. The explicit callback reports bounds relative to this rectangle.
    const textBox = { x: center.x - side * 0.4, y: center.y - side * 0.4, width: side * 0.8, height: side * 0.8 };
    let textFit = "unmeasured", measuredBounds, textOrigin, placedTextBounds;
    if (measureText) {
      step(); const measured = measureText({ id: node.id, paragraphs: [...node.paragraphs], availableWidth: textBox.width, availableHeight: textBox.height });
      plain(measured, ["x", "y", "width", "height"]);
      for (const key of ["x", "y", "width", "height"]) finite(measured[key], key === "width" || key === "height" ? 0 : -1e6, 1e6, `measured ${key}`);
      measuredBounds = { ...measured };
      // Center the measured ink/content bounds, including negative bearings;
      // retain both the provider's bounds and the actual placement transform.
      textOrigin = { x: textBox.x + (textBox.width - measured.width) / 2 - measured.x, y: textBox.y + (textBox.height - measured.height) / 2 - measured.y };
      placedTextBounds = { x: textOrigin.x + measured.x, y: textOrigin.y + measured.y, width: measured.width, height: measured.height };
      textFit = measured.width <= textBox.width + EPS && measured.height <= textBox.height + EPS ? "fits" : "overflow";
      if (textFit === "overflow" && overflow === "error") throw new RangeError(`Cycle text overflows node ${node.id}.`);
    }
    if (textFit !== "fits") findings.push({ nodeId: node.id, code: textFit === "unmeasured" ? "text-unmeasured" : "text-overflow" });
    return { id: node.id, center, diameter, box: { x: center.x - diameter / 2, y: center.y - diameter / 2, width: diameter, height: diameter }, paragraphs: [...node.paragraphs], textBox, textFit, ...(measuredBounds ? { measuredBounds, textOrigin, placedTextBounds } : {}) };
  });
  const edges = unitEdges.map(edge => {
    step(); return edge.kind === "arrow" ? { sourceId: edge.sourceId, targetId: edge.targetId, kind: edge.kind, polygon: edge.polygon.map(project) }
      : { sourceId: edge.sourceId, targetId: edge.targetId, kind: edge.kind, center: { ...orbitCenter }, radius: radius * diameter,
        start: project(point(edge.start, radius)), end: project(point(edge.end, radius)), clockwise: (sign === 1) !== (transform.flipH !== transform.flipV), head: edge.head.map(project) };
  });
  return { profile: "equal-circle-cycle-layout-research-v1", layoutIdentifier: LAYOUT, coordinateSpace: "page-points", frame: capturedFrame, transform: { ...transform, rotation: degrees(transform.rotation) },
    direction, startAngle: degrees(options.startAngle ?? -90), diameter, orbit: { center: orbitCenter, radius: radius * diameter }, nodes, edges, findings, work,
    nativeAdmission: "untested", limitations: ["Owned equal-circle packing, not a complete DiagramML cycle algorithm or native cache writer.", "Two-node curved edges are an explicit research policy, not a verified native layout.", "Text requires an explicit measurement provider; geometric fit does not certify glyph coverage or native text fitting."] };
}
