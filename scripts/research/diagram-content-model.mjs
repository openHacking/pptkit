// Research-only ordered content forest. Not imported by packages or published.
// Native presentation/cache nodes remain source evidence rather than authoring nodes.
import { diagramResearchProfiles } from "./diagram-layout-profiles.mjs";
const profiles = new Map(diagramResearchProfiles.map(profile => [profile.layoutIdentifier, profile.family]));
export class DiagramContentResearchError extends Error {
  constructor(code, message) { super(message); this.name = "DiagramContentResearchError"; this.code = code; }
}
const fail = (code, message) => { throw new DiagramContentResearchError(code, message); };

export function projectDiagramContent(model, layoutIdentifier, options = {}) {
  return projectDiagramContentBudgeted(model, layoutIdentifier, options).projection;
}

// Separate accounting entry point for composed research adapters. The original
// projection shape stays unchanged; consumers must share, not reset, its budget.
export function projectDiagramContentBudgeted(model, layoutIdentifier, options = {}) {
  const limits = { maxNodes: 10000, maxEdges: 20000, maxDepth: 256, maxWork: 200000, ...options };
  for (const [key, value] of Object.entries(limits)) {
    if (!["maxNodes", "maxEdges", "maxDepth", "maxWork"].includes(key) || !Number.isSafeInteger(value) || value < 1) fail("invalid-limit", `Invalid ${key}.`);
  }
  const family = profiles.get(layoutIdentifier);
  if (!family) fail("unknown-layout", "Only three exact research layout identifiers are admitted.");
  if (model.status !== "inspected") fail("incomplete-source", "A partial/source-only graph is not repaired into a content model.");
  let work = 0;
  const step = () => { if (++work > limits.maxWork) fail("work-limit", "Content projection exceeds maxWork."); };
  const points = new Map(), children = new Map(), parents = new Map(), orders = new Map();
  let documentRoot, ignoredPoints = 0, ignoredConnections = 0, edges = 0;
  const allIndices = new Set();
  for (const point of model.points) {
    step();
    if (!Number.isSafeInteger(point.elementIndex) || point.elementIndex < 1 || allIndices.has(point.elementIndex)) fail("source-identity", "Raw point indices must be distinct positive integers.");
    allIndices.add(point.elementIndex);
    if (["pres", "parTrans", "sibTrans"].includes(point.kind)) { ignoredPoints++; continue; }
    if (!["doc", "node", "asst"].includes(point.kind)) fail("unknown-point", "Unknown point kind could carry content.");
    if (point.identityStatus !== "unique") fail("source-identity", "Content point needs a unique source identity.");
    if (points.size >= limits.maxNodes) fail("node-limit", "Content nodes exceed maxNodes (including the document root).");
    if (point.kind === "asst" && family !== "hierarchy") fail("unsupported-assistant", "Assistant policy is admitted only for the hierarchy research profile.");
    points.set(point.elementIndex, point); children.set(point.elementIndex, []); orders.set(point.elementIndex, new Set());
    if (point.kind === "doc") {
      if (documentRoot !== undefined) fail("document-root", "Exactly one virtual document root is required.");
      documentRoot = point.elementIndex;
    }
  }
  if (documentRoot === undefined) fail("document-root", "Exactly one virtual document root is required.");
  for (const paragraph of points.get(documentRoot).text ?? []) for (const run of paragraph.runs) {
    step();
    if (run.kind !== "text" || run.text !== "") fail("document-text", "Nonempty/opaque document-root text is not silently omitted from this candidate.");
  }
  for (const connection of model.connections) {
    step();
    if (connection.kind !== "parOf") {
      if (!["presOf", "presParOf"].includes(connection.kind)) fail("unknown-connection", "Unknown connection kind could carry content topology.");
      ignoredConnections++; continue;
    }
    if (++edges > limits.maxEdges) fail("edge-limit", "Content edges exceed maxEdges.");
    if (connection.identityStatus !== "unique" || connection.source.status !== "resolved" || connection.destination.status !== "resolved") fail("source-reference", "Content edges need unique identities and resolved point endpoints.");
    const from = connection.source.targetElementIndex, to = connection.destination.targetElementIndex, order = connection.sourceOrderValue;
    if (!points.has(from) || !points.has(to)) fail("unsupported-edge", "Parent edges may not borrow presentation/transition points as content.");
    if (to === documentRoot) fail("document-parent", "The virtual document root cannot have a parent.");
    if (parents.has(to)) fail("multiple-parents", "Repeated edges and multiple parents are not silently merged.");
    if (!Number.isSafeInteger(order) || order < 0 || order > 4294967295) fail("sibling-order", "Sibling order must be a recognized uint32 source value.");
    if (orders.get(from).has(order)) fail("sibling-order", "Duplicate sibling order is ambiguous; XML order is not a tiebreaker.");
    if (family === "cycle" && from !== documentRoot) fail("nested-cycle", "The cycle2 research subset requires flat document children; it never invents a closing parent edge.");
    orders.get(from).add(order); parents.set(to, from); children.get(from).push({ to, order });
  }
  for (const siblings of children.values()) { step(); siblings.sort((a, b) => a.order - b.order); }
  const visited = new Set(), traversal = [], stack = [{ index: documentRoot, depth: 0 }];
  while (stack.length) {
    step(); const { index, depth } = stack.pop();
    if (depth > limits.maxDepth) fail("depth-limit", "Content forest exceeds maxDepth.");
    if (visited.has(index)) fail("cycle", "Content topology must be an ordered forest below the virtual root.");
    visited.add(index); if (index !== documentRoot) traversal.push(index);
    const siblings = children.get(index);
    for (let i = siblings.length - 1; i >= 0; i--) stack.push({ index: siblings[i].to, depth: depth + 1 });
  }
  if (visited.size !== points.size) fail("disconnected", "Disconnected nodes/cycles are not attached or discarded.");
  const ids = new Map(traversal.map((index, i) => [index, `n${i + 1}`]));
  const nodes = traversal.map(index => {
    step(); const point = points.get(index), paragraphs = [];
    for (const paragraph of point.text ?? []) {
      step(); let text = "";
      for (const run of paragraph.runs) {
        step();
        if (run.kind === "field") fail("unsupported-field", "Stored field text is not a semantic editable field in this candidate.");
        if (!["text", "break"].includes(run.kind)) fail("unknown-text", "Unknown text records are not flattened.");
        text += run.kind === "break" ? "\n" : run.text;
      }
      paragraphs.push(text);
    }
    return { id: ids.get(index), role: point.kind === "asst" ? "assistant" : "content", paragraphs, children: children.get(index).map(child => ids.get(child.to)) };
  });
  // IDs identify this projection only. Bindings own source identity, and callers
  // must not persist ordinal IDs as stable authoring identities across edits.
  return { work, projection: { candidate: { family, roots: children.get(documentRoot).map(child => ids.get(child.to)), nodes,
    ...(family === "cycle" ? { closure: "layout-generated" } : {}) },
    bindings: traversal.map(index => ({ id: ids.get(index), part: model.part, elementIndex: index })),
    source: { part: model.part, layoutIdentifier, virtualRootElementIndex: documentRoot, ignoredPoints, ignoredConnections },
    textCoverage: "direct-plain-paragraphs-only" } };
}
