// Pure, private source-to-managed identity research. Does not authorize source
// XML writes, adopt cached geometry, or import full formatting/opaque semantics.
import { projectDiagramContentBudgeted } from "./diagram-content-model.mjs";
import { createDiagramIdentityState, reconcileDiagramIdentities } from "./diagram-identity-state.mjs";

const encoder = new TextEncoder();
function record(value, names) {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new TypeError("Expected a plain source-binding record.");
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) if (!names.includes(key) || !("value" in descriptor)) throw new TypeError(`Unknown source-binding field/accessor ${key}.`);
}
function identity(raw) {
  if (typeof raw !== "string" || raw.length > 1024) throw new TypeError("Source model identity is missing or over the lexical budget.");
  const token = raw.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
  if (/^[+-]?\d+$/.test(token)) {
    const number = Number(token);
    if (Number.isSafeInteger(number) && number >= -2147483648 && number <= 2147483647) return `int:${number}`;
  }
  if (/^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/.test(token)) return `guid:${token}`;
  throw new TypeError("Unrecognized source model identity.");
}
const canonical = token => {
  if (typeof token !== "string" || !/^(int:|guid:)/.test(token) || identity(token.slice(token.indexOf(":") + 1)) !== token) throw new TypeError("Invalid canonical source identity.");
  return token;
};
const sourceKey = token => JSON.stringify(["source-point", token]);

export function bindDiagramSourceContent(model, layoutIdentifier, options) {
  record(options, ["scope", "namespace", "state", "restore", "limits"]);
  if (!Array.isArray(model.findings) || model.findings.length) throw new TypeError("Source binding requires a complete recognized graph without source findings.");
  if (typeof options.scope !== "string" || !options.scope.length || options.scope.length > 256) throw new TypeError("An explicit logical source scope is required.");
  if (typeof model.part !== "string" || !model.part.startsWith("/") || model.part.length > 2048) throw new TypeError("Expected a bounded source part identity.");
  const limits = { maxNodes: 10000, maxEdges: 20000, maxDepth: 256, maxWork: 600000, maxKeyBytes: 8000000, maxEntries: 200000, ...options.limits };
  if (options.limits !== undefined) record(options.limits, ["maxNodes", "maxEdges", "maxDepth", "maxWork", "maxKeyBytes", "maxEntries"]);
  for (const [key, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value < 1) throw new TypeError(`Invalid source-binding ${key}.`);
  const projected = projectDiagramContentBudgeted(model, layoutIdentifier, Object.fromEntries(["maxNodes", "maxEdges", "maxDepth", "maxWork"].map(key => [key, limits[key]])));
  const projection = projected.projection, family = projection.candidate.family;
  let work = projected.work, keyBytes = 0;
  const step = () => { if (++work > limits.maxWork) throw new RangeError("Source binding exceeds maxWork."); };
  const charge = value => { step(); keyBytes += encoder.encode(value).byteLength; if (keyBytes > limits.maxKeyBytes) throw new RangeError("Source binding exceeds maxKeyBytes."); };
  charge(options.scope); charge(model.part); charge(layoutIdentifier);
  const points = new Map();
  for (const point of model.points) { step(); points.set(point.elementIndex, point); }
  const root = points.get(projection.source.virtualRootElementIndex), rootIdentity = identity(root.modelId); charge(root.modelId);
  const active = [], rawByOrdinal = new Map(), seen = new Set([rootIdentity]);
  for (const binding of projection.bindings) {
    step(); const raw = points.get(binding.elementIndex).modelId, token = identity(raw);
    if (seen.has(token)) throw new TypeError("Duplicate canonical content/document source identity.");
    seen.add(token); charge(raw); charge(token); active.push(token); rawByOrdinal.set(binding.id, { raw, token });
  }
  let previous;
  if (options.state !== undefined) {
    previous = options.state; record(previous, ["version", "scope", "part", "layoutIdentifier", "rootIdentity", "activeSourceIds", "identities"]);
    if (previous.version !== 1 || previous.scope !== options.scope || previous.part !== model.part || previous.layoutIdentifier !== layoutIdentifier || previous.rootIdentity !== rootIdentity) throw new TypeError("Source-binding context/root mismatch; rebinding must be explicit.");
    if (options.namespace !== undefined && options.namespace !== previous.identities.namespace) throw new TypeError("Source-binding namespace mismatch.");
  } else {
    previous = { identities: createDiagramIdentityState(family, options.namespace), activeSourceIds: [] };
  }
  if (!Array.isArray(previous.activeSourceIds) || !Array.isArray(previous.identities?.entries)) throw new TypeError("Invalid source-binding state collections.");
  if (previous.activeSourceIds.length > limits.maxNodes || previous.identities.entries.length > limits.maxEntries) throw new RangeError("Source-binding state exceeds collection budgets.");
  const reserved = new Set(), oldActive = new Set();
  for (const entry of previous.identities.entries) {
    step(); if (typeof entry.key !== "string" || entry.key.length > 4096) throw new TypeError("Invalid source-binding reservation.");
    charge(entry.key); const tuple = JSON.parse(entry.key);
    if (!Array.isArray(tuple) || tuple.length !== 2 || tuple[0] !== "source-point" || sourceKey(canonical(tuple[1])) !== entry.key) throw new TypeError("Foreign source-binding reservation key.");
    reserved.add(tuple[1]);
  }
  for (const token of previous.activeSourceIds) {
    step(); canonical(token); charge(token);
    if (!reserved.has(token) || token === rootIdentity || oldActive.has(token)) throw new TypeError("Invalid/duplicate active source reservation.");
    oldActive.add(token);
  }
  if (options.state !== undefined && !reserved.has(rootIdentity)) throw new TypeError("Source root reservation is missing.");
  const restored = new Set(), requested = options.restore ?? [];
  if (!Array.isArray(requested)) throw new TypeError("Invalid source restore list.");
  if (requested.length > limits.maxNodes) throw new RangeError("Source restore list exceeds maxNodes.");
  for (const raw of requested) {
    step(); const token = identity(raw); charge(raw); charge(token);
    if (restored.has(token) || !seen.has(token) || token === rootIdentity || !reserved.has(token) || oldActive.has(token)) throw new TypeError("Restore must name each actually returning retired source point exactly once.");
    restored.add(token);
  }
  for (const token of active) { step(); if (reserved.has(token) && !oldActive.has(token) && !restored.has(token)) throw new TypeError("Retired source identity reappeared; explicit restoration is required."); }
  step();
  const reconciled = reconcileDiagramIdentities(previous.identities, family, [rootIdentity, ...active].map(sourceKey), {
    maxEntries: limits.maxEntries, maxKeyBytes: Math.max(1, limits.maxKeyBytes - keyBytes), maxWork: Math.max(1, limits.maxWork - work),
  });
  work += reconciled.work;
  const managed = new Map(active.map((token, index) => [token, reconciled.identities[index + 1].modelId])), remap = new Map();
  for (const [ordinal, source] of rawByOrdinal) { step(); remap.set(ordinal, managed.get(source.token)); }
  const nodes = projection.candidate.nodes.map(node => { step(); return { ...node, id: remap.get(node.id), children: node.children.map(id => { step(); return remap.get(id); }) }; });
  const bindings = projection.bindings.map(binding => { step(); const source = rawByOrdinal.get(binding.id); return { ...binding, id: remap.get(binding.id), sourceModelId: source.raw, canonicalModelId: source.token }; });
  const roots = projection.candidate.roots.map(id => { step(); return remap.get(id); });
  return { ...projection, candidate: { ...projection.candidate, roots, nodes }, bindings,
    state: { version: 1, scope: options.scope, part: model.part, layoutIdentifier, rootIdentity, activeSourceIds: active,
      identities: reconciled.state }, work };
}
