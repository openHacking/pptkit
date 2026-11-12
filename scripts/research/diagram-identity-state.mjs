// Private pure research state. Callers own the namespace and persist the returned
// snapshot only after the complete package transaction succeeds. No global IDs.
const MAX_SERIAL = 0xffffffffffff;
const PREFIX = /^[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}$/;
const families = new Set(["process", "hierarchy", "cycle"]);
const encoder = new TextEncoder();
function record(value, keys) {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new TypeError("Expected a plain identity-state record.");
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    if (!keys.includes(key) || !("value" in descriptor)) throw new TypeError(`Unknown identity-state field/accessor ${key}.`);
  }
}
const identifier = (namespace, serial) => `{${namespace}-${serial.toString(16).toUpperCase().padStart(12, "0")}}`;

export function createDiagramIdentityState(family, namespace) {
  if (!families.has(family) || typeof namespace !== "string" || !PREFIX.test(namespace)) throw new TypeError("Identity state requires an exact family and uppercase UUID v4/variant prefix.");
  return { version: 1, family, namespace, nextSerial: 1, entries: [] };
}

export function reconcileDiagramIdentities(state, family, keys, options = {}) {
  record(options, ["maxEntries", "maxKeyBytes", "maxWork"]);
  const limits = { maxEntries: 200000, maxKeyBytes: 8000000, maxWork: 600000, ...options };
  for (const [name, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value < 1) throw new TypeError(`Invalid identity ${name}.`);
  record(state, ["version", "family", "namespace", "nextSerial", "entries"]);
  createDiagramIdentityState(state.family, state.namespace);
  if (state.version !== 1 || family !== state.family) throw new TypeError("Identity state version/family mismatch.");
  if (!Number.isSafeInteger(state.nextSerial) || state.nextSerial < 1 || state.nextSerial > MAX_SERIAL + 1) throw new TypeError("Invalid identity nextSerial.");
  if (!Array.isArray(state.entries) || !Array.isArray(keys)) throw new TypeError("Identity entries and keys must be arrays.");
  if (state.entries.length > limits.maxEntries || keys.length > limits.maxEntries) throw new RangeError("Identity state exceeds maxEntries.");
  let work = 0, keyBytes = 0;
  const step = () => { if (++work > limits.maxWork) throw new RangeError("Identity state exceeds maxWork."); };
  const checkKey = key => {
    step();
    if (typeof key !== "string" || !key.length || key.length > 4096) throw new TypeError("Invalid identity key.");
    keyBytes += encoder.encode(key).byteLength;
    if (keyBytes > limits.maxKeyBytes) throw new RangeError("Identity state exceeds maxKeyBytes.");
  };
  const assignments = new Map(), entries = [], used = new Set();
  let previousSerial = 0;
  for (const entry of state.entries) {
    step(); record(entry, ["key", "modelId"]); checkKey(entry.key);
    if (assignments.has(entry.key) || typeof entry.modelId !== "string") throw new TypeError("Duplicate/invalid identity assignment.");
    const match = /^\{(.+)-([0-9A-F]{12})\}$/.exec(entry.modelId), serial = match && Number.parseInt(match[2], 16);
    if (!match || match[1] !== state.namespace || serial <= previousSerial || serial >= state.nextSerial || entry.modelId !== identifier(state.namespace, serial)) throw new TypeError("Invalid or nonmonotonic reserved identity.");
    previousSerial = serial; assignments.set(entry.key, entry.modelId); entries.push({ key: entry.key, modelId: entry.modelId });
  }
  // Gaps are valid reservations (e.g. a source import); serials below nextSerial
  // are never reassigned. Deleted keys remain reserved for explicit undo/restore.
  let nextSerial = state.nextSerial;
  const identities = [];
  for (const key of keys) {
    checkKey(key);
    if (used.has(key)) throw new TypeError("Duplicate active identity key.");
    used.add(key);
    let modelId = assignments.get(key);
    if (modelId === undefined) {
      if (entries.length >= limits.maxEntries) throw new RangeError("Identity state exceeds maxEntries.");
      if (nextSerial > MAX_SERIAL) throw new RangeError("Identity serial space exhausted.");
      modelId = identifier(state.namespace, nextSerial++);
      assignments.set(key, modelId); entries.push({ key, modelId });
    }
    identities.push({ key, modelId });
  }
  return { identities, work, state: { version: 1, family, namespace: state.namespace, nextSerial, entries } };
}
