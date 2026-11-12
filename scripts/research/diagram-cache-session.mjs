// Private cache lifecycle research. Current means validated for explicit captured
// dependencies, never native compatibility or source-preservation write authority.
import { diagramResearchProfiles } from "./diagram-layout-profiles.mjs";
const layouts = new Set(diagramResearchProfiles.map(p => p.layoutIdentifier));
const roles = new Set(["data", "layout", "colors", "style", "theme", "resource"]);
const encoder = new TextEncoder();
const byteLength = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(Uint8Array.prototype), "byteLength").get;
const everyByte = Uint8Array.prototype.every;
export class DiagramCacheResearchError extends Error {
  constructor(code, message) { super(message); this.name = "DiagramCacheResearchError"; this.code = code; }
}
const error = (code, message) => { throw new DiagramCacheResearchError(code, message); };
function plain(value, keys) {
  if (!value || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new TypeError("Expected a plain cache research record.");
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) if (!keys.includes(key) || !("value" in descriptor)) throw new TypeError(`Unknown cache field/accessor ${key}.`);
}
function text(value, max, name) { if (typeof value !== "string" || !value.length || value.length > max) throw new TypeError(`Invalid ${name}.`); return value; }

export function createDiagramCacheSession(initial, options = {}) {
  plain(options, ["sourceCache", "limits"]);
  const limits = { maxDependencies: 4096, maxSnapshotBytes: 1000000, maxCacheBytes: 8000000, maxPreparedPlans: 4, ...options.limits };
  if (options.limits !== undefined) plain(options.limits, ["maxDependencies", "maxSnapshotBytes", "maxCacheBytes", "maxPreparedPlans"]);
  for (const [key, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value < 1) throw new TypeError(`Invalid ${key}.`);
  const capture = value => {
    plain(value, ["ownerId", "layoutIdentifier", "coverage", "generatorVersion", "environmentKey", "frame", "dependencies"]);
    text(value.ownerId, 256, "cache owner"); text(value.generatorVersion, 256, "generator version"); text(value.environmentKey, 1024, "environment key");
    if (!layouts.has(value.layoutIdentifier) || !["complete", "partial"].includes(value.coverage)) throw new TypeError("Unknown cache layout or dependency coverage.");
    plain(value.frame, ["width", "height", "rotation", "flipH", "flipV"]);
    const frame = { rotation: 0, flipH: false, flipV: false, ...value.frame };
    for (const key of ["width", "height"]) if (!Number.isFinite(frame[key]) || frame[key] <= 0 || frame[key] > 1e6) throw new TypeError(`Invalid cache frame ${key}.`);
    if (!Number.isFinite(frame.rotation) || Math.abs(frame.rotation) > 1e9 || typeof frame.flipH !== "boolean" || typeof frame.flipV !== "boolean") throw new TypeError("Invalid cache frame transform.");
    const header = { ownerId: value.ownerId, layoutIdentifier: value.layoutIdentifier, coverage: value.coverage, generatorVersion: value.generatorVersion, environmentKey: value.environmentKey,
      frame: { width: frame.width, height: frame.height, rotation: frame.rotation, flipH: frame.flipH, flipV: frame.flipV }, dependencies: [] };
    let snapshotBytes = encoder.encode(JSON.stringify(header)).byteLength;
    if (snapshotBytes > limits.maxSnapshotBytes) throw new RangeError("Cache dependency snapshot exceeds maxSnapshotBytes.");
    // Dependency equality is exact; equivalent-looking transforms are not silently
    // normalized into a semantic-equivalence claim.
    if (!Array.isArray(value.dependencies) || value.dependencies.length > limits.maxDependencies) throw new RangeError("Cache dependency collection exceeds its profile.");
    const counts = new Map(), ids = new Set();
    const dependencies = value.dependencies.map((dependency, index) => {
      plain(dependency, ["role", "id", "sha256"]); text(dependency.id, 2048, "dependency identity");
      if (!roles.has(dependency.role) || typeof dependency.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(dependency.sha256) || ids.has(dependency.id)) throw new TypeError("Invalid/duplicate cache dependency.");
      ids.add(dependency.id); counts.set(dependency.role, (counts.get(dependency.role) ?? 0) + 1);
      const captured = { role: dependency.role, id: dependency.id, sha256: dependency.sha256 };
      // Charge escaped JSON/UTF-8 incrementally, before concatenating the full
      // snapshot; control characters cannot amplify a small-byte budget.
      snapshotBytes += encoder.encode(JSON.stringify(captured)).byteLength + (index ? 1 : 0);
      if (snapshotBytes > limits.maxSnapshotBytes) throw new RangeError("Cache dependency snapshot exceeds maxSnapshotBytes.");
      return captured;
    });
    for (const role of ["data", "layout", "colors", "style"]) if (counts.get(role) !== 1) throw new TypeError(`Exactly one ${role} dependency is required.`);
    if (value.coverage === "complete" && !counts.has("theme")) throw new TypeError("Complete cache dependencies require explicit theme evidence.");
    dependencies.sort((a, b) => a.role < b.role ? -1 : a.role > b.role ? 1 : a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
    const snapshot = { ...header, dependencies };
    const signature = JSON.stringify(snapshot);
    if (encoder.encode(signature).byteLength > limits.maxSnapshotBytes) throw new RangeError("Cache dependency snapshot exceeds maxSnapshotBytes.");
    return { snapshot, signature };
  };
  const bytes = (value, allowEmpty = false) => {
    if (!(value instanceof Uint8Array)) throw new TypeError("Cache payload must be a Uint8Array.");
    const size = byteLength.call(value);
    if (!allowEmpty && !size) throw new TypeError("Cache payload must be nonempty.");
    if (size > limits.maxCacheBytes) throw new RangeError("Cache payload exceeds maxCacheBytes.");
    return new Uint8Array(value);
  };
  let input = capture(initial), revision = 0, epoch = 0, disposed = false, active, count = 0, plans = new WeakMap();
  let cache = options.sourceCache === undefined ? undefined : { bytes: bytes(options.sourceCache, true), signature: input.signature, verified: false };
  const live = () => { if (disposed) error("disposed", "Cache session is disposed."); };
  const state = () => cache === undefined ? "missing" : cache.bytes.length === 0 ? "empty" : cache.signature !== input.signature ? "invalidated" : cache.verified ? "managed-current" : "unknown";
  const fresh = (capturedRevision, capturedEpoch, signal) => {
    live(); if (signal?.aborted) error("aborted", "Cache operation was aborted.");
    if (capturedRevision !== revision || capturedEpoch !== epoch) error("stale-plan", "Cache input or published cache changed.");
  };
  const reserve = (action, payload, signal, capturedRevision, capturedEpoch) => {
    fresh(capturedRevision, capturedEpoch, signal);
    if (count >= limits.maxPreparedPlans) error("plan-limit", "Release or commit a prepared cache plan before preparing more.");
    const plan = Object.freeze({ status: "ready", action, revision: capturedRevision });
    plans.set(plan, { action, payload, signal, revision: capturedRevision, epoch: capturedEpoch }); count++; return plan;
  };
  const lookup = plan => {
    live(); const held = plans.get(plan); if (!held) error("foreign-plan", "Unknown, released or consumed cache plan.");
    fresh(held.revision, held.epoch, held.signal); return held;
  };
  return {
    inspect() { live(); return { revision, epoch, state: state(), coverage: input.snapshot.coverage, ownerId: input.snapshot.ownerId }; },
    update(next) {
      live(); const captured = capture(next);
      if (captured.snapshot.ownerId !== input.snapshot.ownerId || captured.snapshot.layoutIdentifier !== input.snapshot.layoutIdentifier) error("context-mismatch", "Cache owner/layout rebinding must be explicit.");
      if (captured.signature === input.signature) return false;
      if (revision === Number.MAX_SAFE_INTEGER) error("revision-limit", "Cache revision space exhausted.");
      input = captured; revision++; active?.controller.abort("cache input changed"); return true;
    },
    async prepare(request) {
      live(); plain(request, ["policy", "generate", "validate", "signal"]);
      if (!["block", "invalidate", "rebuild"].includes(request.policy)) throw new TypeError("Unknown cache resolution policy.");
      const signal = request.signal;
      if (signal !== undefined && (typeof signal.addEventListener !== "function" || typeof signal.removeEventListener !== "function" || typeof signal.aborted !== "boolean")) throw new TypeError("Expected an AbortSignal.");
      fresh(revision, epoch, signal);
      if (active) error("busy", "A cache generator is still running.");
      if (count >= limits.maxPreparedPlans) error("plan-limit", "Prepared cache plan budget exhausted.");
      const capturedRevision = revision, capturedEpoch = epoch;
      if (state() === "managed-current") return reserve("reuse", new Uint8Array(cache.bytes), signal, capturedRevision, capturedEpoch);
      if (request.policy === "block") return Object.freeze({ status: "blocked", reason: `cache-${state()}`, revision });
      if (request.policy === "invalidate") return reserve("omit", undefined, signal, capturedRevision, capturedEpoch);
      if (input.snapshot.coverage !== "complete") return Object.freeze({ status: "blocked", reason: "dependency-coverage-partial", revision });
      if (typeof request.generate !== "function" || typeof request.validate !== "function") throw new TypeError("Rebuild requires explicit generation and independent validation callbacks.");
      const generate = request.generate, validate = request.validate, controller = new AbortController(), signature = input.signature;
      const relay = () => controller.abort(signal.reason);
      let detached = false;
      const cleanup = () => { if (!detached) { signal?.removeEventListener("abort", relay); detached = true; } };
      signal?.addEventListener("abort", relay, { once: true });
      active = { controller, cleanup };
      try {
        if (signal?.aborted) controller.abort(signal.reason);
        const captured = JSON.parse(signature);
        const generated = bytes(await generate({ snapshot: captured, signal: controller.signal }));
        fresh(capturedRevision, capturedEpoch, signal);
        const checking = new Uint8Array(generated);
        const valid = await validate({ bytes: checking, snapshot: JSON.parse(signature), signal: controller.signal });
        fresh(capturedRevision, capturedEpoch, signal);
        if (byteLength.call(checking) !== generated.length || !everyByte.call(checking, (value, i) => value === generated[i])) error("validator-mutation", "Cache validator mutated the inspected candidate.");
        if (valid !== true) error("validation-failed", "Cache candidate failed independent validation.");
        return reserve("replace", generated, signal, capturedRevision, capturedEpoch);
      } finally { cleanup(); active = undefined; }
    },
    read(plan) { const held = lookup(plan); return held.payload === undefined ? undefined : new Uint8Array(held.payload); },
    commit(plan) {
      const held = lookup(plan);
      if (epoch === Number.MAX_SAFE_INTEGER) error("revision-limit", "Cache publication space exhausted.");
      cache = held.action === "omit" ? undefined : { bytes: new Uint8Array(held.payload), signature: input.signature, verified: true };
      plans.delete(plan); count--; epoch++; active?.controller.abort("cache published"); return { action: held.action, revision, epoch };
    },
    release(plan) { live(); if (plans.has(plan)) { plans.delete(plan); count--; return true; } return false; },
    dispose() { if (disposed) return; disposed = true; active?.cleanup(); active?.controller.abort("cache session disposed"); active = undefined; cache = undefined; input = undefined; plans = new WeakMap(); count = 0; },
  };
}
