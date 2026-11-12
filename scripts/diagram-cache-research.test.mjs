import assert from "node:assert/strict";
import test from "node:test";
import { createDiagramCacheSession } from "./research/diagram-cache-session.mjs";
const hash = digit => digit.repeat(64);
const input = () => ({ ownerId: "slide-A/diagram-A", layoutIdentifier: "urn:microsoft.com/office/officeart/2005/8/layout/cycle2", coverage: "complete", generatorVersion: "owned-preview-v1", environmentKey: "pinned-fonts/config-1", frame: { width: 520, height: 300 }, dependencies: ["data", "layout", "colors", "style", "theme"].map((role, i) => ({ role, id: `/owned/${role}.xml`, sha256: hash(String(i)) })) });
const cacheBytes = () => new Uint8Array([10, 20, 30]);
const request = extra => ({ policy: "rebuild", generate: () => cacheBytes(), validate: ({ bytes }) => bytes.join(",") === "10,20,30", ...extra });
const deferred = () => { let resolve; const promise = new Promise(r => { resolve = r; }); return { promise, resolve }; };

test("source caches remain missing/empty/unknown and explicit policies do not invent freshness", async () => {
  for (const [sourceCache, state] of [[undefined, "missing"], [new Uint8Array(), "empty"], [cacheBytes(), "unknown"]]) {
    const session = createDiagramCacheSession(input(), { sourceCache });
    assert.equal(session.inspect().state, state);
    assert.deepEqual(await session.prepare({ policy: "block" }), { status: "blocked", reason: `cache-${state}`, revision: 0 });
    const omit = await session.prepare({ policy: "invalidate" }); assert.equal(omit.action, "omit"); assert.equal(session.read(omit), undefined); assert.equal(session.inspect().state, state);
    session.commit(omit); assert.equal(session.inspect().state, "missing"); session.dispose();
  }
  const partial = input(); partial.coverage = "partial"; partial.dependencies = partial.dependencies.filter(d => d.role !== "theme");
  const session = createDiagramCacheSession(partial); assert.deepEqual(await session.prepare(request()), { status: "blocked", reason: "dependency-coverage-partial", revision: 0 });
});

test("managed cache rebuild is independently validated and staged before commit; bytes are detached", async () => {
  const source = cacheBytes(), snapshot = input(), session = createDiagramCacheSession(snapshot, { sourceCache: source }); source[0] = 99; snapshot.ownerId = "changed";
  const built = cacheBytes(); let observed;
  const plan = await session.prepare(request({ generate: ({ snapshot }) => { snapshot.dependencies[0].sha256 = hash("a"); return built; }, validate: ({ bytes, snapshot }) => { observed = snapshot; return bytes[0] === 10 && snapshot.dependencies.find(d => d.role === "data").sha256 === hash("0"); } }));
  built[0] = 88; assert.equal(session.inspect().state, "unknown"); assert.equal(observed.ownerId, "slide-A/diagram-A");
  const read = session.read(plan); read[0] = 77; assert.deepEqual(session.read(plan), cacheBytes());
  session.commit(plan); assert.equal(session.inspect().state, "managed-current");
  assert.throws(() => session.commit(plan), { code: "foreign-plan" });
  const reuse = await session.prepare({ policy: "block" }); assert.equal(reuse.action, "reuse"); assert.deepEqual(session.read(reuse), cacheBytes());
  session.release(reuse); assert.throws(() => session.read(reuse), { code: "foreign-plan" }); session.dispose();
});

test("all captured dependencies, frame, environment and generator changes invalidate; ordering and restore are explicit", async () => {
  const session = createDiagramCacheSession(input()); session.commit(await session.prepare(request()));
  const reordered = input(); reordered.dependencies.reverse(); assert.equal(session.update(reordered), false);
  const mutations = [s => { s.dependencies[0].sha256 = hash("a"); }, s => { s.dependencies[1].sha256 = hash("b"); }, s => { s.dependencies[2].sha256 = hash("c"); }, s => { s.dependencies[3].sha256 = hash("d"); }, s => { s.dependencies[4].sha256 = hash("e"); }, s => { s.dependencies.push({ role: "resource", id: "font-A", sha256: hash("f") }); }, s => { s.frame.width = 600; }, s => { s.frame.rotation = 360; }, s => { s.frame.flipH = true; }, s => { s.environmentKey = "new-font-inventory"; }, s => { s.generatorVersion = "owned-preview-v2"; }, s => { s.coverage = "partial"; }];
  for (const mutate of mutations) {
    const next = input(); mutate(next); assert.equal(session.update(next), true); assert.equal(session.inspect().state, "invalidated");
    assert.equal((await session.prepare({ policy: "block" })).status, "blocked");
    session.update(input()); assert.equal(session.inspect().state, "managed-current", "returning to exact witness dependencies restores managed cache status");
  }
  const seeded = createDiagramCacheSession(input(), { sourceCache: cacheBytes() });
  let randomState = 6419, current = 0, witness = 0, present = true, verified = false;
  const random = () => { randomState = (Math.imul(randomState, 1664525) + 1013904223) >>> 0; return randomState >>> 8; };
  const variant = v => { const s = input(); s.frame.width += v; s.environmentKey = `pinned-fonts/config-${v + 1}`; s.dependencies[0].sha256 = hash(String(v)); return s; };
  const expectedState = () => !present ? "missing" : current !== witness ? "invalidated" : verified ? "managed-current" : "unknown";
  for (let round = 0; round < 120; round++) {
    const next = random() % 8; assert.equal(seeded.update(variant(next)), next !== current); current = next;
    assert.equal(seeded.inspect().state, expectedState());
    const policy = ["block", "invalidate", "rebuild"][random() % 3], fail = random() % 5 === 0, before = seeded.inspect();
    if (policy === "rebuild" && expectedState() !== "managed-current" && fail) { await assert.rejects(() => seeded.prepare(request({ policy, validate: () => false })), { code: "validation-failed" }); assert.deepEqual(seeded.inspect(), before); continue; }
    const plan = await seeded.prepare(request({ policy, validate: () => !fail })); assert.deepEqual(seeded.inspect(), before);
    if (plan.status === "blocked") { assert.equal(policy, "block"); assert.notEqual(expectedState(), "managed-current"); continue; }
    assert.equal(plan.action, expectedState() === "managed-current" ? "reuse" : policy === "invalidate" ? "omit" : "replace");
    if (random() % 4 === 0) { seeded.release(plan); continue; }
    if (random() % 7 === 0) { current = (current + 1) % 8; seeded.update(variant(current)); assert.throws(() => seeded.commit(plan), { code: "stale-plan" }); seeded.release(plan); continue; }
    seeded.commit(plan); present = plan.action !== "omit"; if (present) { witness = current; verified = true; }
    assert.equal(seeded.inspect().state, expectedState());
  }
  seeded.dispose();
  const foreign = input(); foreign.ownerId = "different"; assert.throws(() => session.update(foreign), { code: "context-mismatch" });
  assert.equal(session.inspect().state, "managed-current"); session.dispose();
});

test("stale input, ABA and later cache publication prevent old plan publication across sessions", async () => {
  const session = createDiagramCacheSession(input()), other = createDiagramCacheSession(input());
  const a = await session.prepare(request()), b = await session.prepare(request());
  assert.throws(() => session.commit({ status: "ready", action: "replace", revision: 0 }), { code: "foreign-plan" });
  assert.throws(() => other.commit(a), { code: "foreign-plan" }); session.commit(a);
  assert.throws(() => session.read(b), { code: "stale-plan" }); assert.equal(session.release(b), true);
  const c = await session.prepare({ policy: "block" }); const changed = input(); changed.frame.height = 400; session.update(changed); session.update(input());
  assert.equal(session.inspect().state, "managed-current"); assert.throws(() => session.commit(c), { code: "stale-plan" }); session.release(c);
  const wait = deferred(); let generatedSignal;
  const preparing = session.prepare(request({ policy: "rebuild" })); const reuse = await preparing; session.release(reuse);
  session.update(changed);
  const pending = session.prepare(request({ generate: ({ signal }) => { generatedSignal = signal; return wait.promise; } }));
  await assert.rejects(() => session.prepare(request()), { code: "busy" }); session.update(input()); assert.equal(generatedSignal.aborted, true);
  wait.resolve(cacheBytes()); await assert.rejects(pending, { code: "stale-plan" }); assert.equal(session.inspect().state, "managed-current");
  session.update(changed); const checking = deferred(), entered = deferred();
  const validating = session.prepare(request({ validate: () => { entered.resolve(); return checking.promise; } }));
  await entered.promise; session.update(input()); checking.resolve(true); await assert.rejects(validating, { code: "stale-plan" });
  assert.equal(session.inspect().state, "managed-current", "input changes during async validation do not publish an obsolete witness");
  session.dispose(); other.dispose();
});

test("generation, validation, mutation and size failures keep previous cache intact and permit retry", async () => {
  const session = createDiagramCacheSession(input(), { sourceCache: cacheBytes(), limits: { maxCacheBytes: 4 } });
  for (const extra of [{ generate: () => { throw new Error("generation failed"); } }, { validate: () => false }, { validate: () => "true" }, { validate: () => { throw new Error("validation failed"); } }, { validate: ({ bytes }) => { bytes[0] = 0; bytes.every = () => true; return true; } }, { generate: () => new Uint8Array() }, { generate: () => new Uint8Array(5) }]) {
    await assert.rejects(() => session.prepare(request(extra))); assert.equal(session.inspect().state, "unknown");
  }
  const disguised = new Uint8Array(5); Object.defineProperty(disguised, "byteLength", { value: 0 });
  await assert.rejects(() => session.prepare(request({ generate: () => disguised })), RangeError);
  session.commit(await session.prepare(request())); assert.equal(session.inspect().state, "managed-current"); session.dispose();
});

test("abort and dispose propagate to callbacks, reject ready plans and detach external listeners", async () => {
  const controller = new AbortController(), signal = controller.signal, add = signal.addEventListener.bind(signal), remove = signal.removeEventListener.bind(signal); let listeners = 0;
  signal.addEventListener = (...args) => { listeners++; return add(...args); }; signal.removeEventListener = (...args) => { listeners--; return remove(...args); };
  const session = createDiagramCacheSession(input()), wait = deferred(); let child;
  const pending = session.prepare(request({ signal, generate: ({ signal }) => { child = signal; return wait.promise; } })); assert.equal(listeners, 1);
  session.dispose(); assert.equal(child.aborted, true); assert.equal(listeners, 0); wait.resolve(cacheBytes()); await assert.rejects(pending, { code: "disposed" }); assert.equal(listeners, 0);
  const success = createDiagramCacheSession(input()); const completed = await success.prepare(request({ signal })); assert.equal(listeners, 0, "successful preparation detaches its external abort relay"); success.release(completed); success.dispose();
  const next = createDiagramCacheSession(input()), abort = new AbortController(); const ready = await next.prepare(request({ signal: abort.signal })); abort.abort(); assert.throws(() => next.commit(ready), { code: "aborted" }); next.release(ready);
  await assert.rejects(() => next.prepare(request({ signal: abort.signal })), { code: "aborted" });
  const held = deferred(), external = new AbortController(); let child2;
  const p = next.prepare(request({ signal: external.signal, generate: ({ signal }) => { child2 = signal; return held.promise; } })); external.abort(); assert.equal(child2.aborted, true); held.resolve(cacheBytes()); await assert.rejects(p, { code: "aborted" });
  next.commit(await next.prepare(request())); assert.equal(next.inspect().state, "managed-current"); next.dispose();
});

test("prepared plan quota, exact payload/snapshot budgets and malformed dependency profiles reject atomically", async () => {
  const session = createDiagramCacheSession(input(), { limits: { maxPreparedPlans: 1, maxCacheBytes: 3 } });
  const plan = await session.prepare(request()); await assert.rejects(() => session.prepare(request()), { code: "plan-limit" }); assert.equal(session.release(plan), true); assert.equal(session.release(plan), false);
  session.commit(await session.prepare(request())); session.dispose();
  for (const mutate of [s => { s.dependencies.pop(); }, s => { s.dependencies[0].sha256 = "invalid"; }, s => { s.dependencies.push({ ...s.dependencies[0] }); }, s => { s.frame.width = 0; }, s => { s.frame.flipV = 1; }, s => { s.coverage = "unknown"; }, s => { s.layoutIdentifier = "custom"; }, s => { s.extra = true; }]) { const s = input(); mutate(s); assert.throws(() => createDiagramCacheSession(s)); }
  assert.throws(() => createDiagramCacheSession(input(), { limits: { maxDependencies: 4 } }), RangeError);
  assert.throws(() => createDiagramCacheSession(input(), { limits: { maxSnapshotBytes: 1 } }), RangeError);
  assert.throws(() => createDiagramCacheSession(input(), { limits: { maxCacheBytes: 2 }, sourceCache: cacheBytes() }), RangeError);
  const s = createDiagramCacheSession(input()); const baseline = s.inspect(), invalid = input(); invalid.dependencies[0].sha256 = "bad"; assert.throws(() => s.update(invalid)); assert.deepEqual(s.inspect(), baseline); s.dispose();
  let canonicalSize; const measure = createDiagramCacheSession(input()); const measured = await measure.prepare(request({ generate: ({ snapshot }) => { canonicalSize = new TextEncoder().encode(JSON.stringify(snapshot)).byteLength; return cacheBytes(); } })); measure.release(measured); measure.dispose();
  createDiagramCacheSession(input(), { limits: { maxSnapshotBytes: canonicalSize } }).dispose();
  assert.throws(() => createDiagramCacheSession(input(), { limits: { maxSnapshotBytes: canonicalSize - 1 } }), /maxSnapshotBytes/);
  const escaped = input(); escaped.dependencies.push({ role: "resource", id: "\u0000".repeat(2000), sha256: hash("a") });
  assert.throws(() => createDiagramCacheSession(escaped, { limits: { maxSnapshotBytes: canonicalSize + 10000 } }), /maxSnapshotBytes/);
});
