import assert from "node:assert/strict";
import test from "node:test";
import { createDiagramIdentityState, reconcileDiagramIdentities } from "./research/diagram-identity-state.mjs";
const namespace = "12345678-1234-4ABC-8DEF";
const start = () => createDiagramIdentityState("cycle", namespace);
const next = (state, keys, options) => reconcileDiagramIdentities(state, "cycle", keys, options);
const id = serial => `{${namespace}-${serial.toString(16).toUpperCase().padStart(12, "0")}}`;

test("identity reservations survive reorder, insertion, removal, restoration and JSON persistence", () => {
  const initial = start(), before = structuredClone(initial);
  const a = next(initial, ["document", "A", "B"]);
  assert.deepEqual(initial, before);
  assert.deepEqual(a.identities.map(x => x.modelId), [id(1), id(2), id(3)]);
  const b = next(JSON.parse(JSON.stringify(a.state)), ["B", "C", "document", "A"]);
  assert.deepEqual(b.identities.map(x => x.modelId), [id(3), id(4), id(1), id(2)]);
  const c = next(b.state, ["document", "C"]);
  assert.deepEqual(c.state, b.state, "removal keeps reservations; it does not recycle identities");
  const d = next(c.state, ["D", "A", "B"]);
  assert.deepEqual(d.identities.map(x => x.modelId), [id(5), id(2), id(3)]);
  d.state.entries[0].key = "mutated";
  d.identities[0].modelId = "mutated";
  assert.deepEqual(c.state, b.state, "every result and mapping is detached");
});

test("invalid snapshots, duplicate active keys, foreign namespaces and families reject atomically", () => {
  const seed = next(start(), ["A", "B"]).state;
  const cases = [s => { s.version = 2; }, s => { s.namespace = namespace.toLowerCase(); }, s => { s.nextSerial = 2; },
    s => { s.entries[1].modelId = s.entries[0].modelId; }, s => { s.entries[1].key = "A"; },
    s => { s.entries.reverse(); }, s => { s.entries[0].modelId = "{FFFFFFFF-1234-4ABC-8DEF-000000000001}"; },
    s => { s.extra = true; }, s => { s.entries[0].extra = true; }, s => { s.entries[0].modelId = id(0); }];
  for (const mutate of cases) { const s = structuredClone(seed); mutate(s); const before = structuredClone(s); assert.throws(() => next(s, ["A"])); assert.deepEqual(s, before); }
  assert.throws(() => reconcileDiagramIdentities(seed, "hierarchy", ["A"]), /family/);
  assert.throws(() => next(seed, ["A", "A"]), /Duplicate active/);
  for (const key of ["", null, 4, "a".repeat(4097)]) assert.throws(() => next(seed, [key]));
  assert.throws(() => createDiagramIdentityState("unknown", namespace));
  assert.throws(() => createDiagramIdentityState("cycle", "12345678-1234-1ABC-8DEF"));
  const accessor = {}; Object.defineProperty(accessor, "version", { get() { throw new Error("getter ran"); } });
  assert.throws(() => next(accessor, []), /accessor/);
});

test("retained and active keys share bounded work, key bytes and reservation capacity; failures retry", () => {
  const s = next(start(), ["A", "😀"]).state, before = structuredClone(s);
  for (const options of [{ maxEntries: 1 }, { maxEntries: 2 }, { maxKeyBytes: 5 }, { maxWork: 4 }]) {
    assert.throws(() => next(s, ["B"], options), RangeError); assert.deepEqual(s, before);
  }
  for (const options of [{ unknown: 1 }, { maxWork: 0 }, { maxEntries: NaN }]) assert.throws(() => next(s, [], options), TypeError);
  assert.deepEqual(next(s, ["B"]).identities, [{ key: "B", modelId: id(3) }]);
  const a = next(s, ["A"]), bytes = new TextEncoder().encode("A😀A").byteLength;
  assert.deepEqual(next(s, ["A"], { maxKeyBytes: bytes, maxWork: a.work }), a);
  assert.throws(() => next(s, ["A"], { maxKeyBytes: bytes - 1 }), /maxKeyBytes/);
  assert.throws(() => next(s, ["A"], { maxWork: a.work - 1 }), /maxWork/);
});

test("serial exhaustion preserves reservations and permits lookup at the final counter", () => {
  const s = start(); s.nextSerial = 0xffffffffffff;
  const final = next(s, ["last"]);
  assert.equal(final.identities[0].modelId, id(0xffffffffffff));
  assert.equal(final.state.nextSerial, 0x1000000000000);
  assert.equal(next(final.state, ["last"]).identities[0].modelId, id(0xffffffffffff));
  assert.throws(() => next(final.state, ["new"]), /exhausted/);
  assert.equal(s.entries.length, 0);
  assert.throws(() => next({ ...s, nextSerial: Number.MAX_SAFE_INTEGER }, []), /nextSerial/);
});

test("seeded edit sequences match an independent never-reuse assignment oracle", () => {
  let randomState = 37457, state = start(), serial = 1;
  const random = () => (randomState = (Math.imul(randomState, 1664525) + 1013904223) >>> 0);
  const expected = new Map();
  for (let round = 0; round < 150; round++) {
    const active = new Set(); for (let i = 0, count = random() % 40; i < count; i++) active.add(`node-${random() % 100}`);
    const keys = [...active]; if (round % 2) keys.reverse();
    for (const key of keys) if (!expected.has(key)) expected.set(key, id(serial++));
    const result = next(state, keys); state = JSON.parse(JSON.stringify(result.state));
    assert.deepEqual(result.identities, keys.map(key => ({ key, modelId: expected.get(key) })));
    assert.equal(state.nextSerial, serial); assert.equal(new Set(state.entries.map(e => e.modelId)).size, expected.size);
  }
});
