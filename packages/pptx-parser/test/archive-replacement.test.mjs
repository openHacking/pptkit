import assert from "node:assert/strict";
import test from "node:test";
import { inflateRawSync, deflateRawSync } from "node:zlib";
import { replaceArchiveEntries } from "../dist/preservation/replace-archive-entries.js";
import { inspectZipArchive } from "../dist/index.js";
import { crc32, zip } from "./zip-fixture.mjs";

const code = expected => error => { assert.equal(error.code, expected); return true; };
const replace = (input, entries, options = {}) => replaceArchiveEntries(input, entries, { maxOutputBytes: 1024 * 1024, ...options });
// Independent central/local decoder, supporting comments and both descriptors.
function records(input) {
  const b = Buffer.from(input); let end = b.length - 22;
  while (end >= 0 && !(b.readUInt32LE(end) === 0x06054b50 && end + 22 + b.readUInt16LE(end + 20) === b.length)) end--;
  assert.ok(end >= 0);
  let cursor = b.readUInt32LE(end + 16); const rows = [];
  for (let i = 0; i < b.readUInt16LE(end + 10); i++) {
    assert.equal(b.readUInt32LE(cursor), 0x02014b50);
    const n = b.readUInt16LE(cursor + 28), x = b.readUInt16LE(cursor + 30), c = b.readUInt16LE(cursor + 32);
    const name = b.subarray(cursor + 46, cursor + 46 + n).toString(), local = b.readUInt32LE(cursor + 42), method = b.readUInt16LE(cursor + 10);
    assert.equal(b.readUInt32LE(local), 0x04034b50);
    const start = local + 30 + b.readUInt16LE(local + 26) + b.readUInt16LE(local + 28);
    const dataEnd = start + b.readUInt32LE(cursor + 20), compressed = b.subarray(start, dataEnd);
    const payload = method === 8 ? inflateRawSync(compressed) : Buffer.from(compressed);
    assert.equal(payload.length, b.readUInt32LE(cursor + 24)); assert.equal(crc32(payload), b.readUInt32LE(cursor + 16));
    const flags = b.readUInt16LE(cursor + 8), descriptor = flags & 8 ? b.readUInt32LE(dataEnd) === 0x08074b50 ? 16 : 12 : 0;
    rows.push({ name, local, header: b.subarray(local, start), record: b.subarray(local, dataEnd + descriptor), central: b.subarray(cursor, cursor + 46 + n + x + c), payload });
    cursor += 46 + n + x + c;
  }
  assert.equal(cursor, end);
  return { rows, tail: b.subarray(end), comment: b.subarray(end + 22) };
}
function decoratedZip() {
  const name = Buffer.from("字.xml"), payload = Buffer.from("<r>old</r>"), compressed = deflateRawSync(payload), crc = crc32(payload);
  const extra = Buffer.from([0xfe, 0xca, 2, 0, 4, 5]), comment = Buffer.from("entry comment");
  const local = Buffer.alloc(30 + name.length + extra.length); local.writeUInt32LE(0x04034b50); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x802, 6); local.writeUInt16LE(8, 8);
  local.writeUInt16LE(0x5432, 10); local.writeUInt16LE(0x4321, 12); local.writeUInt32LE(crc, 14); local.writeUInt32LE(compressed.length, 18); local.writeUInt32LE(payload.length, 22);
  local.writeUInt16LE(name.length, 26); local.writeUInt16LE(extra.length, 28); name.copy(local, 30); extra.copy(local, 30 + name.length);
  const central = Buffer.alloc(46 + name.length + extra.length + comment.length); central.writeUInt32LE(0x02014b50); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
  local.copy(central, 8, 6, 26); central.writeUInt16LE(name.length, 28); central.writeUInt16LE(extra.length, 30); central.writeUInt16LE(comment.length, 32);
  central.writeUInt32LE(0x20, 38); name.copy(central, 46); extra.copy(central, 46 + name.length); comment.copy(central, 46 + name.length + extra.length);
  const endComment = Buffer.from("archive comment"), end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50); end.writeUInt16LE(1, 8); end.writeUInt16LE(1, 10);
  end.writeUInt32LE(central.length, 12); end.writeUInt32LE(local.length + compressed.length, 16); end.writeUInt16LE(endComment.length, 20);
  return Buffer.concat([local, compressed, central, end, endComment]);
}

test("ZIP replacement retains untouched compressed/descriptor records and reversed central ordering", async () => {
  const source = zip([
    { name: "folder/", method: 0 },
    { name: "a.xml", body: "<a>old</a>", descriptor: "signed" },
    { name: "shared.bin", body: Buffer.from([1, 4, 255]), descriptor: "unsigned" },
    { name: "b.xml", body: "<b/>", method: 0 },
  ], { reverseCentral: true, comment: "keep metadata" }).bytes;
  const changed = await replace(source, [{ name: "a.xml", bytes: Buffer.from("<a>new 長文字</a>") }]), before = records(source), after = records(changed);
  assert.deepEqual(after.rows.map(r => r.name), before.rows.map(r => r.name));
  assert.deepEqual(after.comment, before.comment);
  for (const original of before.rows) {
    const candidate = after.rows.find(r => r.name === original.name);
    if (original.name === "a.xml") {
      assert.equal(candidate.payload.toString(), "<a>new 長文字</a>");
      assert.equal(candidate.header.readUInt16LE(6) & 0x0e, 0);
      assert.equal(candidate.header.readUInt16LE(8), 0);
    } else {
      assert.deepEqual(candidate.record, original.record);
      const central = Buffer.from(candidate.central); central.writeUInt32LE(original.local, 42);
      assert.deepEqual(central, original.central);
    }
  }
  assert.equal((await inspectZipArchive(changed)).entries.length, 4);
});

test("changed entries retain UTF-8 filenames, extra fields, DOS times, attributes and both comments", async () => {
  const original = decoratedZip(), candidate = await replace(original, [{ name: "字.xml", bytes: Buffer.from("<r>new</r>") }]);
  const a = records(original), b = records(candidate);
  assert.deepEqual(a.comment, b.comment);
  const header = Buffer.from(b.rows[0].header), central = Buffer.from(b.rows[0].central);
  for (const [offset, size] of [[6, 2], [8, 2], [14, 4], [18, 4], [22, 4]]) a.rows[0].header.copy(header, offset, offset, offset + size);
  for (const [offset, size] of [[8, 2], [10, 2], [16, 4], [20, 4], [24, 4], [42, 4]]) a.rows[0].central.copy(central, offset, offset, offset + size);
  assert.deepEqual(header, a.rows[0].header); assert.deepEqual(central, a.rows[0].central);
  assert.equal(b.rows[0].payload.toString(), "<r>new</r>");
});

test("no replacements retain the exact ZIP bytes, including empty ZIP and metadata", async () => {
  for (const input of [zip([]).bytes, decoratedZip(), zip([{ name: "a", body: "x", descriptor: "signed" }], { reverseCentral: true, comment: "comment" }).bytes]) {
    assert.deepEqual(Buffer.from(await replace(input, [])), input);
  }
});

test("zero-length and simultaneous payload replacements recompute sizes/CRCs and preserve other bytes", async () => {
  const input = zip([{ name: "a", body: "old" }, { name: "b", body: "old" }, { name: "c", body: "keep" }]).bytes;
  const candidate = await replace(input, [{ name: "b", bytes: Buffer.from("next") }, { name: "a", bytes: new Uint8Array() }]);
  assert.deepEqual(records(candidate).rows.map(r => [r.name, r.payload.toString()]), [["a", ""], ["b", "next"], ["c", "keep"]]);
});

test("ZIP additions retain existing compressed records, comments and metadata alongside replacements", async () => {
  const source=zip([{name:'a.xml',body:'before',descriptor:'signed'},{name:'kept.bin',body:'kept',descriptor:'unsigned'}],{reverseCentral:true,comment:'keep archive comment'}).bytes;
  const before=records(source),candidate=await replace(source,[{name:'a.xml',bytes:Buffer.from('after')}],{additions:[{name:'ppt/media/new.png',bytes:Buffer.from([0,2,255])},{name:'ppt/media/empty.png',bytes:new Uint8Array()}]});
  const after=records(candidate);assert.deepEqual(after.comment,before.comment);assert.deepEqual(after.rows.map(r=>r.name),[...before.rows.map(r=>r.name),'ppt/media/new.png','ppt/media/empty.png']);
  assert.deepEqual(after.rows.find(r=>r.name==='kept.bin').record,before.rows.find(r=>r.name==='kept.bin').record);
  assert.deepEqual(after.rows.find(r=>r.name==='ppt/media/new.png').payload,Buffer.from([0,2,255]));assert.equal(after.rows.find(r=>r.name==='ppt/media/empty.png').payload.length,0);
  assert.equal(after.rows.find(r=>r.name==='a.xml').payload.toString(),'after');assert.equal((await inspectZipArchive(candidate)).entries.length,4);
  const empty=await replace(zip([]).bytes,[],{additions:[{name:'first',bytes:Buffer.from('x')}]});assert.equal(records(empty).rows[0].payload.toString(),'x');
});

test("ZIP additions reject aliases, unsafe names and budgets without publishing partial output", async () => {
  const source=zip([{name:'media/A.png',body:'old'},{name:'parent',body:'file'}]).bytes,original=Buffer.from(source);
  for(const name of ['../bad','/bad','a//b','a/./b','a/../b','bad?query','bad%20name','trailing/','字'])await assert.rejects(replace(source,[],{additions:[{name,bytes:Buffer.from('x')}]}),code('archive-path-invalid'));
  await assert.rejects(replace(source,[],{additions:[{name:'MEDIA/a.PNG',bytes:Buffer.from('x')}]}),code('archive-duplicate-path'));
  await assert.rejects(replace(source,[],{additions:[{name:'new',bytes:Buffer.from('x')},{name:'NEW',bytes:Buffer.from('y')}]}),code('archive-duplicate-path'));
  await assert.rejects(replace(source,[],{additions:[{name:'parent/child',bytes:Buffer.from('x')}]}),code('archive-path-invalid'));
  const encoded=zip([{name:'media/%41.png',body:'old'},{name:'parent',body:'x'},{name:'parent-',body:'x'}]).bytes;
  await assert.rejects(replace(encoded,[],{additions:[{name:'media/a.png',bytes:Buffer.from('x')}],maxOutputBytes:2}),code('archive-duplicate-path'));
  await assert.rejects(replace(encoded,[],{additions:[{name:'parent/child',bytes:Buffer.from('x')}],maxOutputBytes:2}),code('archive-path-invalid'));
  await assert.rejects(replace(source,[],{additions:[{name:'new',bytes:Buffer.from('x')},{name:'new-',bytes:Buffer.from('x')},{name:'new/child',bytes:Buffer.from('x')}],maxOutputBytes:3}),code('archive-path-invalid'));
  await assert.rejects(replace(source,[],{additions:[{name:'new',bytes:Buffer.alloc(5)}],limits:{maxEntryBytes:4}}),code('archive-limit-exceeded'));
  await assert.rejects(replace(source,[],{additions:[{name:'new',bytes:Buffer.alloc(5)}],limits:{maxExpandedBytes:11}}),code('archive-limit-exceeded'));
  await assert.rejects(replace(source,[],{additions:[{name:'new',bytes:Buffer.from('x')}],limits:{maxEntries:2}}),code('archive-limit-exceeded'));
  await assert.rejects(replace(source,[],{additions:[{name:'new',bytes:Buffer.from('x')}],maxOutputBytes:source.length}),code('archive-limit-exceeded'));
  assert.deepEqual(source,original);
});

test("ZIP additions capture actual views and records once before asynchronous work", async () => {
  class Hostile extends Uint8Array{get byteLength(){return 0}get buffer(){return new ArrayBuffer(0)}*[Symbol.iterator](){throw Error('iterator')}}
  const source=zip([{name:'old',body:'old'}]).bytes,raw=new Hostile(Buffer.from('payload'));let names=0,values=0;
  const record={get name(){names++;return 'added'},get bytes(){values++;return raw}};
  const pending=replace(source,[],{additions:[record]});raw.fill(0);const candidate=await pending;
  assert.equal(names,1);assert.equal(values,1);assert.equal(records(candidate).rows.find(r=>r.name==='added').payload.toString(),'payload');
  if(typeof SharedArrayBuffer!=='undefined')await assert.rejects(replace(source,[],{additions:[{name:'shared',bytes:new Hostile(new SharedArrayBuffer(5))}]}),/ordinary byte buffers/);
  const controller=new AbortController();controller.abort();await assert.rejects(replace(source,[],{additions:[{name:'new',bytes:Buffer.from('x')}],signal:controller.signal}),code('operation-aborted'));
});

test("missing, duplicate and directory targets reject without mutating the source", async () => {
  const input = zip([{ name: "a", body: "old" }, { name: "dir/", method: 0 }]).bytes, original = Buffer.from(input);
  for (const entries of [
    [{ name: "missing", bytes: Buffer.from("x") }],
    [{ name: "a", bytes: Buffer.from("x") }, { name: "a", bytes: Buffer.from("y") }],
    [{ name: "dir/", bytes: new Uint8Array() }],
  ]) await assert.rejects(replace(input, entries), code("archive-invalid"));
  assert.deepEqual(input, original);
  assert.equal(records(await replace(input, [{ name: "a", bytes: Buffer.from("retry") }])).rows[0].payload.toString(), "retry");
});

test("source validation cannot be bypassed by replacing a corrupt payload", async () => {
  const source = zip([{ name: "a", body: "broken", method: 0, crc: 1 }]).bytes;
  await assert.rejects(replace(source, [{ name: "a", bytes: Buffer.from("fixed") }]), code("archive-crc-mismatch"));
  await assert.rejects(replace(Buffer.from("bad"), []), code("archive-invalid"));
});

test("exact output/entry/expanded/input budgets fail before returning any candidate", async () => {
  const source = zip([{ name: "a", body: "old", method: 0 }]).bytes, entry = [{ name: "a", bytes: Buffer.from("newer") }];
  const expected = source.length + 2;
  assert.equal((await replace(source, entry, { maxOutputBytes: expected })).length, expected);
  await assert.rejects(replace(source, entry, { maxOutputBytes: expected - 1 }), code("archive-limit-exceeded"));
  await assert.rejects(replace(source, entry, { limits: { maxEntryBytes: 4 } }), code("archive-limit-exceeded"));
  await assert.rejects(replace(source, entry, { limits: { maxExpandedBytes: 4 } }), code("archive-limit-exceeded"));
  await assert.rejects(replace(source, entry, { limits: { maxInputBytes: source.length - 1 } }), code("archive-limit-exceeded"));
  await assert.rejects(replace(source, entry, { maxOutputBytes: 0 }), RangeError);
  await assert.rejects(replace(source, entry, { limits: { maxEntryBytes: NaN } }), RangeError);
});

test("source/replacement/options capture, cancellation and independent retries do not share output state", async () => {
  const input = zip([{ name: "a", body: "old" }]).bytes, replacement = Buffer.from("new"), entries = [{ name: "a", bytes: replacement }], options = { maxOutputBytes: 10000 };
  const pending = replace(input, entries, options);
  input.fill(0); replacement.fill(0); entries[0].name = "missing"; options.maxOutputBytes = 1;
  const output = await pending;
  assert.equal(records(output).rows[0].payload.toString(), "new");
  const source = zip([{ name: "a", body: "x" }]).bytes, controller = new AbortController();
  const editing = replace(source, [{ name: "a", bytes: Buffer.alloc(3 * 1024 * 1024, 3) }], { maxOutputBytes: 4 * 1024 * 1024, signal: controller.signal });
  setTimeout(() => controller.abort(), 0);
  await assert.rejects(editing, code("operation-aborted"));
  assert.equal(records(await replace(source, [{ name: "a", bytes: Buffer.from("retry") }])).rows[0].payload.toString(), "retry");
  const cancelled = new AbortController(); cancelled.abort();
  await assert.rejects(replace(source, [], { signal: cancelled.signal }), code("operation-aborted"));
});
