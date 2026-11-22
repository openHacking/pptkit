import { metadataFixture, encodeMetadata, custom } from "./metadata-fixture.mjs";
import assert from "node:assert/strict";
import test from "node:test";
import * as fs from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { importPptxFile, writePptxFile, PptxFileError } from "../dist/node.js";
import { importPptx } from "../dist/index.js";
import { createPptxFileAdapter } from "../dist/node/file-adapter.js";
import { tableFixture, table, tableRow, tableBody, cell } from "./table-fixture.mjs";
import { encodeStructure, structureFixture } from "./structure-fixture.mjs";
import { withSections } from "./sections-fixture.mjs";
import { bytes as contentBytes, shape as contentShape, body as contentBody } from './content-fixture.mjs';

const options = { projection: "content", editSaveProfile: "text-runs-v3" };
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const code = expected => error => { assert.equal(error.code, expected, error.stack); return true; };
const fail = (code = "EIO") => Object.assign(new Error(`injected ${code}`), { code });
const temporary = file => path.basename(file).startsWith(".pptkit-") && file.endsWith(".tmp");
const defer = () => { let resolve; const promise = new Promise(r => { resolve = r; }); return { promise, resolve }; };
function instrument(intercept) {
  const wrap = (handle, file) => new Proxy(handle, { get(target, key) {
    const value = Reflect.get(target, key, target);
    return typeof value === "function" ? (...args) => intercept({ method: key, file, args }, () => value.apply(target, args)) : value;
  } });
  return Object.fromEntries(["open", "lstat", "realpath", "link", "rename", "unlink"].map(method => [method, (...args) => intercept(
    { method, file: String(args[0]), args },
    async () => { const result = await fs[method](...args); return method === "open" ? wrap(result, String(args[0])) : result; },
  )]));
}
async function fixture(t, adapter = { importFile: importPptxFile, writeFile: writePptxFile }, bytes = encodeStructure(tableFixture())) {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(tmpdir(), "pptkit-file-save-")));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, "source.pptx"); await fs.writeFile(source, bytes);
  const imported = await adapter.importFile(source, options); t.after(() => imported.session.dispose());
  const slide = imported.session.document.slides[0], original = slide.elements[0], changed = structuredClone(original);
  changed.rows[0].cells[0].content[0].runs[0].text = "File edit <&🙂";
  slide.removeElement(original.id); slide.insertElement(0, changed);
  const prepared = await imported.session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
  return { ...imported, plan: prepared.plan, source, directory, input: bytes, adapter };
}
const noTemps = async directory => assert.deepEqual((await fs.readdir(directory)).filter(temporary), []);

test("Node file adapter captures the explicit section collection projection", async t => {
  const directory = await fs.mkdtemp(path.join(tmpdir(), "pptkit-section-file-"));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, "source.pptx"), bytes = encodeStructure(withSections(structureFixture()));
  await fs.writeFile(file, bytes);
  const imported = await importPptxFile(file, { projection: "structure", sectionProjection: "collection-v1" });
  t.after(() => imported.session.dispose());
  assert.deepEqual(imported.session.document.sections.map(s => [s.name, s.slideIds]), [["Opening", ["source-slide-300"]], ["Empty", []], ["Closing", ["source-slide-256"]]]);
  const plan = await imported.session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
  assert.deepEqual(Buffer.from((await imported.session.save(plan.plan)).bytes), Buffer.from(bytes));
  const collectionOnly = await importPptxFile(file, { projection: "structure", sectionProjection: "collection-v2" });
  t.after(() => collectionOnly.session.dispose());
  assert.deepEqual(collectionOnly.session.document.sections, imported.session.document.sections);
  assert.deepEqual(collectionOnly.session.document.slides.map(slide => slide.section), [undefined, undefined]);
});

test('Node file adapter captures and uses the explicit existing-description profile', async t => {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(tmpdir(), 'pptkit-description-file-')));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const input = path.join(directory, 'source.pptx'), output = path.join(directory, 'output.pptx'); await fs.writeFile(input, contentBytes());
  const imported = await importPptxFile(input, { projection: 'content', descriptionEditSaveProfile: 'existing-v1' }); t.after(() => imported.session.dispose());
  const slide = imported.session.document.slides[0], element = slide.elements[0];
  slide.replaceElement(element.id, { ...element, accessibility: { ...element.accessibility, description: 'File description' } });
  const plan = await imported.session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
  assert.equal((await writePptxFile(imported.session, plan.plan, output, { mode: 'create' })).status, 'committed');
  const reopened = await importPptx(await fs.readFile(output), { projection: 'content' });
  assert.equal(reopened.document.slides[0].elements[0].accessibility.description, 'File description'); reopened.dispose();
  const removing = await importPptxFile(input, { projection: 'content', descriptionEditSaveProfile: 'attribute-v2' }); t.after(() => removing.session.dispose());
  const removingSlide = removing.session.document.slides[0], current = removingSlide.elements[0], next = structuredClone(current);
  delete next.accessibility.description; removingSlide.replaceElement(current.id, next);
  const removalPlan = await removing.session.prepareSave({ mode: 'edit' }); assert.equal(removalPlan.status, 'ready', JSON.stringify(removalPlan.report));
  const removedOutput = path.join(directory, 'removed.pptx');
  assert.equal((await writePptxFile(removing.session, removalPlan.plan, removedOutput, { mode: 'create' })).status, 'committed');
  const removed = await importPptx(await fs.readFile(removedOutput), { projection: 'content' });
  assert.equal(removed.document.slides[0].elements[0].accessibility?.description, undefined); removed.dispose();
  await noTemps(directory);
});

test('Node file adapter saves a batch of descriptions atomically', async t => {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(tmpdir(), 'pptkit-description-batch-')));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const input = path.join(directory, 'source.pptx'), output = path.join(directory, 'output.pptx');
  await fs.writeFile(input, contentBytes(contentShape({ id: '2' }) + contentShape({ id: '3' })));
  const imported = await importPptxFile(input, { projection: 'content', descriptionEditSaveProfile: 'batch-v3' }); t.after(() => imported.session.dispose());
  const slide = imported.session.document.slides[0];
  for (const [index, element] of [...slide.elements].entries()) slide.replaceElement(element.id, { ...element, accessibility: { ...element.accessibility, description: `Batch ${index}` } });
  const plan = await imported.session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
  assert.equal((await writePptxFile(imported.session, plan.plan, output, { mode: 'create' })).status, 'committed');
  const reopened = await importPptx(await fs.readFile(output), { projection: 'content' });
  assert.deepEqual(reopened.document.slides[0].elements.map(element => element.accessibility.description), ['Batch 0', 'Batch 1']); reopened.dispose();
  await noTemps(directory);
});

test('Node file adapter saves combined text and description edits', async t => {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(tmpdir(), 'pptkit-description-mixed-')));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const input = path.join(directory, 'source.pptx'), output = path.join(directory, 'output.pptx');
  await fs.writeFile(input, contentBytes(contentShape({ id: '2', txBox: true, text: contentBody() })));
  const imported = await importPptxFile(input, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1' }); t.after(() => imported.session.dispose());
  const slide = imported.session.document.slides[0], original = slide.elements[0], next = structuredClone(original);
  next.content[0].runs[0].text = 'File text'; next.accessibility.description = 'File alt'; next.name = 'File name'; slide.replaceElement(original.id, next);
  const plan = await imported.session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
  assert.equal((await writePptxFile(imported.session, plan.plan, output, { mode: 'create' })).status, 'committed');
  const reopened = await importPptx(await fs.readFile(output), { projection: 'content' });
  assert.equal(reopened.document.slides[0].elements[0].content[0].runs[0].text, 'File text');
  assert.equal(reopened.document.slides[0].elements[0].accessibility.description, 'File alt'); assert.equal(reopened.document.slides[0].elements[0].name, 'File name'); reopened.dispose(); await noTemps(directory);
});

test("public Node create/replace round-trips validated table bytes, retains ordinary permissions and issues reusable snapshots", async t => {
  const f = await fixture(t), output = path.join(f.directory, "created.pptx");
  const created = await writePptxFile(f.session, f.plan, output, { mode: "create" });
  assert.equal(created.status, "committed"); assert.equal(created.directorySynced, true); assert.deepEqual(created.warnings, []);
  assert.equal(hash(await fs.readFile(output)), created.file.sha256);
  assert.ok(Object.isFrozen(created) && Object.isFrozen(created.file) && Object.isFrozen(created.warnings));
  assert.throws(() => JSON.stringify(created.file), /not serializable/);
  assert.equal((await fs.stat(output)).mode & 0o777, 0o600);
  const imported = await importPptxFile(output, options);
  assert.equal(imported.session.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0].text, "File edit <&🙂"); imported.session.dispose();
  await fs.chmod(f.source, 0o640);
  const current = await importPptxFile(f.source, options); t.after(() => current.session.dispose());
  const replaced = await writePptxFile(f.session, f.plan, f.source, { mode: "replace", expected: current.file });
  assert.equal((await fs.stat(f.source)).mode & 0o777, 0o640);
  const repeated = await writePptxFile(f.session, f.plan, f.source, { mode: "replace", expected: replaced.file });
  assert.equal(repeated.file.sha256, replaced.file.sha256);
  await assert.rejects(writePptxFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file }), code("file-target-changed"));
  assert.equal((await f.session.prepareSave({ mode: "no-op" })).status, "blocked");
  await noTemps(f.directory);
});

test("synchronous plan guards enforce ownership, consent, stale/disposed state and semantic undo without copying or consuming plans", async t => {
  const f = await fixture(t);
  assert.equal(f.session.assertSavePlanCurrent(f.plan), undefined);
  const other = await importPptx(f.input, options); t.after(() => other.dispose());
  assert.throws(() => other.assertSavePlanCurrent(f.plan), code("plan-session-mismatch"));
  assert.throws(() => f.session.assertSavePlanCurrent(f.plan, { acceptLossIds: ["foreign"] }), code("save-loss-consent-mismatch"));
  const slide = f.session.document.slides[0], saved = structuredClone(slide.elements[0]), changed = structuredClone(saved);
  changed.rows[0].cells[0].content[0].runs[0].text = "stale";
  slide.removeElement(saved.id); slide.insertElement(0, changed);
  assert.throws(() => f.session.assertSavePlanCurrent(f.plan), code("save-snapshot-stale"));
  slide.removeElement(saved.id); slide.insertElement(0, saved);
  assert.equal(f.session.assertSavePlanCurrent(f.plan), undefined);
  assert.equal((await f.session.save(f.plan)).bytes.length, f.plan.byteLength);
  f.plan.dispose(); assert.throws(() => f.session.assertSavePlanCurrent(f.plan), code("plan-disposed"));
});

test("invalid modes/options/foreign or forged snapshots and stale/foreign plans create no output or temp file", async t => {
  const f = await fixture(t), output = path.join(f.directory, "never.pptx");
  for (const option of [{}, { mode: "unknown" }, { mode: "create", expected: f.file }, { mode: "replace" }, { mode: "replace", expected: { ...f.file } }]) {
    await assert.rejects(writePptxFile(f.session, f.plan, output, option));
  }
  await assert.rejects(writePptxFile(f.session, f.plan, output, { mode: "replace", expected: f.file }), code("file-snapshot-invalid"));
  await assert.rejects(writePptxFile(f.session, f.plan, output, { mode: "create", acceptLossIds: ["unknown"] }), code("save-loss-consent-mismatch"));
  await assert.rejects(writePptxFile(f.session, {}, output, { mode: "create" }), code("plan-session-mismatch"));
  await assert.rejects(fs.stat(output), { code: "ENOENT" }); await noTemps(f.directory);
  assert.equal(hash(await fs.readFile(f.source)), hash(f.input));
});

test("exclusive create preserves existing regular files and symlinks without modifying their referents", async t => {
  const f = await fixture(t), link = path.join(f.directory, "link.pptx");
  await fs.symlink(f.source, link);
  for (const target of [f.source, link]) await assert.rejects(writePptxFile(f.session, f.plan, target, { mode: "create" }), code("file-target-exists"));
  await assert.rejects(importPptxFile(link, options), code("file-not-regular"));
  await assert.rejects(importPptxFile(f.directory, options), code("file-not-regular"));
  assert.ok((await fs.lstat(link)).isSymbolicLink()); assert.equal(hash(await fs.readFile(f.source)), hash(f.input)); await noTemps(f.directory);
});

test("file import captures options at invocation and enforces byte limits before opening the source", async t => {
  const entered = defer(), release = defer(); let opens = 0;
  const io = instrument(async (event, perform) => {
    if (event.method === "realpath") { entered.resolve(); await release.promise; }
    if (event.method === "open") opens++;
    return perform();
  });
  const adapter = createPptxFileAdapter(io), directory = await fs.mkdtemp(path.join(tmpdir(), "pptkit-file-options-"));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, "source.pptx"), bytes = encodeStructure(tableFixture()); await fs.writeFile(source, bytes);
  const config = { ...options, zipLimits: { maxInputBytes: bytes.length }, sessionLimits: { maxReportFindings: 1000 } };
  const pending = adapter.importFile(source, config); await entered.promise;
  config.projection = "structure"; config.editSaveProfile = "unknown"; config.zipLimits.maxInputBytes = 1; config.sessionLimits.maxReportFindings = 1;
  release.resolve(); const result = await pending;
  assert.equal(result.session.report.projection, "content"); assert.equal(result.session.report.tableTextRuns.length, 4); result.session.dispose();
  opens = 0; await assert.rejects(adapter.importFile(source, { ...options, zipLimits: { maxInputBytes: bytes.length - 1 } }), code("file-limit-exceeded"));
  assert.equal(opens, 0);
  for (const config of [{ projection: "unknown" }, { ...options, editSaveProfile: "" }, { ...options, sessionLimits: { maxLivePlans: 0 } }]) await assert.rejects(adapter.importFile(source, config));
});

test("source growth, shrinkage, replacement and cancellation during bounded read reject without returning a session", async t => {
  for (const kind of ["grow", "shrink", "replace", "abort"]) {
    let armed = true; const abort = new AbortController();
    const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
      const result = await perform();
      if (event.method === "read" && armed) {
        armed = false;
        if (kind === "grow") await fs.appendFile(event.file, "extra");
        if (kind === "shrink") await fs.truncate(event.file, 1);
        if (kind === "replace") { const swap = event.file + ".new"; await fs.copyFile(event.file, swap); await fs.rename(swap, event.file); }
        if (kind === "abort") abort.abort();
      }
      return result;
    }));
    const directory = await fs.mkdtemp(path.join(tmpdir(), "pptkit-source-race-")); t.after(() => fs.rm(directory, { recursive: true, force: true }));
    const source = path.join(directory, "input.pptx"); await fs.writeFile(source, encodeStructure(tableFixture()));
    await assert.rejects(adapter.importFile(source, { ...options, signal: abort.signal }), code(kind === "abort" ? "operation-aborted" : "file-source-changed"));
  }
});

test("replacement rejects changed contents even when inode/size/mode and exposed mtime still match", async t => {
  const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
    const result = await perform();
    if ((event.method === "lstat" || event.method === "stat") && result && "mtimeNs" in result) result.mtimeNs = 0n;
    return result;
  }));
  const f = await fixture(t, adapter), changed = Buffer.from(f.input); changed[changed.length - 1] ^= 1;
  await fs.writeFile(f.source, changed);
  await assert.rejects(adapter.writeFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file }), code("file-target-changed"));
  assert.equal(hash(await fs.readFile(f.source)), hash(changed)); await noTemps(f.directory);
});

test("replacement rejects equal-byte inode replacement, missing targets and symlink redirection", async t => {
  for (const kind of ["equal-new-inode", "missing", "symlink"]) {
    const f = await fixture(t);
    const backup = f.source + ".backup"; await fs.rename(f.source, backup);
    if (kind === "equal-new-inode") await fs.writeFile(f.source, f.input);
    if (kind === "symlink") await fs.symlink(backup, f.source);
    await assert.rejects(writePptxFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file }), code("file-target-changed"));
    assert.equal(hash(await fs.readFile(backup)), hash(f.input)); await noTemps(f.directory);
  }
});

test("every precommit IO stage failure leaves the original intact, cleans temp files and permits retry", async t => {
  for (const stage of ["open-parent", "open-temp", "write", "sync-temp", "close-temp", "verify-read", "final-parent", "rename"]) {
    let armed = false, fired = false, target, directory;
    const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
      const hit = stage === "open-parent" ? event.method === "open" && event.file === directory
        : stage === "open-temp" ? event.method === "open" && temporary(event.file)
        : stage === "write" ? event.method === "write"
        : stage === "sync-temp" ? event.method === "sync" && temporary(event.file)
        : stage === "close-temp" ? event.method === "close" && temporary(event.file)
        : stage === "verify-read" ? event.method === "read" && temporary(event.file)
        : stage === "final-parent" ? event.method === "lstat" && event.file === directory
        : event.method === "rename" && event.args[1] === target;
      if (armed && !fired && hit) { fired = true; throw fail(); }
      return perform();
    }));
    const f = await fixture(t, adapter); target = f.source; directory = f.directory; armed = true;
    await assert.rejects(adapter.writeFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file }), code("file-io"), stage);
    assert.equal(fired, true, stage); assert.equal(hash(await fs.readFile(f.source)), hash(f.input), stage);
    await noTemps(f.directory); f.session.assertSavePlanCurrent(f.plan);
    const retried = await adapter.writeFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file });
    assert.equal(hash(await fs.readFile(f.source)), retried.file.sha256);
  }
});

test("short writes are completed; zero progress and same-size corruption are rejected before publication", async t => {
  for (const kind of ["short", "zero", "corrupt"]) {
    const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
      if (event.method === "write") {
        if (kind === "zero") return { bytesWritten: 0 };
        if (kind === "short") event.args[2] = Math.min(event.args[2], 97);
        if (kind === "corrupt") { event.args[0] = new Uint8Array(event.args[0]); event.args[0][event.args[1]] ^= 1; }
      }
      return perform();
    }));
    const f = await fixture(t, adapter), output = path.join(f.directory, kind + ".pptx");
    if (kind === "short") assert.equal((await adapter.writeFile(f.session, f.plan, output, { mode: "create" })).status, "committed");
    else { await assert.rejects(adapter.writeFile(f.session, f.plan, output, { mode: "create" }), code("file-io")); await assert.rejects(fs.stat(output), { code: "ENOENT" }); }
    await noTemps(f.directory);
  }
});

test("cancellation before/during temp work always preserves original bytes and releases the path lock", async t => {
  for (const stage of ["before", "write", "sync", "verify-read", "final-temp"]) {
    let armed = false, counts = 0; const abort = new AbortController();
    const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
      const result = await perform();
      if (armed && temporary(event.file)) {
        if (event.method === "lstat") counts++;
        if (event.method === stage || (stage === "verify-read" && event.method === "read") || (stage === "final-temp" && event.method === "lstat" && counts === 3)) abort.abort();
      }
      return result;
    }));
    const f = await fixture(t, adapter); armed = true; if (stage === "before") abort.abort();
    await assert.rejects(adapter.writeFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file, signal: abort.signal }), code("operation-aborted"));
    assert.equal(hash(await fs.readFile(f.source)), hash(f.input)); await noTemps(f.directory);
    armed = false; await adapter.writeFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file });
  }
});

test("edits/session disposal/plan disposal at the final async boundary fail the synchronous commit guard", async t => {
  for (const kind of ["edit", "session", "plan"]) {
    let armed = false, f, counts = 0;
    const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
      const result = await perform();
      if (armed && event.method === "lstat" && temporary(event.file) && ++counts === 3) {
        if (kind === "session") f.session.dispose();
        if (kind === "plan") f.plan.dispose();
        if (kind === "edit") f.session.document.slides[0].setNotes("mutation at commit boundary");
      }
      return result;
    }));
    f = await fixture(t, adapter); armed = true;
    await assert.rejects(adapter.writeFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file }), code(kind === "edit" ? "save-snapshot-stale" : kind === "session" ? "session-disposed" : "plan-disposed"));
    assert.equal(hash(await fs.readFile(f.source)), hash(f.input)); await noTemps(f.directory);
  }
});

test("target changes after initial guard and exclusive-create races preserve the competing writer", async t => {
  for (const kind of ["replace", "create"]) {
    let armed = false, target; const competing = Buffer.from("external writer");
    const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
      if (armed && ((kind === "replace" && event.method === "sync" && temporary(event.file)) || (kind === "create" && event.method === "link"))) { armed = false; await fs.writeFile(target, competing); }
      return perform();
    }));
    const f = await fixture(t, adapter); target = kind === "replace" ? f.source : path.join(f.directory, "raced.pptx"); armed = true;
    await assert.rejects(adapter.writeFile(f.session, f.plan, target, kind === "replace" ? { mode: "replace", expected: f.file } : { mode: "create" }), code(kind === "replace" ? "file-target-changed" : "file-target-exists"));
    assert.equal(hash(await fs.readFile(target)), hash(competing)); await noTemps(f.directory);
  }
});

test("same-adapter path/inode contention fails explicitly and independent adapters still publish new names exclusively", async t => {
  const entered = defer(), release = defer(); let pause = true;
  const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
    if (event.method === "write" && pause) { pause = false; entered.resolve(); await release.promise; }
    return perform();
  }));
  const f = await fixture(t, adapter), output = path.join(f.directory, "one.pptx");
  const first = adapter.writeFile(f.session, f.plan, output, { mode: "create" }); await entered.promise;
  await assert.rejects(adapter.writeFile(f.session, f.plan, output, { mode: "create" }), code("file-write-busy"));
  release.resolve(); await first;
  const ready = defer(); let links = 0;
  const io = instrument(async (event, perform) => { if (event.method === "link") { if (++links === 2) ready.resolve(); await ready.promise; } return perform(); });
  const a = createPptxFileAdapter(io), b = createPptxFileAdapter(io), target = path.join(f.directory, "exclusive.pptx");
  const results = await Promise.allSettled([a.writeFile(f.session, f.plan, target, { mode: "create" }), b.writeFile(f.session, f.plan, target, { mode: "create" })]);
  assert.equal(results.filter(r => r.status === "fulfilled").length, 1);
  assert.equal(results.find(r => r.status === "rejected").reason.code, "file-target-exists"); await noTemps(f.directory);
});

test("post-dispatch cancellation and directory-sync failure return committed outcomes with usable receipts", async t => {
  const abort = new AbortController(); let directory, armed = true;
  const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
    if (event.method === "rename") { const result = await perform(); abort.abort(); return result; }
    if (event.method === "sync" && event.file === directory && armed) throw fail("EINVAL");
    return perform();
  }));
  const f = await fixture(t, adapter); directory = f.directory;
  const result = await adapter.writeFile(f.session, f.plan, f.source, { mode: "replace", expected: f.file, signal: abort.signal });
  assert.equal(result.status, "committed"); assert.equal(result.directorySynced, false);
  assert.equal(result.warnings[0].code, "directory-sync-failed"); assert.equal(result.warnings[0].errorCode, "EINVAL");
  assert.equal(hash(await fs.readFile(f.source)), result.file.sha256);
  armed = false; assert.equal((await adapter.writeFile(f.session, f.plan, f.source, { mode: "replace", expected: result.file })).directorySynced, true);
  await noTemps(f.directory);
});

test("cleanup failures are explicit before commit; after commit they are warnings and do not discard successful output", async t => {
  for (const committed of [false, true]) {
    let armed = true, leftover;
    const adapter = createPptxFileAdapter(instrument(async (event, perform) => {
      if (!committed && event.method === "write" && armed) throw fail();
      if (event.method === "unlink" && temporary(event.file) && armed) { leftover = event.file; throw fail("EACCES"); }
      return perform();
    }));
    const f = await fixture(t, adapter), output = path.join(f.directory, "cleanup.pptx");
    if (committed) {
      const result = await adapter.writeFile(f.session, f.plan, output, { mode: "create" });
      assert.equal(result.status, "committed"); assert.equal(result.warnings[0].code, "temporary-cleanup-failed");
      assert.equal(hash(await fs.readFile(output)), result.file.sha256);
      armed = false; await adapter.writeFile(f.session, f.plan, output, { mode: "replace", expected: result.file });
    } else {
      await assert.rejects(adapter.writeFile(f.session, f.plan, output, { mode: "create" }), error => {
        assert.ok(error instanceof PptxFileError); assert.equal(error.code, "file-cleanup-failed"); assert.equal(error.committed, false); assert.equal(error.temporaryPath, leftover); return true;
      });
      await assert.rejects(fs.stat(output), { code: "ENOENT" }); armed = false;
      await adapter.writeFile(f.session, f.plan, output, { mode: "create" });
    }
    await fs.unlink(leftover); await noTemps(f.directory);
  }
});

test("large original-byte no-op file writes use bounded chunks and preserve every source byte", async t => {
  let state = 0x6d2b79f5;
  const alphabet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  const text = () => Array.from({ length: 1000 }, () => {
    state = Math.imul(state ^ (state >>> 15), 1 | state);
    state ^= state + Math.imul(state ^ (state >>> 7), 61 | state);
    return alphabet[(state ^ (state >>> 14)) & 63];
  }).join("");
  const runs = `<a:p>${Array.from({ length: 200 }, () => `<a:r><a:t>${text()}</a:t></a:r>`).join("")}</a:p>`;
  const bytes = encodeStructure(tableFixture(table({ rows: [tableRow([cell({ body: tableBody(runs) }), cell()], 50)] })));
  const sizes = [], adapter = createPptxFileAdapter(instrument(async (event, perform) => {
    if (event.method === "read" || event.method === "write") sizes.push(event.args[2]);
    return perform();
  }));
  const directory = await fs.mkdtemp(path.join(tmpdir(), "pptkit-large-file-")); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, "source.pptx"), output = path.join(directory, "copy.pptx"); await fs.writeFile(source, bytes);
  const imported = await adapter.importFile(source, options); t.after(() => imported.session.dispose());
  const plan = await imported.session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
  await adapter.writeFile(imported.session, plan.plan, output, { mode: "create" });
  assert.equal(hash(await fs.readFile(output)), hash(bytes)); assert.ok(sizes.length > 6); assert.ok(sizes.every(n => n <= 64 * 1024)); await noTemps(directory);
});

test('Node file adapter retains the drawing name save profile', async t => {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(tmpdir(), 'pptkit-drawing-name-'))); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const input = path.join(directory, 'source.pptx'), output = path.join(directory, 'output.pptx'); await fs.writeFile(input, contentBytes(contentShape({ id: '2' })));
  const imported = await importPptxFile(input, { projection: 'content', drawingNameSaveProfile: 'existing-v1' }); t.after(() => imported.session.dispose());
  const slide = imported.session.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.name = 'File name'; slide.replaceElement(old.id, edit);
  const plan = await imported.session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal((await writePptxFile(imported.session, plan.plan, output, { mode: 'create' })).status, 'committed');
  const reopened = await importPptx(await fs.readFile(output), { projection: 'content' }); try { assert.equal(reopened.document.slides[0].elements[0].name, 'File name'); } finally { reopened.dispose(); } await noTemps(directory);
});

test('Node file adapter retains the decorative source save profile', async t => {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(tmpdir(), 'pptkit-decorative-value-'))); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const input = path.join(directory, 'source.pptx'), output = path.join(directory, 'output.pptx');
  const ext = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>';
  const xml = contentShape({ id: '2' }).replace('<p:cNvPr id="2" name="Object 2" descr="Description" />', `<p:cNvPr id="2" name="Object 2" descr="Description">${ext}</p:cNvPr>`); await fs.writeFile(input, contentBytes(xml));
  const imported = await importPptxFile(input, { projection: 'content', decorativeSaveProfile: 'existing-v1' }); t.after(() => imported.session.dispose());
  const slide = imported.session.document.slides[0], old = slide.elements[0], edit = structuredClone(old); edit.accessibility.decorative = false; slide.replaceElement(old.id, edit);
  const plan = await imported.session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal((await writePptxFile(imported.session, plan.plan, output, { mode: 'create' })).status, 'committed');
  const reopened = await importPptx(await fs.readFile(output), { projection: 'content' }); try { assert.equal(reopened.document.slides[0].elements[0].accessibility.decorative, false); } finally { reopened.dispose(); } await noTemps(directory);
});

test('Node file adapter retains the slide visibility save profile', async t => {
  const directory = await fs.realpath(await fs.mkdtemp(path.join(tmpdir(), 'pptkit-slide-visibility-'))); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const input = path.join(directory, 'source.pptx'), output = path.join(directory, 'output.pptx'); await fs.writeFile(input, contentBytes(contentShape({ id: '2' })));
  const imported = await importPptxFile(input, { projection: 'content', slideVisibilitySaveProfile: 'attribute-v1' }); t.after(() => imported.session.dispose());
  imported.session.document.slides[0].setHidden(false); const plan = await imported.session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal((await writePptxFile(imported.session, plan.plan, output, { mode: 'create' })).status, 'committed');
  const reopened = await importPptx(await fs.readFile(output), { projection: 'content' }); try { assert.equal(reopened.document.slides[0].hidden, false); } finally { reopened.dispose(); } await noTemps(directory);
});

test('Node file adapter retains the core metadata source save profile', async t => {
  const { metadataFixture, encodeMetadata } = await import('./metadata-fixture.mjs'); const fixtureData = metadataFixture(); fixtureData.add('core', '<d:title>Original</d:title><d:creator>Alice</d:creator><c:revision>7</c:revision>');
  const directory = await fs.realpath(await fs.mkdtemp(path.join(tmpdir(), 'pptkit-core-metadata-'))); t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const input = path.join(directory, 'source.pptx'), output = path.join(directory, 'output.pptx'); await fs.writeFile(input, encodeMetadata(fixtureData));
  const imported = await importPptxFile(input, { projection: 'content', metadataSaveProfile: 'core-strings-v1' }); t.after(() => imported.session.dispose());
  imported.session.document.setMetadata({ ...imported.session.document.metadata, title: 'File title' }); const plan = await imported.session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); assert.equal((await writePptxFile(imported.session, plan.plan, output, { mode: 'create' })).status, 'committed');
  const reopened = await importPptx(await fs.readFile(output), { projection: 'content' }); try { assert.equal(reopened.document.metadata.title, 'File title'); assert.equal(reopened.document.metadata.author, 'Alice'); assert.equal(reopened.document.metadata.revision, 7); } finally { reopened.dispose(); } await noTemps(directory);
});

test('Node file adapter atomically publishes explicit default metadata removals and reinsertions', async t => {
  const directory=await fs.realpath(await fs.mkdtemp(path.join(tmpdir(),'pptkit-default-removal-file-')));t.after(()=>fs.rm(directory,{recursive:true,force:true}));
  for(const strict of [false,true]){
    const fixtureData=metadataFixture(strict);fixtureData.add('core','<d:title>Title</d:title><d:creator>Author</d:creator><d:language>fr-FR</d:language><c:revision>7</c:revision><d:subject>Sentinel</d:subject>');
    const input=path.join(directory,strict?'strict.pptx':'transitional.pptx');await fs.writeFile(input,encodeMetadata(fixtureData),{mode:0o640});const fields=['title','author','language','revision'];
    for(let round=0;round<3;round++){
      const imported=await importPptxFile(input,{projection:'content',metadataSaveProfile:'metadata-default-remove-v17'}),session=imported.session;
      try{
        session.document.setMetadata({...session.document.metadata,...(round===1?{title:'PPTKit Presentation',author:'PPTKit',language:'en-US',revision:1}:{title:undefined,author:undefined,language:undefined,revision:undefined})});
        const plan=await session.prepareSave(round===1?{mode:'edit',coreMetadataInsertions:fields}:{mode:'edit',coreMetadataRemovals:fields});assert.equal(plan.status,'ready',JSON.stringify(plan.report));const expected=(await session.save(plan.plan)).bytes;
        const result=await writePptxFile(session,plan.plan,input,{mode:'replace',expected:imported.file});assert.equal(result.status,'committed');const actual=await fs.readFile(input);assert.deepEqual(actual,Buffer.from(expected));assert.equal(result.file.sha256,hash(actual));assert.equal((await fs.stat(input)).mode&0o777,0o640);await noTemps(directory);
        const reopened=await importPptx(actual,{projection:'content'});try{for(const field of fields)assert.equal(reopened.document.metadata[field],round===1?({title:'PPTKit Presentation',author:'PPTKit',language:'en-US',revision:1})[field]:undefined);assert.equal(reopened.document.metadata.subject,'Sentinel');}finally{reopened.dispose();}
      }finally{session.dispose();}
    }
  }
});

test('stale destination prevents explicit metadata removal from replacing a changed file', async t => {
  const directory=await fs.realpath(await fs.mkdtemp(path.join(tmpdir(),'pptkit-default-removal-stale-')));t.after(()=>fs.rm(directory,{recursive:true,force:true}));const fixtureData=metadataFixture();fixtureData.add('core','<d:title>Title</d:title>');const input=path.join(directory,'source.pptx');await fs.writeFile(input,encodeMetadata(fixtureData));
  const imported=await importPptxFile(input,{projection:'content',metadataSaveProfile:'metadata-default-remove-v17'}),session=imported.session;try{session.document.setMetadata({...session.document.metadata,title:undefined});const plan=await session.prepareSave({mode:'edit',coreMetadataRemovals:['title']});assert.equal(plan.status,'ready');const changed=Buffer.from('external replacement');await fs.writeFile(input,changed);await assert.rejects(writePptxFile(session,plan.plan,input,{mode:'replace',expected:imported.file}),code('file-target-changed'));assert.deepEqual(await fs.readFile(input),changed);await noTemps(directory);}finally{session.dispose();}
});


test("file import captures custom property source inventory option", async t => {
  const directory = await fs.mkdtemp(path.join(tmpdir(), "pptkit-custom-source-"));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const f = metadataFixture(); f.add("custom", custom(2, "lpwstr", "value", 'name="A"'));
  const source = path.join(directory, "source.pptx"); await fs.writeFile(source, encodeMetadata(f));
  const imported = await importPptxFile(source, { projection: "content", customPropertySourceProfile: "raw-v1" });
  try { assert.equal(imported.session.report.customPropertySources[0].status, "bound"); }
  finally { imported.session.dispose(); }
});


test("file import forwards custom scalar preservation profile", async t => {
  const directory = await fs.mkdtemp(path.join(tmpdir(), "pptkit-custom-save-"));
  t.after(() => fs.rm(directory, { recursive: true, force: true }));
  const f = metadataFixture(); f.add("custom", custom(2, "lpwstr", "value", 'name="A"'));
  const source = path.join(directory, "source.pptx"); await fs.writeFile(source, encodeMetadata(f));
  const imported = await importPptxFile(source, { projection: "content", customPropertySaveProfile: "scalar-v1" });
  try {
    imported.session.document.setMetadata({ ...imported.session.document.metadata, customProperties: [{ name: "A", type: "string", value: "new" }] });
    const prepared = await imported.session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
  } finally { imported.session.dispose(); }
});
