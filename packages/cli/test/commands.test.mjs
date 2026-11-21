import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, readFile, readdir, symlink, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { spawnSync, execFile } from "node:child_process";
import { promisify } from "node:util";
import { createHash } from "node:crypto";
import { runCliAsync, runCli, CLI_REPORT_SCHEMA_VERSION } from "../dist/index.js";
import { textCandidates } from "../dist/commands/text-candidates.js";
import { visibilityCandidates } from "../dist/commands/slide-visibility-edits.js";
import { metadataCandidates } from "../dist/commands/core-metadata-edits.js";
import { importPptx, compareOpcPackages, inspectPptx, inspectOpcClosure, inspectPptxNotes } from "../../pptx-parser/dist/index.js";
import { contentFixture, shape, body } from "../../pptx-parser/test/content-fixture.mjs";
import { group } from "../../pptx-parser/test/reference-fixture.mjs";
import { notesFixture, notesShape } from "../../pptx-parser/test/notes-fixture.mjs";
import { notesInventoryFixture } from "../../pptx-parser/test/notes-inventory-fixture.mjs";
import { table, cell, tableRow, tableFixture } from "../../pptx-parser/test/table-fixture.mjs";
import { chartFixture, chartXml, series, workbook, encodeChart, frame } from "../../pptx-parser/test/chart-fixture.mjs";
import { metadataFixture, encodeMetadata, custom } from "../../pptx-parser/test/metadata-fixture.mjs";
import { imageFixture, encodeImages, picture, png, validGif } from "../../pptx-parser/test/image-fixture.mjs";
import { encodeStructure } from "../../pptx-parser/test/structure-fixture.mjs";
import { relationships } from "../../pptx-parser/test/opc-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { diagramTextFixture } from "../../pptx-parser/test/diagram-text-fixture.mjs";
const run = promisify(execFile), sha = b => createHash("sha256").update(b).digest("hex");
async function fixture(t, bytes = encodeStructure(contentFixture())) {
  const directory = await mkdtemp(path.join(tmpdir(), "pptkit-cli-")); t.after(() => rm(directory, { recursive: true, force: true }));
  await writeFile(path.join(directory, "source.pptx"), bytes); return { directory, bytes, options: { cwd: directory } };
}
const decode = result => { const json = JSON.parse(result.output); assert.equal(json.schemaVersion, CLI_REPORT_SCHEMA_VERSION); return json; };

test('optional CLI notes inventory matches public Parser output and preserves the ordinary report', async t => {
  for (const strict of [false, true]) {
    const bytes = encodeStructure(notesInventoryFixture(strict)), f = await fixture(t, bytes);
    const run = (...args) => runCliAsync(['inspect', 'source.pptx', ...args, '--json'], f.options);
    const plainResult = await run(); assert.equal(plainResult.exitCode, 0); const plain = decode(plainResult);
    const result = await run('--notes-inventory'); assert.equal(result.exitCode, 0, result.output);
    const actual = decode(result);
    assert.deepEqual(actual.report, plain.report);
    assert.equal(actual.inputs[0].sha256, sha(bytes));
    assert.deepEqual(actual.notesInventory, await inspectPptxNotes(bytes));
    assert.equal(actual.notesInventory.fieldEvaluation, 'not-performed');
    assert.ok(actual.notesInventory.parts.some(part => part.fields.length && part.drawings.some(drawing => drawing.role === 'slide-image')));
    const combined = decode(await run('--capabilities', '--package-summary', '--notes-inventory'));
    assert.equal(combined.reportKind, 'capabilities'); assert.ok(combined.packageSummary);
    assert.deepEqual(combined.notesInventory, actual.notesInventory);
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), bytes);
    const human = await runCliAsync(['inspect', 'source.pptx', '--notes-inventory'], f.options);
    assert.equal(human.exitCode, 0); assert.match(human.output, /Notes source inventory: 2 slide\(s\), 2 notes\/master part\(s\)/);
    assert.match(human.output, /body: unique-direct/); assert.match(human.output, /body: absent/);
    assert.match(human.output, /field records: 3/); assert.match(human.output, /fields were not evaluated/);
    assert.doesNotMatch(human.output, /2042-01-02 & stored/);
    const humanCapabilities = await runCliAsync(['inspect', 'source.pptx', '--capabilities', '--notes-inventory'], f.options);
    assert.equal(humanCapabilities.exitCode, 0); assert.match(humanCapabilities.output, /PPTKit capability view/); assert.match(humanCapabilities.output, /Notes source inventory/);
    const humanBase = await runCliAsync(['inspect', 'source.pptx'], f.options);
    const humanLimit = String(Buffer.byteLength(humanBase.output) + 1);
    const humanLimited = await runCliAsync(['inspect', 'source.pptx', '--notes-inventory', '--max-output-bytes', humanLimit], f.options);
    assert.equal(humanLimited.exitCode, 4); assert.doesNotMatch(humanLimited.output, /Notes source inventory/);
    const plainLimit = String(Buffer.byteLength(plainResult.output) + 1);
    assert.equal((await run('--max-output-bytes', plainLimit)).exitCode, 0);
    const limited = await run('--notes-inventory', '--max-output-bytes', plainLimit);
    assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
  }
  for (const args of [['compare', 'a', 'b', '--notes-inventory', '--json'], ['validate', 'a', '--notes-inventory', '--json'], ['inspect', 'missing', '--notes-inventory', '--notes-inventory', '--json']]) {
    assert.equal((await runCliAsync(args)).exitCode, 2);
  }
  const abort = new AbortController(); abort.abort();
  const result = await runCliAsync(['inspect', 'missing', '--notes-inventory', '--json'], { signal: abort.signal });
  assert.equal(result.exitCode, 130);
});

test('CLI notes inspection budgets enforce exact Parser thresholds and reject invalid options before I/O', async t => {
  const bytes = encodeStructure(notesInventoryFixture()), f = await fixture(t, bytes);
  const inventory = await inspectPptxNotes(bytes);
  for (const [flag, field] of [['--max-notes-work', 'workUnits'], ['--max-notes-records', 'recordCount'], ['--max-notes-text', 'textCharacters']]) {
    const run = value => runCliAsync(['inspect', 'source.pptx', '--notes-inventory', flag, String(value), '--json'], f.options);
    const exact = await run(inventory[field]); assert.equal(exact.exitCode, 0, exact.output);
    assert.deepEqual(decode(exact).notesInventory, inventory);
    const exhausted = await run(inventory[field] - 1); assert.equal(exhausted.exitCode, 4);
    assert.equal(decode(exhausted).error.code, 'package-limit-exceeded');
    const human = await runCliAsync(['inspect', 'source.pptx', '--notes-inventory', flag, String(inventory[field] - 1)], f.options);
    assert.equal(human.exitCode, 4); assert.doesNotMatch(human.output, /Notes source inventory:/);
  }
  const bad = notesFixture(notesShape({ id: 'bad' }) + notesShape({ id: 'wrong', role: 'undefined-kind' }));
  const badBytes = encodeStructure(bad), badFile = await fixture(t, badBytes);
  const badInventory = await inspectPptxNotes(badBytes);
  const findingsCount = badInventory.findings.length + badInventory.structureFindings.length;
  assert.ok(findingsCount > 1);
  const findingsExact = await runCliAsync(['inspect', 'source.pptx', '--notes-inventory', '--max-notes-findings', String(findingsCount), '--json'], badFile.options);
  assert.equal(findingsExact.exitCode, 0, findingsExact.output);
  assert.deepEqual(decode(findingsExact).notesInventory, badInventory);
  const exhausted = await runCliAsync(['inspect', 'source.pptx', '--notes-inventory', '--max-notes-findings', '1', '--json'], badFile.options);
  assert.equal(exhausted.exitCode, 4); assert.equal(decode(exhausted).error.code, 'package-limit-exceeded');
  for (const flag of ['--max-notes-work', '--max-notes-records', '--max-notes-findings', '--max-notes-text']) {
    for (const value of ['0', '-1', '01', '1.2', '1e3', '100000001']) {
      assert.equal((await runCliAsync(['inspect', 'missing', '--notes-inventory', flag, value, '--json'])).exitCode, 2);
    }
    assert.equal((await runCliAsync(['inspect', 'missing', flag, '1', '--json'])).exitCode, 2);
    assert.equal((await runCliAsync(['compare', 'a', 'b', '--notes-inventory', flag, '1', '--json'])).exitCode, 2);
    assert.equal((await runCliAsync(['inspect', 'missing', '--notes-inventory', flag])).exitCode, 2);
  }
  assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), bytes);
});

test('CLI preview propagates JPEG precision warnings across authoring, publication and PPTX import', async t => {
  const f = await fixture(t), twelve = await readFile(new URL('../../../tests/fixtures/jpeg-12bit-progressive.jpg', import.meta.url));
  const eight = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const asset = (id, bytes) => ({ id, kind: 'image', mimeType: 'image/jpeg', source: { type: 'bytes', encoding: 'base64', value: bytes.toString('base64') } });
  const slide = (id, assetId) => ({ id, elements: [{ type: 'image', assetId, box: { x: 0, y: 0, width: 80, height: 40 } }] });
  const authoring = JSON.stringify({ version: 'pptkit-authoring-v1', assets: [asset('twelve', twelve), asset('eight', eight)], slides: [slide('s1', 'twelve'), slide('s2', 'eight')] });
  await writeFile(path.join(f.directory, 'jpeg.json'), authoring);
  const run = (...args) => runCliAsync([...args, '--json'], f.options);
  const first = await run('preview', 'jpeg.json', '--slide', '1'); assert.equal(first.exitCode, 0, first.output);
  const report = decode(first).report;
  assert.equal(report.status, 'rendered-with-warnings'); assert.deepEqual(report.warnings.map(w => [w.code, w.slideId, w.assetId]), [['jpeg-preview-precision-unverified', 's1', 'twelve']]);
  assert.ok(report.slide.svg.includes(twelve.toString('base64')));
  const second = await run('preview', 'jpeg.json', '--slide', '2'); assert.equal(second.exitCode, 0, second.output); assert.deepEqual(decode(second).report.warnings, []);
  const output = await run('preview', 'jpeg.json', '--slide', '1', '--output', 'preview.svg'); assert.equal(output.exitCode, 0, output.output);
  assert.deepEqual(decode(output).report.warnings, report.warnings); assert.equal(await readFile(path.join(f.directory, 'preview.svg'), 'utf8'), report.slide.svg);
  const generated = await run('generate', 'jpeg.json', 'jpeg.pptx'); assert.equal(generated.exitCode, 0, generated.output);
  const imported = await run('preview', 'jpeg.pptx', '--input-format', 'pptx', '--slide', '1'); assert.equal(imported.exitCode, 0, imported.output);
  assert.ok(decode(imported).report.warnings.some(w => w.code === 'jpeg-preview-precision-unverified'));
  assert.ok(decode(imported).report.slide.svg.includes(twelve.toString('base64')));
  assert.equal(await readFile(path.join(f.directory, 'jpeg.json'), 'utf8'), authoring);
});

test("strict arguments reject unsupported/duplicate/missing/cross-command values before I/O", async () => {
  for (const args of [["other"], ["inspect"], ["compare", "a"], ["inspect", "a", "b"], ["inspect", "a", "--bad"], ["inspect", "a", "--json", "--json"], ["inspect", "a", "--projection"], ["inspect", "a", "--profile", "opc-xml-v1"], ["compare", "a", "b", "--projection", "content"], ["inspect", "a", "--diagram-preview", "cache"], ["inspect", "a", "--diagram-text-comparison", "true"], ["inspect", "a", "--projection", "content", "--edit-save-profile", "v4"], ["inspect", "a", "--edit-save-profile", "text-runs-v3"], ...["0", "-1", "NaN", "1.5", "1e3", "01", "9007199254740992", "536870913"].map(n => ["inspect", "missing", "--max-input-bytes", n]), ["compare", "a", "b", "--max-differences", "1000001"], ["inspect", "a", "--max-output-bytes", "511"], ["inspect", "a", "--max-output-bytes", "67108865"], ["inspect", "a\0b"]]) {
    const r = await runCliAsync([...args, "--json"]); assert.equal(r.exitCode, 2, JSON.stringify(args)); assert.equal(r.channel, "stderr"); assert.equal(decode(r).error.code, "usage");
  }
  assert.equal(runCli(["inspect", "a"]).exitCode, 2);
  assert.equal((await runCliAsync(["compare", "--help"])).exitCode, 0);
  assert.equal((await runCliAsync()).exitCode, 0);
});

test("inspect JSON is the exact public import report with deterministic relative identities and preserved files", async t => {
  for (const strict of [false, true]) {
    const f = await fixture(t, encodeStructure(contentFixture(undefined, strict)));
    for (const projection of ["structure", "content"]) {
      const result = await runCliAsync(["inspect", "source.pptx", "--projection", projection, "--json"], f.options); assert.equal(result.exitCode, 0);
      const body = decode(result), expected = await importPptx(f.bytes, { projection });
      try { assert.deepEqual(body.report, expected.report); } finally { expected.dispose(); }
      assert.deepEqual(body.inputs, [{ path: "source.pptx", byteLength: f.bytes.length, sha256: sha(f.bytes) }]);
      assert.equal(body.nativeVerification, "not-performed"); assert.equal(body.command, "inspect"); assert.equal(body.status, "ok");
      assert.equal((await runCliAsync(["inspect", path.join(f.directory, "source.pptx"), "--projection", projection, "--json"], f.options)).output, result.output);
      assert.deepEqual(await readFile(path.join(f.directory, "source.pptx")), Buffer.from(f.bytes));
    }
  }
});

test('package summary reports declared media sizes and relationship counts without changing the import report', async t => {
  for (const strict of [false, true]) {
    const f = await fixture(t, encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }), strict)));
    const base = decode(await runCliAsync(['inspect', 'source.pptx', '--json'], f.options));
    const result = await runCliAsync(['inspect', 'source.pptx', '--package-summary', '--json'], f.options);
    assert.equal(result.exitCode, 0);
    const actual = decode(result), summary = actual.packageSummary, pkg = base.report.inspection.package;
    assert.deepEqual(actual.report, base.report);
    assert.equal(summary.schemaVersion, 'pptkit-cli-package-summary-v1');
    assert.equal(summary.mediaClassification, 'declared-content-type');
    assert.equal(summary.relationships.total, pkg.relationships.length);
    assert.equal(summary.relationships.internal + summary.relationships.external, summary.relationships.total);
    assert.equal(summary.archive.inputBytes, f.bytes.length);
    assert.equal(summary.media.length, 1);
    const media = summary.media[0], entry = pkg.archive.entries.find(item => `/${item.name}` === media.part);
    assert.equal(media.contentType, 'image/png');
    assert.equal(media.expandedBytes, png.length);
    assert.equal(media.compressedBytes, entry.compressedBytes);
    assert.equal(media.incomingRelationships, pkg.relationships.filter(r => r.targetPart === media.part).length);
    assert.deepEqual(summary.mediaTotals, { parts: 1, compressedBytes: entry.compressedBytes, expandedBytes: png.length });
    const capabilities = decode(await runCliAsync(['inspect', 'source.pptx', '--capabilities', '--package-summary', '--json'], f.options));
    assert.deepEqual(capabilities.packageSummary, summary);
    assert.equal(capabilities.reportKind, 'capabilities');
    const limited = await runCliAsync(['inspect', 'source.pptx', '--package-summary', '--max-output-bytes', '512', '--json'], f.options);
    assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), Buffer.from(f.bytes));
  }
  for (const args of [['inspect', 'missing', '--package-summary'], ['compare', 'missing', 'missing', '--package-summary', '--json'], ['inspect', 'missing', '--package-summary', '--package-summary', '--json']]) {
    const result = await runCliAsync(args); assert.equal(result.exitCode, 2);
  }
});

test('package summary counts orphaned declarations and external relationships without fetching', async t => {
  const source = imageFixture();
  source.parts['media/unlinked.bin'] = png;
  source.mediaTypes['media/unlinked.bin'] = 'image/jpeg';
  source.parts['deck/_rels/z.xml.rels'] = relationships([source.rel('layout', 'slideLayout', 'layout.xml'), source.rel('image', 'image', '../media/raster.bin'), { ...source.rel('remote', 'image', 'https://example.invalid/image.png'), TargetMode: 'External' }]);
  const f = await fixture(t, encodeImages(source));
  const result = await runCliAsync(['inspect', 'source.pptx', '--package-summary', '--json'], f.options);
  assert.equal(result.exitCode, 0);
  const summary = decode(result).packageSummary;
  assert.equal(summary.relationships.external, 1);
  assert.equal(summary.mediaTotals.parts, 2);
  const unlinked = summary.media.find(item => item.part === '/media/unlinked.bin');
  assert.equal(unlinked.incomingRelationships, 0);
  assert.equal(unlinked.contentType, 'image/jpeg');
  assert.equal(unlinked.expandedBytes, png.length);
});

test('documented diagnostic workflow keeps source identity through validation, preview and no-op comparison', async t => {
  const f = await fixture(t);
  const authoring = { version: 'pptkit-authoring-v1', presentation: { size: { width: 960, height: 540 } }, slides: [{ id: 'page', elements: [{ type: 'text', id: 'title', box: { x: 40, y: 40, width: 400, height: 80 }, content: 'Diagnostic example' }] }] };
  await writeFile(path.join(f.directory, 'deck.json'), JSON.stringify(authoring));
  const runStep = async args => { const result = await runCliAsync([...args, '--json'], f.options); assert.equal(result.exitCode, 0, result.output); return decode(result); };
  const validation = await runStep(['validate', 'deck.json', '--export']);
  assert.equal(validation.report.valid, true); assert.equal(validation.report.export.status, 'checked');
  const generated = await runStep(['generate', 'deck.json', 'generated.pptx']);
  assert.equal(generated.report.status, 'committed');
  const inspect = await runStep(['inspect', 'generated.pptx', '--projection', 'content', '--package-summary']);
  const capabilities = await runStep(['inspect', 'generated.pptx', '--projection', 'content', '--capabilities']);
  const preview = await runStep(['preview', 'generated.pptx', '--input-format', 'pptx', '--slide', '1']);
  const dryRun = await runStep(['save', 'generated.pptx', 'copy.pptx', '--mode', 'no-op', '--dry-run']);
  assert.equal(dryRun.status, 'ready'); await assert.rejects(readFile(path.join(f.directory, 'copy.pptx')), { code: 'ENOENT' });
  const saved = await runStep(['save', 'generated.pptx', 'copy.pptx', '--mode', 'no-op']);
  assert.equal(saved.report.status, 'committed');
  const compare = await runStep(['compare', 'generated.pptx', 'copy.pptx']);
  assert.equal(compare.report.equal, true);
  const sourceBytes = await readFile(path.join(f.directory, 'generated.pptx'));
  const copyBytes = await readFile(path.join(f.directory, 'copy.pptx'));
  const closure = await inspectOpcClosure(sourceBytes, { roots: ['/ppt/slides/slide1.xml'] });
  assert.deepEqual(await inspectOpcClosure(copyBytes, { roots: ['/ppt/slides/slide1.xml'] }), closure);
  assert.equal(sha(sourceBytes), sha(copyBytes));
  await assert.rejects(inspectOpcClosure(sourceBytes, { roots: ['/missing.xml'] }));
  assert.deepEqual(await readFile(path.join(f.directory, 'generated.pptx')), sourceBytes);
  assert.match(preview.report.slide.svg, /Diagnostic example/);
  assert.equal(inspect.packageSummary.archive.inputBytes, (await readFile(path.join(f.directory, 'generated.pptx'))).length);
  assert.equal(capabilities.report.nativeVerification, 'not-performed');
  for (const result of [inspect, capabilities, preview, dryRun, saved, compare]) assert.equal(result.inputs[0].sha256, inspect.inputs[0].sha256);
});

test("inspect propagates explicit cache/text/save profiles without granting native edit authority", async t => {
  const input = diagramTextFixture(); input.setText(["Data marker"], ["Stored cache"]); const f = await fixture(t, input.bytes());
  const args = ["inspect", "source.pptx", "--projection", "content", "--diagram-preview", "cache", "--diagram-text-comparison", "direct-paragraphs", "--edit-save-profile", "text-runs-v3", "--json"];
  const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0); const body = decode(result);
  assert.equal(body.report.diagramPreviews[0].freshness, "unknown"); assert.equal(body.report.inspection.diagramInventory.textComparisons[0].shapes[0].status, "different");
  assert.ok(body.report.textRuns.length); assert.ok(body.report.findings.some(f => f.feature === "smartart.cache-preview" && f.operation === "edit-save" && f.tier === "unsupported"));
});

test("compare defaults to exact payloads, XML equality is explicit, and differences/truncation stay independent", async t => {
  const before = contentFixture(), after = contentFixture(); after.parts["deck/z.xml"] = after.parts["deck/z.xml"].replace('show="0"', 'show = "0"');
  const f = await fixture(t, encodeStructure(before)); const changed = encodeStructure(after); await writeFile(path.join(f.directory, "after.pptx"), changed);
  for (const profile of ["opc-payload-v1", "opc-xml-v1"]) {
    const result = await runCliAsync(["compare", "source.pptx", "after.pptx", "--profile", profile, "--json"], f.options), body = decode(result);
    assert.deepEqual(body.report, await compareOpcPackages(f.bytes, changed, { profile, comparisonLimits: { maxDifferences: 1000 } }));
    assert.equal(result.exitCode, profile === "opc-payload-v1" ? 1 : 0); assert.notEqual(result.channel, "stderr");
  }
  const base = await runCliAsync(["compare", "source.pptx", "source.pptx", "--json"], f.options); assert.equal(base.exitCode, 0); assert.equal(decode(base).report.equal, true);
  for (const key of ["deck/z.xml", "deck/a.xml", "deck/layout.xml"]) after.parts[key] = after.parts[key].replace('</p:', '<!-- changed --></p:');
  await writeFile(path.join(f.directory, "after.pptx"), encodeStructure(after));
  const limited = await runCliAsync(["compare", "source.pptx", "after.pptx", "--max-differences", "1", "--json"], f.options), body = decode(limited);
  assert.equal(limited.exitCode, 1); assert.equal(body.report.truncated, true); assert.equal(body.report.differences.length, 1); assert.ok(body.report.differenceCount > 1);
  assert.equal(body.report.profile, "opc-payload-v1");
});

test("input/output limits and invalid packages yield bounded machine errors instead of partial success", async t => {
  const f = await fixture(t);
  for (const [args, code, error] of [
    [["inspect", "source.pptx", "--max-input-bytes", "1"], 4, "input-limit-exceeded"],
    [["inspect", "source.pptx", "--max-output-bytes", "512"], 4, "output-limit-exceeded"],
    [["inspect", "missing.pptx"], 3, "file-io"],
    [["inspect", "."], 3, "file-not-regular"],
  ]) { const r = await runCliAsync([...args, "--json"], f.options); assert.equal(r.exitCode, code); assert.equal(decode(r).error.code, error); assert.ok(Buffer.byteLength(r.output) + 1 <= 512); }
  await writeFile(path.join(f.directory, "broken.pptx"), f.bytes.subarray(0, 100));
  const bad = await runCliAsync(["inspect", "broken.pptx", "--json"], f.options); assert.equal(bad.exitCode, 3); assert.equal(decode(bad).error.code, "import-invalid-package");
  const compared = await runCliAsync(["compare", "source.pptx", "broken.pptx", "--json"], f.options); assert.equal(compared.exitCode, 3); assert.match(decode(compared).error.code, /^archive-/);
  const retry = await runCliAsync(["inspect", "source.pptx", "--json"], f.options); assert.equal(retry.exitCode, 0);
});

test("option/argument capture and cancellation retain invocation-time ownership", async t => {
  const f = await fixture(t), args = ["inspect", "source.pptx", "--json"], options = { ...f.options };
  const pending = runCliAsync(args, options); args[1] = "missing"; options.cwd = "/missing";
  assert.equal((await pending).exitCode, 0);
  for (const timing of ["before", "during"]) {
    const controller = new AbortController(); if (timing === "before") controller.abort();
    const pending = runCliAsync(["inspect", "source.pptx", "--json"], { ...f.options, signal: controller.signal }); if (timing === "during") controller.abort();
    const result = await pending; assert.equal(result.exitCode, 130); assert.equal(decode(result).error.code, "operation-aborted");
  }
  assert.equal((await runCliAsync(["inspect", "source.pptx"], f.options)).exitCode, 0);
});

test("dash paths, Unicode/control names and symlinks are file arguments without terminal injection", async t => {
  const f = await fixture(t); await writeFile(path.join(f.directory, "--json"), f.bytes);
  assert.equal((await runCliAsync(["inspect", "--", "--json"], f.options)).exitCode, 0);
  const name = "中文\u001b[31m.pptx"; await writeFile(path.join(f.directory, name), f.bytes);
  const result = await runCliAsync(["inspect", name], f.options); assert.equal(result.exitCode, 0); assert.ok(!result.output.includes("\u001b")); assert.ok(result.output.includes("\\u001b"));
  if (process.platform !== "win32") {
    await symlink("source.pptx", path.join(f.directory, "link.pptx")); const linked = await runCliAsync(["inspect", "link.pptx", "--json"], f.options); assert.equal(linked.exitCode, 0); assert.equal(decode(linked).inputs[0].path, "link.pptx");
    const fifo = path.join(f.directory, "pipe"); assert.equal(spawnSync("mkfifo", [fifo]).status, 0);
    const nonregular = await runCliAsync(["inspect", "pipe", "--json"], f.options); assert.equal(nonregular.exitCode, 3); assert.equal(decode(nonregular).error.code, "file-not-regular");
  }
});

test("actual executable routes JSON success/differences to stdout and errors to stderr with stable status", async t => {
  const f = await fixture(t), executable = new URL("../dist/cli.js", import.meta.url).pathname;
  const good = await run(process.execPath, [executable, "inspect", "source.pptx", "--json"], { cwd: f.directory }); assert.equal(good.stderr, ""); assert.equal(JSON.parse(good.stdout).command, "inspect");
  const changed = contentFixture(); changed.parts["deck/z.xml"] = changed.parts["deck/z.xml"].replace('show="0"', 'show="1"'); await writeFile(path.join(f.directory, "after.pptx"), encodeStructure(changed));
  for (const [args, code, channel] of [[["compare", "source.pptx", "after.pptx", "--json"], 1, "stdout"], [["inspect", "missing", "--json"], 3, "stderr"], [["inspect", "source.pptx", "--max-output-bytes", "512", "--json"], 4, "stderr"], [["bad", "--json"], 2, "stderr"]]) {
    let failure; try { await run(process.execPath, [executable, ...args], { cwd: f.directory }); } catch (e) { failure = e; }
    assert.equal(failure?.code, code); assert.equal(failure[channel === "stdout" ? "stderr" : "stdout"], ""); assert.equal(JSON.parse(failure[channel]).schemaVersion, CLI_REPORT_SCHEMA_VERSION);
  }
});


test("bounded JSON/human encoders charge exact escaped UTF-8 bytes and final newline", async () => {
  const { boundedJson, boundedText } = await import("../dist/commands/format.js");
  const value = { text: "中文 😀\r\n\t\u0000\\\"", nested: [true, false, null, 12.5, { empty: "", absent: undefined }] };
  const expected = JSON.stringify(value), budget = Buffer.byteLength(expected) + 1;
  assert.equal(boundedJson(value, budget), expected);
  assert.throws(() => boundedJson(value, budget - 1), e => e.code === "output-limit-exceeded");
  assert.throws(() => boundedJson({ huge: "x".repeat(10000) }, 512), e => e.code === "output-limit-exceeded");
  const lines = ["中 😀", "next"], human = lines.join("\n"), textBudget = Buffer.byteLength(human) + 1;
  assert.equal(boundedText(lines, textBudget), human); assert.throws(() => boundedText(lines, textBudget - 1), e => e.code === "output-limit-exceeded");
  assert.equal(boundedText(["\u001b[31m\u202e"], 512), "\\u001b[31m\\u202e");
});


test("versioned capability view supports exact filters and counts-only truncation without changing default report", async t => {
  const input = diagramTextFixture(), f = await fixture(t, input.bytes());
  const base = ["inspect", "source.pptx", "--projection", "content", "--diagram-preview", "cache", "--json"];
  const standard = decode(await runCliAsync(base, f.options)); assert.equal(standard.reportKind, undefined); assert.ok(standard.report.inspection);
  const result = await runCliAsync([...base, "--capabilities", "--capability-feature", "smartart.cache-preview", "--capability-part", "/deck/z.xml"], f.options); assert.equal(result.exitCode, 0);
  const body = decode(result); assert.equal(body.reportKind, "capabilities"); assert.equal(body.report.schemaVersion, "pptkit-import-capabilities-v1"); assert.equal(body.report.totalSubjects, 1);
  assert.equal(body.report.subjects[0].operations[6].tier, "unsupported"); assert.equal(body.report.nativeVerification, "not-performed");
  const count = decode(await runCliAsync([...base, "--capabilities", "--max-capability-subjects", "0"], f.options)); assert.ok(count.report.totalSubjects > 0); assert.equal(count.report.truncated, true); assert.deepEqual(count.report.subjects, []);
  for (const args of [["inspect", "source.pptx", "--capability-feature", "x"], ["inspect", "source.pptx", "--max-capability-subjects", "1"], ["inspect", "source.pptx", "--capabilities", "--max-capability-subjects", "150001"], ["compare", "source.pptx", "source.pptx", "--capabilities"]]) assert.equal((await runCliAsync([...args, "--json"], f.options)).exitCode, 2);
});

test('save no-op creates byte-identical files and never replaces an existing destination', async t => {
  const f = await fixture(t);
  const args = ['save', 'source.pptx', 'copy.pptx', '--mode', 'no-op', '--json'];
  const result = await runCliAsync(args, f.options), body = decode(result);
  assert.equal(result.exitCode, 0); assert.equal(body.command, 'save'); assert.equal(body.destination, 'copy.pptx');
  assert.equal(body.report.status, 'committed'); assert.equal(body.report.assurance, 'original-bytes');
  assert.deepEqual(await readFile(path.join(f.directory, 'copy.pptx')), Buffer.from(f.bytes));
  const again = await runCliAsync(args, f.options); assert.equal(again.exitCode, 3); assert.equal(decode(again).error.code, 'file-target-exists');
  const same = await runCliAsync(['save', 'source.pptx', 'source.pptx', '--mode', 'no-op', '--json'], f.options);
  assert.equal(same.exitCode, 3); assert.equal(decode(same).error.code, 'file-target-exists');
  assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), Buffer.from(f.bytes));
  await symlink('source.pptx', path.join(f.directory, 'link.pptx'));
  const linked = await runCliAsync(['save', 'source.pptx', 'link.pptx', '--mode', 'no-op', '--json'], f.options);
  assert.equal(linked.exitCode, 3); assert.equal(decode(linked).error.code, 'file-target-exists');
});

test('save rejects unsafe requests before publication and concurrent create has exactly one winner', async t => {
  const f = await fixture(t);
  for (const args of [['save', 'source.pptx', 'out.pptx'], ['save', 'source.pptx', 'out.pptx', '--mode', 'edit'], ['save', 'source.pptx', 'out.pptx', '--mode', 'no-op', '--projection', 'content']]) {
    const r = await runCliAsync([...args, '--json'], f.options); assert.equal(r.exitCode, 2);
  }
  const longDestination = 'x'.repeat(230) + '.pptx';
  const limited = await runCliAsync(['save', 'source.pptx', longDestination, '--mode', 'no-op', '--max-output-bytes', '512', '--json'], f.options);
  assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
  await assert.rejects(readFile(path.join(f.directory, longDestination)), { code: 'ENOENT' });
  await writeFile(path.join(f.directory, 'bad.pptx'), Buffer.from('bad'));
  const bad = await runCliAsync(['save', 'bad.pptx', 'bad-out.pptx', '--mode', 'no-op', '--json'], f.options); assert.equal(bad.exitCode, 3);
  await assert.rejects(readFile(path.join(f.directory, 'bad-out.pptx')), { code: 'ENOENT' });
  const abort = new AbortController(); abort.abort();
  const cancelled = await runCliAsync(['save', 'source.pptx', 'cancel.pptx', '--mode', 'no-op', '--json'], { ...f.options, signal: abort.signal });
  assert.equal(cancelled.exitCode, 130); await assert.rejects(readFile(path.join(f.directory, 'cancel.pptx')), { code: 'ENOENT' });
  const args = ['save', 'source.pptx', 'race.pptx', '--mode', 'no-op', '--json'];
  const outcomes = await Promise.all([runCliAsync(args, f.options), runCliAsync(args, f.options)]);
  assert.equal(outcomes.filter(r => r.exitCode === 0).length, 1); assert.equal(outcomes.filter(r => r.exitCode === 3).length, 1);
  assert.deepEqual(await readFile(path.join(f.directory, 'race.pptx')), Buffer.from(f.bytes));
  const mutableArgs = ['save', 'source.pptx', 'captured.pptx', '--mode', 'no-op'], mutableOptions = { ...f.options };
  const pending = runCliAsync(mutableArgs, mutableOptions); mutableArgs[2] = 'changed.pptx'; mutableOptions.cwd = '/missing';
  const human = await pending; assert.equal(human.exitCode, 0); assert.match(human.output, /save committed: captured\.pptx/);
  assert.deepEqual(await readFile(path.join(f.directory, 'captured.pptx')), Buffer.from(f.bytes));
  await assert.rejects(readFile(path.join(f.directory, 'changed.pptx')), { code: 'ENOENT' });
});

test('CLI edits starting numbers through scoped preservation, including explicit default and clear', async t => {
  for (const strict of [false, true]) {
    const f = await fixture(t, encodeStructure(contentFixture(undefined, strict)));
    let source = 'source.pptx';
    for (const [index, value] of ['1', '0', '9999', 'clear'].entries()) {
      const target = `round-${index}.pptx`;
      const result = await runCliAsync(['save', source, target, '--mode', 'edit', '--first-slide-number', value, '--json'], f.options);
      assert.equal(result.exitCode, 0, result.output); const body = decode(result);
      assert.equal(body.edit.firstSlideNumber, value === 'clear' ? null : Number(value));
      assert.equal(body.report.assurance, 'scoped-preservation'); assert.equal(body.report.changes.length, 1);
      const output = await readFile(path.join(f.directory, target)), info = await inspectPptx(output);
      assert.equal(info.firstSlideNumber, value === 'clear' ? 1 : Number(value));
      assert.equal(info.firstSlideNumberDeclared, value !== 'clear');
      const input = await readFile(path.join(f.directory, source));
      const diff = await compareOpcPackages(input, output, { profile: 'opc-payload-v1' });
      assert.deepEqual([...new Set(diff.differences.map(d => d.part))], ['/deck/main.data']);
      source = target;
    }
    const noChange = await runCliAsync(['save', source, 'unchanged.pptx', '--mode', 'edit', '--first-slide-number', 'clear', '--json'], f.options);
    assert.equal(noChange.exitCode, 3); assert.equal(decode(noChange).error.code, 'save-edit-no-changes');
    await assert.rejects(readFile(path.join(f.directory, 'unchanged.pptx')), { code: 'ENOENT' });
    const existing = await runCliAsync(['save', source, 'source.pptx', '--mode', 'edit', '--first-slide-number', '17', '--json'], f.options);
    assert.equal(existing.exitCode, 3); assert.equal(decode(existing).error.code, 'file-target-exists');
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), Buffer.from(f.bytes));
    const oversized = 'y'.repeat(230) + '.pptx';
    const bounded = await runCliAsync(['save', source, oversized, '--mode', 'edit', '--first-slide-number', '17', '--max-output-bytes', '512', '--json'], f.options);
    assert.equal(bounded.exitCode, 4); assert.equal(decode(bounded).error.code, 'output-limit-exceeded');
    await assert.rejects(readFile(path.join(f.directory, oversized)), { code: 'ENOENT' });
  }
});

test('CLI rejects ambiguous or invalid starting-number requests before reading or writing', async () => {
  for (const value of ['-1', '10000', '01', '+1', '1.0', '1e2', 'CLEAR']) {
    const r = await runCliAsync(['save', 'missing', 'out', '--mode', 'edit', '--first-slide-number', value, '--json']);
    assert.equal(r.exitCode, 2, value); assert.equal(decode(r).error.code, 'usage');
  }
  for (const args of [['save', 'missing', 'out', '--mode', 'no-op', '--first-slide-number', '17'], ['inspect', 'missing', '--first-slide-number', '17'], ['compare', 'a', 'b', '--first-slide-number', '17']]) {
    assert.equal((await runCliAsync(args)).exitCode, 2);
  }
});

test('source-bound text manifests save three rounds and reject stale or ambiguous targets', async t => {
  for (const strict of [false, true]) {
    const source = encodeStructure(contentFixture(undefined, strict)), f = await fixture(t, source); let bytes = source, from = 'source.pptx';
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v1' });
      const binding = session.report.textRuns.find(b => b.status === 'bound'), element = session.document.slides.find(s => s.id === binding.slideId).elements.find(e => e.id === binding.elementId);
      const expected = element.content[binding.paragraphIndex].runs[binding.runIndex].text; session.dispose();
      const manifest = { version: 'pptkit-text-edits-v1', sourceSha256: sha(bytes), edits: [{ slideId: binding.slideId, elementId: binding.elementId, paragraphIndex: binding.paragraphIndex, runIndex: binding.runIndex, expected, replacement: `Round ${round} <&>🙂` }] };
      await writeFile(path.join(f.directory, 'edits.json'), JSON.stringify(manifest));
      const to = `round ${round}.pptx`, result = await runCliAsync(['save', from, to, '--mode', 'edit', '--text-edits', 'edits.json', '--json'], f.options);
      assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).report.assurance, 'scoped-preservation'); assert.equal(decode(result).inputs.length, 2); assert.equal(decode(result).edit.textEdits, true);
      const saved = await readFile(path.join(f.directory, to)), before = readGeneratedZip(bytes), after = readGeneratedZip(saved);
      assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort()); for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
      const xmlText = text => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      assert.equal(after.get('deck/z.xml').toString(), before.get('deck/z.xml').toString().replace(`>${xmlText(expected)}</a:t>`, `>${xmlText(manifest.edits[0].replacement)}</a:t>`));
      const reopened = await importPptx(saved, { projection: 'content', editSaveProfile: 'text-runs-v1' }); assert.equal(reopened.document.slides.find(s => s.id === binding.slideId).elements.find(e => e.id === binding.elementId).content[binding.paragraphIndex].runs[binding.runIndex].text, manifest.edits[0].replacement); reopened.dispose();
      bytes = saved; from = to;
    }
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
    const originalSession = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v1' }); const b = originalSession.report.textRuns.find(x => x.status === 'bound'); originalSession.dispose();
    const valid = { version: 'pptkit-text-edits-v1', sourceSha256: sha(source), edits: [{ slideId: b.slideId, elementId: b.elementId, paragraphIndex: b.paragraphIndex, runIndex: b.runIndex, expected: ' A & B ', replacement: 'changed' }] };
    const cases = [
      [{ ...valid, sourceSha256: '0'.repeat(64) }, 'text-edits-source-mismatch'],
      [{ ...valid, edits: [{ ...valid.edits[0], expected: 'wrong' }] }, 'text-edits-baseline-mismatch'],
      [{ ...valid, edits: [{ ...valid.edits[0], elementId: 'missing' }] }, 'text-edits-target-unbound'],
      [{ ...valid, edits: [valid.edits[0], { ...valid.edits[0], elementId: 'missing' }] }, 'text-edits-target-unbound'],
      [{ ...valid, edits: [valid.edits[0], valid.edits[0]] }, 'text-edits-invalid'],
      [{ ...valid, extra: true }, 'text-edits-invalid'],
      [{ ...valid, edits: [{ ...valid.edits[0], replacement: '\u0001' }] }, 'save-text-candidate-unsupported'],
      [{ ...valid, edits: [{ ...valid.edits[0], replacement: valid.edits[0].expected }] }, 'save-edit-no-changes'],
    ];
    for (const [manifest, code] of cases) {
      await writeFile(path.join(f.directory, 'bad-edits.json'), JSON.stringify(manifest));
      const result = await runCliAsync(['save', 'source.pptx', 'rejected.pptx', '--mode', 'edit', '--text-edits', 'bad-edits.json', '--json'], f.options);
      assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'rejected.pptx')), { code: 'ENOENT' });
    }
    for (const args of [['save', 'source.pptx', 'x', '--mode', 'no-op', '--text-edits', 'edits.json'], ['inspect', 'source.pptx', '--text-edits', 'edits.json']]) assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
    const abort = new AbortController(); abort.abort(); assert.equal((await runCliAsync(['save', 'source.pptx', 'cancelled.pptx', '--mode', 'edit', '--text-edits', 'edits.json'], { ...f.options, signal: abort.signal })).exitCode, 130); await assert.rejects(readFile(path.join(f.directory, 'cancelled.pptx')), { code: 'ENOENT' });
  }
});

test('one text manifest updates distinct repeated runs without changing their other XML', async t => {
  const paragraphs = '<a:p><a:r><a:rPr b="1"/><a:t>same</a:t></a:r><a:r><a:rPr i="1"/><a:t>same</a:t></a:r></a:p>';
  const fixtureData = contentFixture(shape({ txBox: true, text: body(paragraphs) })), source = encodeStructure(fixtureData), f = await fixture(t, source);
  const imported = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v1' });
  const edits = imported.report.textRuns.map((binding, i) => ({ slideId: binding.slideId, elementId: binding.elementId, paragraphIndex: binding.paragraphIndex, runIndex: binding.runIndex, expected: 'same', replacement: `new${i}` })); imported.dispose();
  await writeFile(path.join(f.directory, 'batch.json'), JSON.stringify({ version: 'pptkit-text-edits-v1', sourceSha256: sha(source), edits }));
  const result = await runCliAsync(['save', 'source.pptx', 'batch.pptx', '--mode', 'edit', '--text-edits', 'batch.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).report.changes.length, 2);
  const actual = readGeneratedZip(await readFile(path.join(f.directory, 'batch.pptx'))).get('deck/z.xml').toString(); assert.equal(actual, fixtureData.parts['deck/z.xml'].replace('>same</a:t>', '>new0</a:t>').replace('>same</a:t>', '>new1</a:t>'));
});

test('text manifest does not bypass Parser range guards', async t => {
  const fixtureData = contentFixture(); fixtureData.parts['deck/z.xml'] = fixtureData.parts['deck/z.xml'].replace('</p:sld>', '<p:timing><p:spTgt spid="2"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt></p:timing></p:sld>');
  const source = encodeStructure(fixtureData), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v1' }); const binding = session.report.textRuns[0]; session.dispose();
  await writeFile(path.join(f.directory, 'edits.json'), JSON.stringify({ version: 'pptkit-text-edits-v1', sourceSha256: sha(source), edits: [{ slideId: binding.slideId, elementId: binding.elementId, paragraphIndex: binding.paragraphIndex, runIndex: binding.runIndex, expected: ' A & B ', replacement: 'changed' }] }));
  const result = await runCliAsync(['save', 'source.pptx', 'blocked.pptx', '--mode', 'edit', '--text-edits', 'edits.json', '--json'], f.options); assert.equal(result.exitCode, 3); assert.equal(decode(result).error.code, 'save-text-range-reference'); await assert.rejects(readFile(path.join(f.directory, 'blocked.pptx')), { code: 'ENOENT' }); assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
  const planned = await runCliAsync(['save', 'source.pptx', 'blocked.pptx', '--mode', 'edit', '--text-edits', 'edits.json', '--dry-run', '--json'], f.options); assert.equal(planned.exitCode, 1); assert.equal(decode(planned).status, 'blocked'); assert.ok(decode(planned).report.diagnostics.some(d => d.code === 'save-text-range-reference')); assert.ok(decode(planned).report.changes.some(change => change.kind === 'text')); await assert.rejects(readFile(path.join(f.directory, 'blocked.pptx')), { code: 'ENOENT' });
});

test('grouped text manifests update separate nested owners and retain the source tree', async t => {
  for (const strict of [false, true]) {
    const textA = shape({ id: '2', txBox: true, text: body() }), textB = shape({ id: '3', txBox: true, text: body() });
    const fixtureData = contentFixture(group(textA + group(textB, 11), 10), strict), source = encodeStructure(fixtureData), f = await fixture(t, source);
    const imported = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v1' });
    const bindings = imported.report.textRuns.filter(b => b.status === 'bound'); assert.equal(bindings.length, 2); imported.dispose();
    const edits = bindings.map((b, i) => ({ slideId: b.slideId, elementId: b.elementId, paragraphIndex: b.paragraphIndex, runIndex: b.runIndex, expected: ' A & B ', replacement: `Nested ${i}` }));
    const manifest = { version: 'pptkit-text-edits-v1', sourceSha256: sha(source), edits }; await writeFile(path.join(f.directory, 'group-edits.json'), JSON.stringify(manifest));
    const result = await runCliAsync(['save', 'source.pptx', 'grouped.pptx', '--mode', 'edit', '--text-edits', 'group-edits.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output);
    const output = await readFile(path.join(f.directory, 'grouped.pptx')), before = readGeneratedZip(source), after = readGeneratedZip(output);
    for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
    assert.equal(after.get('deck/z.xml').toString(), fixtureData.parts['deck/z.xml'].replace('> A &amp; B </a:t>', '>Nested 0</a:t>').replace('> A &amp; B </a:t>', '>Nested 1</a:t>'));
    const reopened = await importPptx(output, { projection: 'content', editSaveProfile: 'text-runs-v1' }); const outer = reopened.document.slides[0].elements[0]; assert.equal(outer.type, 'group'); assert.equal(outer.children[0].content[0].runs[0].text, 'Nested 0'); assert.equal(outer.children[1].children[0].content[0].runs[0].text, 'Nested 1'); reopened.dispose();
    let current = output, inputFile = 'grouped.pptx';
    for (let round = 2; round <= 3; round++) {
      const nextEdits = edits.map((edit, i) => ({ ...edit, expected: round === 2 ? `Nested ${i}` : `Round ${round - 1}-${i}`, replacement: `Round ${round}-${i}` }));
      await writeFile(path.join(f.directory, 'next-group-edits.json'), JSON.stringify({ ...manifest, sourceSha256: sha(current), edits: nextEdits }));
      const outputFile = `grouped-${round}.pptx`, next = await runCliAsync(['save', inputFile, outputFile, '--mode', 'edit', '--text-edits', 'next-group-edits.json', '--json'], f.options); assert.equal(next.exitCode, 0, next.output);
      const nextBytes = await readFile(path.join(f.directory, outputFile)), oldParts = readGeneratedZip(current), newParts = readGeneratedZip(nextBytes);
      for (const [name, payload] of oldParts) if (name !== 'deck/z.xml') assert.deepEqual(newParts.get(name), payload, name);
      let expectedXml = oldParts.get('deck/z.xml').toString(); for (const edit of nextEdits) expectedXml = expectedXml.replace(`>${edit.expected}</a:t>`, `>${edit.replacement}</a:t>`);
      assert.equal(newParts.get('deck/z.xml').toString(), expectedXml); current = nextBytes; inputFile = outputFile;
    }
    const finalSession = await importPptx(current, { projection: 'content', editSaveProfile: 'text-runs-v1' }); const finalGroup = finalSession.document.slides[0].elements[0]; assert.equal(finalGroup.children[0].content[0].runs[0].text, 'Round 3-0'); assert.equal(finalGroup.children[1].children[0].content[0].runs[0].text, 'Round 3-1'); finalSession.dispose();
    manifest.edits[1].expected = 'wrong'; await writeFile(path.join(f.directory, 'bad-group-edits.json'), JSON.stringify(manifest));
    const rejected = await runCliAsync(['save', 'source.pptx', 'no-group.pptx', '--mode', 'edit', '--text-edits', 'bad-group-edits.json', '--json'], f.options); assert.equal(rejected.exitCode, 3); assert.equal(decode(rejected).error.code, 'text-edits-baseline-mismatch'); await assert.rejects(readFile(path.join(f.directory, 'no-group.pptx')), { code: 'ENOENT' }); assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
  }
});

test('v2 manifest preserves mixed shape, grouped text-box and notes edits across three rounds', async t => {
  const textBody = text => body(`<a:p><a:r><a:rPr b="1"/><a:t>${text}</a:t></a:r></a:p>`);
  for (const strict of [false, true]) {
    const fixtureData = notesFixture(notesShape({ text: textBody('Notes old') }), strict);
    fixtureData.parts['deck/z.xml'] = fixtureData.parts['deck/z.xml'].replace(shape({ txBox: true, text: body() }), shape({ text: textBody('Shape old'), preset: 'roundRect' }) + group(shape({ id: '3', txBox: true, text: textBody('Box old') }), 10));
    const source = encodeStructure(fixtureData), f = await fixture(t, source); let current = source, inputFile = 'source.pptx';
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(current, { projection: 'content', editSaveProfile: 'text-runs-v2' });
      const draw = session.report.textRuns.filter(b => b.status === 'bound'), notes = session.report.notesTextRuns[0]; assert.equal(draw.length, 2); assert.equal(notes.status, 'bound');
      const edits = draw.map(b => ({ owner: b.owner, slideId: b.slideId, elementId: b.elementId, paragraphIndex: b.paragraphIndex, runIndex: b.runIndex, expected: round === 1 ? (b.owner === 'shape-text' ? 'Shape old' : 'Box old') : `${b.owner} ${round - 1}`, replacement: `${b.owner} ${round}` }));
      edits.push({ owner: 'notes', slideId: notes.slideId, paragraphIndex: notes.paragraphIndex, runIndex: notes.runIndex, expected: round === 1 ? 'Notes old' : `notes ${round - 1}`, replacement: `notes ${round}` }); session.dispose();
      await writeFile(path.join(f.directory, 'v2.json'), JSON.stringify({ version: 'pptkit-text-edits-v2', sourceSha256: sha(current), edits }));
      const outputFile = `v2-${round}.pptx`, result = await runCliAsync(['save', inputFile, outputFile, '--mode', 'edit', '--text-edits', 'v2.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).report.changes.length, 3);
      const next = await readFile(path.join(f.directory, outputFile)), before = readGeneratedZip(current), after = readGeneratedZip(next);
      for (const [name, bytes] of before) if (!['deck/z.xml', 'deck/notes/z.xml'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
      let slideXml = before.get('deck/z.xml').toString(); for (const edit of edits.filter(e => e.owner !== 'notes')) slideXml = slideXml.replace(`>${edit.expected}</a:t>`, `>${edit.replacement}</a:t>`);
      assert.equal(after.get('deck/z.xml').toString(), slideXml); assert.equal(after.get('deck/notes/z.xml').toString(), before.get('deck/notes/z.xml').toString().replace(`>${edits[2].expected}</a:t>`, `>${edits[2].replacement}</a:t>`));
      current = next; inputFile = outputFile;
    }
    const final = await importPptx(current, { projection: 'content', editSaveProfile: 'text-runs-v2' }); const slide = final.document.slides[0]; assert.equal(slide.notes[0].runs[0].text, 'notes 3'); assert.equal(slide.elements[0].text.content[0].runs[0].text, 'shape-text 3'); assert.equal(slide.elements[1].children[0].content[0].runs[0].text, 'text-box 3'); final.dispose();
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
  }
});

test('v2 text manifest rejects mismatched owners, unmapped notes and conflicting mixed batches', async t => {
  const fixtureData = notesFixture(notesShape({ text: body('<a:p><a:r><a:t>Note <!--opaque-->value</a:t></a:r></a:p>') }));
  fixtureData.parts['deck/z.xml'] = fixtureData.parts['deck/z.xml'].replace(shape({ txBox: true, text: body() }), shape({ text: body(), preset: 'roundRect' }));
  const source = encodeStructure(fixtureData), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v2' });
  const draw = session.report.textRuns[0], note = session.report.notesTextRuns[0]; assert.equal(draw.owner, 'shape-text'); assert.equal(note.status, 'unmapped'); session.dispose();
  const shapeEdit = { owner: 'shape-text', slideId: draw.slideId, elementId: draw.elementId, paragraphIndex: draw.paragraphIndex, runIndex: draw.runIndex, expected: ' A & B ', replacement: 'changed' };
  const noteEdit = { owner: 'notes', slideId: note.slideId, paragraphIndex: note.paragraphIndex, runIndex: note.runIndex, expected: 'Note value', replacement: 'changed' };
  const cases = [
    [{ ...shapeEdit, owner: 'text-box' }, 'text-edits-target-unbound'],
    [noteEdit, 'text-edits-target-unbound'],
    [[shapeEdit, noteEdit], 'text-edits-target-unbound'],
    [{ ...noteEdit, elementId: draw.elementId }, 'text-edits-invalid'],
    [{ ...shapeEdit, owner: ['shape-text'] }, 'text-edits-invalid'],
    [{ ...shapeEdit, owner: undefined }, 'text-edits-invalid'],
  ];
  for (const [candidate, code] of cases) {
    await writeFile(path.join(f.directory, 'bad-v2.json'), JSON.stringify({ version: 'pptkit-text-edits-v2', sourceSha256: sha(source), edits: Array.isArray(candidate) ? candidate : [candidate] }));
    const result = await runCliAsync(['save', 'source.pptx', 'no-v2.pptx', '--mode', 'edit', '--text-edits', 'bad-v2.json', '--json'], f.options); assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'no-v2.pptx')), { code: 'ENOENT' });
  }
  assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
});

test('v3 text manifest binds merged table anchors by compact and physical indices', async t => {
  for (const strict of [false, true]) {
    const merged = table({ columns: [30, 30, 40], rows: [tableRow([cell({ text: 'Same', attributes: 'gridSpan="2" rowSpan="2"' }), cell({ body: '', attributes: 'hMerge="1"' }), cell({ text: 'Top' })]), tableRow([cell({ body: '', attributes: 'vMerge="1"' }), cell({ body: '', attributes: 'vMerge="1" hMerge="1"' }), cell({ text: 'Same' })], 30)] });
    const fixtureData = tableFixture(merged + shape({ id: '3', txBox: true, text: body('<a:p><a:r><a:t>Outside</a:t></a:r></a:p>') }), strict), source = encodeStructure(fixtureData), f = await fixture(t, source); let current = source, inputFile = 'source.pptx';
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(current, { projection: 'content', editSaveProfile: 'text-runs-v3' });
      const anchors = session.report.tableTextRuns.filter(b => b.rowIndex === 0 && b.cellIndex === 0 || b.rowIndex === 1 && b.cellIndex === 0), drawing = session.report.textRuns[0];
      assert.deepEqual(anchors.map(b => [b.rowIndex, b.cellIndex, b.columnIndex]), [[0, 0, 0], [1, 0, 2]]);
      const edits = anchors.map((b, i) => ({ owner: 'table-cell', slideId: b.slideId, elementId: b.elementId, rowIndex: b.rowIndex, cellIndex: b.cellIndex, columnIndex: b.columnIndex, paragraphIndex: b.paragraphIndex, runIndex: b.runIndex, expected: round === 1 ? 'Same' : `Cell ${round - 1}-${i}`, replacement: `Cell ${round}-${i}` }));
      edits.push({ owner: 'text-box', slideId: drawing.slideId, elementId: drawing.elementId, paragraphIndex: drawing.paragraphIndex, runIndex: drawing.runIndex, expected: round === 1 ? 'Outside' : `Outside ${round - 1}`, replacement: `Outside ${round}` }); session.dispose();
      const manifest = { version: 'pptkit-text-edits-v3', sourceSha256: sha(current), edits }; await writeFile(path.join(f.directory, 'v3.json'), JSON.stringify(manifest));
      const outputFile = `v3-${round}.pptx`, result = await runCliAsync(['save', inputFile, outputFile, '--mode', 'edit', '--text-edits', 'v3.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).report.changes.length, 3);
      const next = await readFile(path.join(f.directory, outputFile)), before = readGeneratedZip(current), after = readGeneratedZip(next);
      for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
      let expected = before.get('deck/z.xml').toString(); for (const edit of edits) expected = expected.replace(`>${edit.expected}</a:t>`, `>${edit.replacement}</a:t>`);
      assert.equal(after.get('deck/z.xml').toString(), expected); current = next; inputFile = outputFile;
      if (round === 1) {
        const bad = { ...manifest, edits: [edits[2], { ...edits[1], columnIndex: 0 }] }; await writeFile(path.join(f.directory, 'v3-bad.json'), JSON.stringify(bad));
        const rejected = await runCliAsync(['save', 'source.pptx', 'no-v3.pptx', '--mode', 'edit', '--text-edits', 'v3-bad.json', '--json'], f.options); assert.equal(rejected.exitCode, 3); assert.equal(decode(rejected).error.code, 'text-edits-target-unbound'); await assert.rejects(readFile(path.join(f.directory, 'no-v3.pptx')), { code: 'ENOENT' });
      }
    }
    const final = await importPptx(current, { projection: 'content', editSaveProfile: 'text-runs-v3' }); const cells = final.document.slides[0].elements[0].rows; assert.equal(cells[0].cells[0].content[0].runs[0].text, 'Cell 3-0'); assert.equal(cells[1].cells[0].content[0].runs[0].text, 'Cell 3-1'); final.dispose();
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
  }
});

test('v3 grouped tables reject covered cells, duplicate targets and older-profile escalation', async t => {
  const merged = table({ columns: [30, 30, 40], rows: [tableRow([cell({ text: 'Anchor', attributes: 'gridSpan="2" rowSpan="2"' }), cell({ body: '', attributes: 'hMerge="1"' }), cell({ text: 'Top' })]), tableRow([cell({ body: '', attributes: 'vMerge="1"' }), cell({ body: '', attributes: 'vMerge="1" hMerge="1"' }), cell({ text: 'Bottom' })], 30)] });
  const fixtureData = tableFixture(group(merged, 10)), source = encodeStructure(fixtureData), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v3' });
  const b = session.report.tableTextRuns.find(item => item.rowIndex === 1); assert.deepEqual([b.cellIndex, b.columnIndex], [0, 2]); session.dispose();
  const edit = { owner: 'table-cell', slideId: b.slideId, elementId: b.elementId, rowIndex: b.rowIndex, cellIndex: b.cellIndex, columnIndex: b.columnIndex, paragraphIndex: b.paragraphIndex, runIndex: b.runIndex, expected: 'Bottom', replacement: 'Changed' };
  const cases = [
    [{ ...edit, columnIndex: 0 }, 'pptkit-text-edits-v3', 'text-edits-target-unbound'],
    [{ ...edit, cellIndex: 1 }, 'pptkit-text-edits-v3', 'text-edits-target-unbound'],
    [{ ...edit, rowIndex: -1 }, 'pptkit-text-edits-v3', 'text-edits-invalid'],
    [{ ...edit, columnIndex: undefined }, 'pptkit-text-edits-v3', 'text-edits-invalid'],
    [edit, 'pptkit-text-edits-v2', 'text-edits-invalid'],
    [[edit, edit], 'pptkit-text-edits-v3', 'text-edits-invalid'],
  ];
  for (const [candidate, version, code] of cases) {
    await writeFile(path.join(f.directory, 'bad-group-table.json'), JSON.stringify({ version, sourceSha256: sha(source), edits: Array.isArray(candidate) ? candidate : [candidate] }));
    const result = await runCliAsync(['save', 'source.pptx', 'bad-table.pptx', '--mode', 'edit', '--text-edits', 'bad-group-table.json', '--json'], f.options); assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad-table.pptx')), { code: 'ENOENT' });
  }
  await writeFile(path.join(f.directory, 'good-group-table.json'), JSON.stringify({ version: 'pptkit-text-edits-v3', sourceSha256: sha(source), edits: [edit] }));
  const saved = await runCliAsync(['save', 'source.pptx', 'group-table.pptx', '--mode', 'edit', '--text-edits', 'good-group-table.json', '--json'], f.options); assert.equal(saved.exitCode, 0, saved.output);
  assert.equal(readGeneratedZip(await readFile(path.join(f.directory, 'group-table.pptx'))).get('deck/z.xml').toString(), fixtureData.parts['deck/z.xml'].replace('>Bottom</a:t>', '>Changed</a:t>')); assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
});

test('v3 table text manifest respects source range guards', async t => {
  const fixtureData = tableFixture(); fixtureData.parts['deck/z.xml'] = fixtureData.parts['deck/z.xml'].replace('</p:sld>', '<p:timing><p:spTgt spid="2"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt></p:timing></p:sld>');
  const source = encodeStructure(fixtureData), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v3' }); const b = session.report.tableTextRuns[0]; session.dispose();
  const edit = { owner: 'table-cell', slideId: b.slideId, elementId: b.elementId, rowIndex: b.rowIndex, cellIndex: b.cellIndex, columnIndex: b.columnIndex, paragraphIndex: b.paragraphIndex, runIndex: b.runIndex, expected: 'R1C1', replacement: 'changed' };
  await writeFile(path.join(f.directory, 'guarded.json'), JSON.stringify({ version: 'pptkit-text-edits-v3', sourceSha256: sha(source), edits: [edit] }));
  const result = await runCliAsync(['save', 'source.pptx', 'guarded.pptx', '--mode', 'edit', '--text-edits', 'guarded.json', '--json'], f.options); assert.equal(result.exitCode, 3); assert.equal(decode(result).error.code, 'save-text-range-reference'); await assert.rejects(readFile(path.join(f.directory, 'guarded.pptx')), { code: 'ENOENT' });
});

test('v3 manifest commits table, shape and notes text in one guarded plan', async t => {
  const fixtureData = notesFixture(notesShape({ text: body('<a:p><a:r><a:t>Note</a:t></a:r></a:p>') }));
  fixtureData.parts['deck/z.xml'] = fixtureData.parts['deck/z.xml'].replace(shape({ txBox: true, text: body() }), table() + shape({ id: '3', text: body(), preset: 'rect' }));
  const source = encodeStructure(fixtureData), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v3' });
  const c = session.report.tableTextRuns[0], s = session.report.textRuns.find(b => b.owner === 'shape-text'), n = session.report.notesTextRuns[0]; assert.ok(c && s && n); session.dispose();
  const edits = [
    { owner: 'table-cell', slideId: c.slideId, elementId: c.elementId, rowIndex: c.rowIndex, cellIndex: c.cellIndex, columnIndex: c.columnIndex, paragraphIndex: c.paragraphIndex, runIndex: c.runIndex, expected: 'R1C1', replacement: 'Table changed' },
    { owner: 'shape-text', slideId: s.slideId, elementId: s.elementId, paragraphIndex: s.paragraphIndex, runIndex: s.runIndex, expected: ' A & B ', replacement: 'Shape changed' },
    { owner: 'notes', slideId: n.slideId, paragraphIndex: n.paragraphIndex, runIndex: n.runIndex, expected: 'Note', replacement: 'Note changed' },
  ];
  await writeFile(path.join(f.directory, 'all-owners.json'), JSON.stringify({ version: 'pptkit-text-edits-v3', sourceSha256: sha(source), edits }));
  const result = await runCliAsync(['save', 'source.pptx', 'all-owners.pptx', '--mode', 'edit', '--text-edits', 'all-owners.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).report.changes.length, 3);
  const after = readGeneratedZip(await readFile(path.join(f.directory, 'all-owners.pptx'))), before = readGeneratedZip(source);
  assert.equal(after.get('deck/z.xml').toString(), before.get('deck/z.xml').toString().replace('>R1C1</a:t>', '>Table changed</a:t>').replace('> A &amp; B </a:t>', '>Shape changed</a:t>'));
  assert.equal(after.get('deck/notes/z.xml').toString(), before.get('deck/notes/z.xml').toString().replace('>Note</a:t>', '>Note changed</a:t>'));
  for (const [name, bytes] of before) if (!['deck/z.xml', 'deck/notes/z.xml'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
});

test('opt-in inspect text values build source-bound manifests for every profile', async t => {
  for (const strict of [false, true]) {
    const fixtureData = notesFixture(notesShape({ text: body('<a:p><a:r><a:t>Note</a:t></a:r></a:p>') }), strict);
    fixtureData.parts['deck/z.xml'] = fixtureData.parts['deck/z.xml'].replace(shape({ txBox: true, text: body() }), table() + group(shape({ id: '3', text: body(), preset: 'rect' }) + shape({ id: '4', txBox: true, text: body() }), 10));
    const source = encodeStructure(fixtureData), f = await fixture(t, source);
    for (const version of [1, 2, 3]) {
      const args = ['inspect', 'source.pptx', '--projection', 'content', '--edit-save-profile', `text-runs-v${version}`, '--json'];
      const ordinary = await runCliAsync(args, f.options); assert.equal(ordinary.exitCode, 0, ordinary.output); assert.equal(decode(ordinary).textEditCandidates, undefined);
      const inspected = await runCliAsync([...args, '--text-values'], f.options); assert.equal(inspected.exitCode, 0, inspected.output); const candidates = decode(inspected).textEditCandidates;
      assert.equal(candidates.version, `pptkit-text-edits-v${version}`); assert.equal(candidates.sourceSha256, sha(source)); assert.equal(candidates.unboundCount, 0);
      assert.ok(candidates.targets.every(target => typeof target.expected === 'string' && !Object.hasOwn(target, 'replacement')));
      if (version === 1) assert.ok(candidates.targets.every(target => !Object.hasOwn(target, 'owner')));
      else assert.ok(candidates.targets.some(target => target.owner === 'shape-text') && candidates.targets.some(target => target.owner === 'notes' && !Object.hasOwn(target, 'elementId')));
      if (version === 3) assert.ok(candidates.targets.some(target => target.owner === 'table-cell' && Number.isInteger(target.columnIndex)));
      const selected = candidates.targets.map(target => ({ ...target, replacement: `${target.expected}!` })); await writeFile(path.join(f.directory, `candidate-${version}.json`), JSON.stringify({ version: candidates.version, sourceSha256: candidates.sourceSha256, edits: selected }));
      const saved = await runCliAsync(['save', 'source.pptx', `candidate-${version}.pptx`, '--mode', 'edit', '--text-edits', `candidate-${version}.json`, '--json'], f.options); assert.equal(saved.exitCode, 0, saved.output); assert.equal(decode(saved).report.changes.length, selected.length);
    }
    for (const args of [['inspect', 'source.pptx', '--text-values', '--json'], ['inspect', 'source.pptx', '--projection', 'content', '--edit-save-profile', 'text-runs-v1', '--text-values'], ['inspect', 'source.pptx', '--projection', 'content', '--edit-save-profile', 'text-runs-v1', '--text-values', '--capabilities', '--json'], ['compare', 'source.pptx', 'source.pptx', '--text-values', '--json']]) assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
    const limited = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--edit-save-profile', 'text-runs-v3', '--text-values', '--json', '--max-output-bytes', '512'], f.options); assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
  }
});

test('inspect text values reports unbound atoms without presenting them as editable targets', async t => {
  const fixtureData = contentFixture(shape({ txBox: true, text: body('<a:p><a:r><a:t>mixed<!--opaque-->text</a:t></a:r><a:r><a:t>bound</a:t></a:r></a:p>') })), f = await fixture(t, encodeStructure(fixtureData));
  const result = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--edit-save-profile', 'text-runs-v1', '--text-values', '--json'], f.options); assert.equal(result.exitCode, 0, result.output);
  const candidates = decode(result).textEditCandidates; assert.equal(candidates.unboundCount, 1); assert.deepEqual(candidates.targets.map(t => t.expected), ['bound']); assert.equal(candidates.targets[0].runIndex, 1);
});

test('text candidate projection rejects ambiguous or inconsistent bound identities', () => {
  const binding = { slideId: 's', elementId: 'e', paragraphIndex: 0, runIndex: 0, status: 'bound' };
  const session = { document: { slides: [{ id: 's', elements: [{ id: 'e', type: 'text', content: [{ runs: [{ text: 'before' }] }] }] }] }, report: { textRuns: [binding] } };
  assert.equal(textCandidates(session, 'text-runs-v1', 'hash').targets[0].expected, 'before');
  session.report.textRuns = [binding, { ...binding, status: 'unmapped' }]; assert.throws(() => textCandidates(session, 'text-runs-v1', 'hash'), e => e.code === 'text-candidates-inconsistent');
  session.report.textRuns = [binding]; session.document.slides[0].elements[0].content = []; assert.throws(() => textCandidates(session, 'text-runs-v1', 'hash'), e => e.code === 'text-candidates-inconsistent');
  const controller = new AbortController(); controller.abort(); assert.throws(() => textCandidates(session, 'text-runs-v1', 'hash', controller.signal), e => e.code === 'operation-aborted');
});

test('chart manifest synchronizes cache and workbook in three rounds across families and dialects', async t => {
  for (const strict of [false, true]) for (const family of ['bar', 'line', 'pie']) {
    const initial = family === 'pie' ? 2.5 : -2.5, original = encodeChart(chartFixture({ strict, xml: chartXml({ family, data: series({ refs: true, values: [0, initial] }), external: true }), book: workbook({ values: [0, initial] }) })), f = await fixture(t, original); let current = original, input = 'source.pptx';
    for (const [i, value] of ['12.25', '-0', '37'].entries()) {
      const inspect = await runCliAsync(['inspect', input, '--projection', 'content', '--chart-edit-save-profile', 'existing-values-v1', '--chart-values', '--json'], f.options); assert.equal(inspect.exitCode, 0, inspect.output);
      const candidates = decode(inspect).chartEditCandidates, target = candidates.targets.find(x => x.seriesIndex === 0 && x.pointIndex === 1); assert.equal(candidates.sourceSha256, sha(current)); assert.ok(target, `${family}/${strict}/${i}: ${JSON.stringify(candidates)} ${JSON.stringify(decode(inspect).report.chartValues)}`);
      assert.equal(target.expected, i === 0 ? String(initial) : ['12.25', '-0'][i - 1]);
      await writeFile(path.join(f.directory, 'chart-edits.json'), JSON.stringify({ version: candidates.version, sourceSha256: candidates.sourceSha256, edits: [{ ...target, replacement: value }] }));
      const output = `chart-${i}.pptx`, result = await runCliAsync(['save', input, output, '--mode', 'edit', '--chart-edits', 'chart-edits.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).report.assurance, 'scoped-preservation');
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(current), after = readGeneratedZip(next);
      for (const [name, bytes] of before) if (!['deck/chart.xml', 'deck/book.xlsx'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
      assert.equal(after.get('deck/chart.xml').toString(), before.get('deck/chart.xml').toString().replace(/(<c:val>[\s\S]*?<c:pt idx="1"><c:v>)[^<]*(<\/c:v>)/, `$1${value}$2`));
      const previousBook = readGeneratedZip(before.get('deck/book.xlsx')), nextBook = readGeneratedZip(after.get('deck/book.xlsx'));
      for (const [name, bytes] of previousBook) if (name !== 'xl/sheets/one.xml') assert.deepEqual(nextBook.get(name), bytes, name);
      assert.equal(nextBook.get('xl/sheets/one.xml').toString(), previousBook.get('xl/sheets/one.xml').toString().replace(/(<c r="B3"[^>]*><v>)[^<]*(<\/v>)/, `$1${value}$2`));
      const reopened = await importPptx(next, { projection: 'content', chartEditSaveProfile: 'existing-values-v1' }); assert.ok(Object.is(reopened.document.slides[0].elements[0].series[0].values[1], Number(value))); reopened.dispose(); current = next; input = output;
    }
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), original);
  }
});

test('chart, text and starting number can share one guarded save and reject mixed failures', async t => {
  for (const strict of [false, true]) {
    const fdata = chartFixture({ strict, xml: chartXml({ family: 'bar', data: series({ refs: true, values: [0, -2.5] }), external: true }), book: workbook({ values: [0, -2.5] }), content: frame() + shape({ id: '20', txBox: true, text: body('<a:p><a:r><a:t>Text old</a:t></a:r></a:p>') }) }), source = encodeChart(fdata), f = await fixture(t, source);
    const inspection = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--edit-save-profile', 'text-runs-v3', '--chart-edit-save-profile', 'existing-values-v1', '--chart-values', '--text-values', '--json'], f.options); assert.equal(inspection.exitCode, 0, inspection.output);
    const inspected = decode(inspection), chart = inspected.chartEditCandidates.targets.find(c => c.pointIndex === 1), text = inspected.textEditCandidates.targets[0]; assert.equal(chart.expected, '-2.5'); assert.equal(text.expected, 'Text old');
    const chartManifest = { version: 'pptkit-chart-edits-v1', sourceSha256: sha(source), edits: [{ ...chart, replacement: '9' }] }, textManifest = { version: 'pptkit-text-edits-v3', sourceSha256: sha(source), edits: [{ ...text, replacement: 'Text new' }] };
    await writeFile(path.join(f.directory, 'chart.json'), JSON.stringify(chartManifest)); await writeFile(path.join(f.directory, 'text.json'), JSON.stringify(textManifest));
    const args = ['save', 'source.pptx', 'combined.pptx', '--mode', 'edit', '--chart-edits', 'chart.json', '--text-edits', 'text.json', '--first-slide-number', '17', '--json'];
    const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).report.changes.length, 3); assert.equal(decode(result).inputs.length, 3);
    const output = await readFile(path.join(f.directory, 'combined.pptx')), before = readGeneratedZip(source), after = readGeneratedZip(output);
    for (const [name, payload] of before) if (!['deck/main.data', 'deck/z.xml', 'deck/chart.xml', 'deck/book.xlsx'].includes(name)) assert.deepEqual(after.get(name), payload, name);
    assert.equal(after.get('deck/z.xml').toString(), before.get('deck/z.xml').toString().replace('>Text old</a:t>', '>Text new</a:t>'));
    assert.match(after.get('deck/chart.xml').toString(), /<c:pt idx="1"><c:v>9<\/c:v>/); assert.match(readGeneratedZip(after.get('deck/book.xlsx')).get('xl/sheets/one.xml').toString(), /<c r="B3"[^>]*><v>9<\/v>/);
    const reopened = await importPptx(output, { projection: 'content', editSaveProfile: 'text-runs-v3', chartEditSaveProfile: 'existing-values-v1', firstSlideNumberSaveProfile: 'attribute-v2' }); assert.equal(reopened.document.firstSlideNumber, 17); assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], 9); assert.equal(reopened.document.slides[0].elements[1].content[0].runs[0].text, 'Text new'); reopened.dispose();
    let current = output, input = 'combined.pptx';
    for (const [round, number] of [[2, '0'], [3, 'clear']]) {
      const nextInspection = await runCliAsync(['inspect', input, '--projection', 'content', '--edit-save-profile', 'text-runs-v3', '--chart-edit-save-profile', 'existing-values-v1', '--chart-values', '--text-values', '--json'], f.options); assert.equal(nextInspection.exitCode, 0, nextInspection.output);
      const nextReport = decode(nextInspection), nextChart = nextReport.chartEditCandidates.targets.find(c => c.pointIndex === 1), nextText = nextReport.textEditCandidates.targets[0];
      await writeFile(path.join(f.directory, 'chart-next.json'), JSON.stringify({ version: nextReport.chartEditCandidates.version, sourceSha256: sha(current), edits: [{ ...nextChart, replacement: String(round + 10) }] }));
      await writeFile(path.join(f.directory, 'text-next.json'), JSON.stringify({ version: nextReport.textEditCandidates.version, sourceSha256: sha(current), edits: [{ ...nextText, replacement: `Text ${round}` }] }));
      const nextFile = `combined-${round}.pptx`, nextResult = await runCliAsync(['save', input, nextFile, '--mode', 'edit', '--chart-edits', 'chart-next.json', '--text-edits', 'text-next.json', '--first-slide-number', number, '--json'], f.options); assert.equal(nextResult.exitCode, 0, nextResult.output); assert.equal(decode(nextResult).report.changes.length, 3);
      const nextBytes = await readFile(path.join(f.directory, nextFile)), oldParts = readGeneratedZip(current), newParts = readGeneratedZip(nextBytes);
      for (const [name, payload] of oldParts) if (!['deck/main.data', 'deck/z.xml', 'deck/chart.xml', 'deck/book.xlsx'].includes(name)) assert.deepEqual(newParts.get(name), payload, name);
      assert.equal(newParts.get('deck/z.xml').toString(), oldParts.get('deck/z.xml').toString().replace(`>${nextText.expected}</a:t>`, `>Text ${round}</a:t>`));
      assert.match(newParts.get('deck/chart.xml').toString(), new RegExp(`<c:pt idx="1"><c:v>${round + 10}</c:v>`)); assert.match(readGeneratedZip(newParts.get('deck/book.xlsx')).get('xl/sheets/one.xml').toString(), new RegExp(`<c r="B3"[^>]*><v>${round + 10}</v>`));
      const nextSession = await importPptx(nextBytes, { projection: 'content', editSaveProfile: 'text-runs-v3', chartEditSaveProfile: 'existing-values-v1', firstSlideNumberSaveProfile: 'attribute-v2' }); assert.equal(nextSession.document.firstSlideNumber, number === 'clear' ? undefined : Number(number)); assert.equal(nextSession.document.slides[0].elements[0].series[0].values[1], round + 10); nextSession.dispose(); current = nextBytes; input = nextFile;
    }
    chartManifest.edits[0].expected = '0'; await writeFile(path.join(f.directory, 'chart.json'), JSON.stringify(chartManifest));
    const rejected = await runCliAsync(['save', 'source.pptx', 'rejected.pptx', '--mode', 'edit', '--chart-edits', 'chart.json', '--text-edits', 'text.json', '--first-slide-number', '17', '--json'], f.options); assert.equal(rejected.exitCode, 3); assert.equal(decode(rejected).error.code, 'chart-edits-baseline-mismatch'); await assert.rejects(readFile(path.join(f.directory, 'rejected.pptx')), { code: 'ENOENT' }); assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
  }
});

test('chart edit manifests reject stale, ambiguous and noncanonical numeric inputs', async t => {
  const source = encodeChart(chartFixture({ xml: chartXml({ family: 'bar', data: series({ refs: true, values: [0, -2.5] }), external: true }), book: workbook({ values: [0, -2.5] }) })), f = await fixture(t, source);
  const inspect = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--chart-edit-save-profile', 'existing-values-v1', '--chart-values', '--json'], f.options), target = decode(inspect).chartEditCandidates.targets.find(x => x.pointIndex === 1);
  const edit = { ...target, replacement: '9' }, manifest = { version: 'pptkit-chart-edits-v1', sourceSha256: sha(source), edits: [edit] };
  const cases = [
    [{ ...manifest, sourceSha256: '0'.repeat(64) }, 'chart-edits-source-mismatch'],
    [{ ...manifest, edits: [{ ...edit, expected: '0' }] }, 'chart-edits-baseline-mismatch'],
    [{ ...manifest, edits: [{ ...edit, elementId: 'missing' }] }, 'chart-edits-target-unbound'],
    [{ ...manifest, edits: [edit, edit] }, 'chart-edits-invalid'],
    ...['01', '1e3', 'NaN', 'Infinity', '-0.0', 9].map(replacement => [{ ...manifest, edits: [{ ...edit, replacement }] }, 'chart-edits-invalid']),
    [{ ...manifest, edits: [{ ...edit, replacement: edit.expected }] }, 'save-edit-no-changes'],
  ];
  for (const [bad, code] of cases) {
    await writeFile(path.join(f.directory, 'bad-chart.json'), JSON.stringify(bad)); const result = await runCliAsync(['save', 'source.pptx', 'bad-chart.pptx', '--mode', 'edit', '--chart-edits', 'bad-chart.json', '--json'], f.options);
    assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad-chart.pptx')), { code: 'ENOENT' });
  }
  for (const args of [['save', 'source.pptx', 'x', '--mode', 'no-op', '--chart-edits', 'bad-chart.json'], ['inspect', 'source.pptx', '--chart-values', '--json'], ['inspect', 'source.pptx', '--chart-edit-save-profile', 'existing-values-v1', '--json'], ['compare', 'source.pptx', 'source.pptx', '--chart-values', '--json']]) assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
  assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
});

test('chart candidate inspection counts mismatched cache/workbook values as unbound', async t => {
  const source = encodeChart(chartFixture({ xml: chartXml({ data: series({ refs: true, values: [0, -2.5] }), external: true }), book: workbook({ values: [8, -2.5] }) })), f = await fixture(t, source);
  const result = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--chart-edit-save-profile', 'existing-values-v1', '--chart-values', '--json'], f.options); assert.equal(result.exitCode, 0, result.output);
  const body = decode(result), unbound = body.report.chartValues.filter(item => item.status === 'unmapped'); assert.ok(unbound.length); assert.equal(body.chartEditCandidates.unboundCount, unbound.length);
  assert.ok(unbound.every(item => !body.chartEditCandidates.targets.some(candidate => candidate.slideId === item.slideId && candidate.elementId === item.elementId && candidate.seriesIndex === item.seriesIndex && candidate.pointIndex === item.pointIndex)));
});

test('save dry run reports complete ready and blocked plans without publishing', async t => {
  const source = encodeStructure(contentFixture()), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', editSaveProfile: 'text-runs-v1' }); const b = session.report.textRuns[0]; session.dispose();
  const manifest = { version: 'pptkit-text-edits-v1', sourceSha256: sha(source), edits: [{ slideId: b.slideId, elementId: b.elementId, paragraphIndex: b.paragraphIndex, runIndex: b.runIndex, expected: ' A & B ', replacement: 'Reviewed' }] };
  await writeFile(path.join(f.directory, 'edit.json'), JSON.stringify(manifest)); const args = ['save', 'source.pptx', 'output.pptx', '--mode', 'edit', '--text-edits', 'edit.json', '--json'];
  const planned = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(planned.exitCode, 0, planned.output); const plan = decode(planned); assert.equal(plan.status, 'ready'); assert.equal(plan.dryRun, true); assert.equal(plan.report.assurance, 'scoped-preservation'); assert.equal(plan.report.changes[0].text.after, 'Reviewed'); await assert.rejects(readFile(path.join(f.directory, 'output.pptx')), { code: 'ENOENT' });
  const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output); assert.deepEqual(decode(saved).report.changes, plan.report.changes);
  const existing = await readFile(path.join(f.directory, 'output.pptx')); const existingPlan = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(existingPlan.exitCode, 0); assert.deepEqual(await readFile(path.join(f.directory, 'output.pptx')), existing);
  const noOp = await runCliAsync(['save', 'source.pptx', 'unused.pptx', '--mode', 'no-op', '--dry-run', '--json'], f.options); assert.equal(noOp.exitCode, 0); assert.equal(decode(noOp).report.assurance, 'original-bytes'); await assert.rejects(readFile(path.join(f.directory, 'unused.pptx')), { code: 'ENOENT' });
  manifest.edits[0].replacement = manifest.edits[0].expected; await writeFile(path.join(f.directory, 'edit.json'), JSON.stringify(manifest));
  const blocked = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(blocked.exitCode, 1); assert.equal(decode(blocked).status, 'blocked'); assert.equal(decode(blocked).report.diagnostics[0].code, 'save-edit-no-changes'); assert.deepEqual(await readFile(path.join(f.directory, 'output.pptx')), existing);
  const human = await runCliAsync(['save', 'source.pptx', 'unused.pptx', '--mode', 'edit', '--text-edits', 'edit.json', '--dry-run'], f.options); assert.equal(human.exitCode, 1); assert.match(human.output, /save-edit-no-changes/);
  manifest.edits[0].replacement = 'x'.repeat(1000); await writeFile(path.join(f.directory, 'edit.json'), JSON.stringify(manifest));
  const limited = await runCliAsync(['save', 'source.pptx', 'unused.pptx', '--mode', 'edit', '--text-edits', 'edit.json', '--dry-run', '--json', '--max-output-bytes', '512'], f.options); assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded'); await assert.rejects(readFile(path.join(f.directory, 'unused.pptx')), { code: 'ENOENT' });
  const abort = new AbortController(); abort.abort(); assert.equal((await runCliAsync(['save', 'source.pptx', 'unused.pptx', '--mode', 'no-op', '--dry-run'], { ...f.options, signal: abort.signal })).exitCode, 130);
  for (const bad of [['inspect', 'source.pptx', '--dry-run'], ['generate', 'x', 'y', '--dry-run'], ['save', 'source.pptx', 'x', '--mode', 'no-op', '--dry-run', '--dry-run']]) assert.equal((await runCliAsync(bad, f.options)).exitCode, 2);
  assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
});

test('picture replacement manifest isolates shared media over three rounds', async t => {
  for (const strict of [false, true]) {
    const source = encodeImages(imageFixture(picture({ id: 2, crop: '<a:srcRect l="12500"/>' }) + picture({ id: 3 }), strict)), f = await fixture(t, source); let current = source, input = 'source.pptx';
    for (let round = 1; round <= 3; round++) {
      const inspection = await runCliAsync(['inspect', input, '--projection', 'content', '--picture-source-profile', 'embedded-v1', '--json'], f.options); assert.equal(inspection.exitCode, 0, inspection.output);
      const binding = decode(inspection).report.pictureSources[0], replacement = round % 2 ? validGif : png, mimeType = round % 2 ? 'image/gif' : 'image/png', width = round % 2 ? 1 : 8, height = round % 2 ? 1 : 4;
      const manifest = { version: 'pptkit-picture-edit-v1', sourceSha256: sha(current), replacementSha256: sha(replacement), slideId: binding.slideId, elementId: binding.elementId, expectedAssetId: binding.assetId, mimeType, width, height };
      await writeFile(path.join(f.directory, 'image.json'), JSON.stringify(manifest)); await writeFile(path.join(f.directory, 'replacement image'), replacement);
      const output = `picture-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--picture-edit', 'image.json', '--picture-bytes', 'replacement image', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); assert.equal(decode(dry).report.assurance, 'scoped-preservation'); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0, result.output); assert.deepEqual(decode(result).report.changes, decode(dry).report.changes); assert.equal(decode(result).inputs.length, 3);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(current), after = readGeneratedZip(next);
      for (const [name, bytes] of before) if (!['deck/z.xml', 'deck/_rels/z.xml.rels', '[Content_Types].xml'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
      const session = await importPptx(next, { projection: 'content', pictureEditSaveProfile: 'replace-v1' }); const [changed, sibling] = session.report.pictureSources;
      assert.notEqual(changed.mediaPart, sibling.mediaPart); assert.equal(sibling.mediaPart, '/media/raster.bin'); assert.deepEqual(after.get(changed.mediaPart.slice(1)), Buffer.from(replacement)); assert.equal(session.document.slides[0].elements[0].crop.left, 0.125); session.dispose(); current = next; input = output;
    }
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
  }
});

test('picture replacement rejects stale identities, bad bytes and incompatible edits', async t => {
  const source = encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', pictureEditSaveProfile: 'replace-v1' }), b = session.report.pictureSources[0]; session.dispose();
  const manifest = { version: 'pptkit-picture-edit-v1', sourceSha256: sha(source), replacementSha256: sha(validGif), slideId: b.slideId, elementId: b.elementId, expectedAssetId: b.assetId, mimeType: 'image/gif', width: 1, height: 1 };
  await writeFile(path.join(f.directory, 'replacement.gif'), validGif);
  const cases = [
    [{ ...manifest, sourceSha256: '0'.repeat(64) }, 'picture-edit-source-mismatch'],
    [{ ...manifest, replacementSha256: '0'.repeat(64) }, 'picture-edit-source-mismatch'],
    [{ ...manifest, expectedAssetId: 'missing' }, 'picture-edit-target-unbound'],
    [{ ...manifest, mimeType: ['image/gif'] }, 'picture-edit-invalid'],
    [{ ...manifest, width: 0 }, 'picture-edit-invalid'],
    [{ ...manifest, extra: true }, 'picture-edit-invalid'],
  ];
  for (const [bad, code] of cases) {
    await writeFile(path.join(f.directory, 'bad-picture.json'), JSON.stringify(bad)); const result = await runCliAsync(['save', 'source.pptx', 'bad-picture.pptx', '--mode', 'edit', '--picture-edit', 'bad-picture.json', '--picture-bytes', 'replacement.gif', '--json'], f.options);
    assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad-picture.pptx')), { code: 'ENOENT' });
  }
  const invalidBytes = Buffer.from('bad image'); await writeFile(path.join(f.directory, 'bad-image.png'), invalidBytes); await writeFile(path.join(f.directory, 'bad-picture.json'), JSON.stringify({ ...manifest, replacementSha256: sha(invalidBytes), mimeType: 'image/png' }));
  const invalid = await runCliAsync(['save', 'source.pptx', 'bad-picture.pptx', '--mode', 'edit', '--picture-edit', 'bad-picture.json', '--picture-bytes', 'bad-image.png', '--dry-run', '--json'], f.options); assert.equal(invalid.exitCode, 1, invalid.output); assert.equal(decode(invalid).status, 'blocked'); assert.ok(decode(invalid).report.diagnostics.some(d => d.code.startsWith('save-picture'))); await assert.rejects(readFile(path.join(f.directory, 'bad-picture.pptx')), { code: 'ENOENT' });
  for (const args of [['save', 'source.pptx', 'x', '--mode', 'edit', '--picture-edit', 'bad-picture.json'], ['save', 'source.pptx', 'x', '--mode', 'edit', '--picture-bytes', 'replacement.gif'], ['save', 'source.pptx', 'x', '--mode', 'edit', '--picture-edit', 'bad-picture.json', '--picture-bytes', 'replacement.gif', '--text-edits', 'anything'], ['inspect', 'source.pptx', '--picture-source-profile', 'embedded-v1']]) assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
  assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
});

test('grouped picture replacement combines atomically with first-slide-number', async t => {
  const source = encodeImages(imageFixture(group(picture({ id: 2 }) + picture({ id: 3 }), 10))), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', pictureEditSaveProfile: 'replace-v1', firstSlideNumberSaveProfile: 'attribute-v2' }); const b = session.report.pictureSources[0]; session.dispose();
  await writeFile(path.join(f.directory, 'replacement.gif'), validGif); await writeFile(path.join(f.directory, 'picture.json'), JSON.stringify({ version: 'pptkit-picture-edit-v1', sourceSha256: sha(source), replacementSha256: sha(validGif), slideId: b.slideId, elementId: b.elementId, expectedAssetId: b.assetId, mimeType: 'image/gif', width: 1, height: 1 }));
  const args = ['save', 'source.pptx', 'combined-picture.pptx', '--mode', 'edit', '--picture-edit', 'picture.json', '--picture-bytes', 'replacement.gif', '--first-slide-number', '17', '--json'];
  const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); assert.equal(decode(dry).report.changes.length, 3);
  const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0, result.output); assert.deepEqual(decode(result).report.changes, decode(dry).report.changes);
  const bytes = await readFile(path.join(f.directory, 'combined-picture.pptx')), after = await importPptx(bytes, { projection: 'content', pictureEditSaveProfile: 'replace-v1', firstSlideNumberSaveProfile: 'attribute-v2' });
  assert.equal(after.document.firstSlideNumber, 17); assert.equal(after.document.slides[0].elements[0].type, 'group'); assert.notEqual(after.report.pictureSources[0].mediaPart, after.report.pictureSources[1].mediaPart); assert.equal(after.report.pictureSources[1].mediaPart, '/media/raster.bin'); after.dispose();
});

test('picture manifest preserves Parser rejection of opaque blip effects', async t => {
  const source = encodeImages(imageFixture(picture({ effects: '<a:extLst/>' }))), f = await fixture(t, source), session = await importPptx(source, { projection: 'content', pictureEditSaveProfile: 'replace-v1' }); const b = session.report.pictureSources[0]; session.dispose();
  await writeFile(path.join(f.directory, 'replacement.gif'), validGif); await writeFile(path.join(f.directory, 'picture.json'), JSON.stringify({ version: 'pptkit-picture-edit-v1', sourceSha256: sha(source), replacementSha256: sha(validGif), slideId: b.slideId, elementId: b.elementId, expectedAssetId: b.assetId, mimeType: 'image/gif', width: 1, height: 1 }));
  const result = await runCliAsync(['save', 'source.pptx', 'blocked-picture.pptx', '--mode', 'edit', '--picture-edit', 'picture.json', '--picture-bytes', 'replacement.gif', '--dry-run', '--json'], f.options); assert.equal(result.exitCode, 1); assert.equal(decode(result).status, 'blocked'); assert.ok(decode(result).report.diagnostics.some(d => d.code === 'save-picture-candidate-unsupported')); await assert.rejects(readFile(path.join(f.directory, 'blocked-picture.pptx')), { code: 'ENOENT' });
});

test('generate creates an editable deck from bounded JSON with embedded assets', async t => {
  const f = await fixture(t);
  const image = (await import('../../pptx-parser/test/image-fixture.mjs')).png;
  const spec = { version: 'pptkit-authoring-v1', presentation: { id: 'doc', metadata: { title: 'CLI deck' }, size: { width: 500, height: 250 }, firstSlideNumber: 17 }, assets: [{ id: 'img', kind: 'image', source: { type: 'bytes', encoding: 'base64', value: Buffer.from(image).toString('base64') }, mimeType: 'image/png', width: 8, height: 4 }], slides: [{ id: 'page', elements: [{ id: 'title', type: 'text', box: { x: 10, y: 10, width: 200, height: 40 }, content: 'Editable title' }, { id: 'photo', type: 'image', assetId: 'img', box: { x: 220, y: 10, width: 100, height: 50 } }] }] };
  await writeFile(path.join(f.directory, 'deck.json'), JSON.stringify(spec));
  const result = await runCliAsync(['generate', 'deck.json', 'generated.pptx', '--json'], f.options); assert.equal(result.exitCode, 0, result.output);
  const body = decode(result); assert.equal(body.command, 'generate'); assert.equal(body.report.status, 'committed'); assert.equal(body.report.slideCount, 1); assert.deepEqual(body.report.exportWarnings, []);
  const bytes = await readFile(path.join(f.directory, 'generated.pptx')), session = await importPptx(bytes, { projection: 'content' });
  try { assert.equal(session.document.metadata.title, 'CLI deck'); assert.equal(session.document.firstSlideNumber, 17); assert.equal(session.document.slides[0].elements[0].type, 'text'); assert.equal(session.document.slides[0].elements[1].type, 'image'); } finally { session.dispose(); }
  const repeat = await runCliAsync(['generate', 'deck.json', 'generated.pptx', '--json'], f.options); assert.equal(repeat.exitCode, 3); assert.equal(decode(repeat).error.code, 'file-target-exists');
  assert.deepEqual(await readFile(path.join(f.directory, 'generated.pptx')), bytes);
  spec.assets[0].source.value = Buffer.from('not an image').toString('base64');
  await writeFile(path.join(f.directory, 'warning.json'), JSON.stringify(spec));
  const warning = await runCliAsync(['generate', 'warning.json', 'warning.pptx', '--json'], f.options);
  assert.equal(warning.exitCode, 0, warning.output); assert.ok(decode(warning).report.exportWarnings.length > 0);
});

test('opt-in image diagnostics expose incomplete PNG, GIF and JPEG before publication', async t => {
  const f = await fixture(t);
  const inputs = [['png', Buffer.from([137,80,78,71,13,10,26,10])], ['gif', Buffer.from('GIF89a')], ['jpeg', Buffer.from([255,216,255])]];
  const spec = {version: 'pptkit-authoring-v1', presentation: {size: {width: 200, height: 100}}, assets: inputs.map(([id, bytes]) => ({id, kind: 'image', source: {type: 'bytes', encoding: 'base64', value: bytes.toString('base64')}})), slides: [{elements: inputs.map(([id], i) => ({type: 'image', assetId: id, box: {x: i * 40, y: 10, width: 30, height: 30}}))}]};
  await writeFile(path.join(f.directory, 'images.json'), JSON.stringify(spec));
  const baseline = await runCliAsync(['validate', 'images.json', '--export', '--json'], f.options); assert.equal(baseline.exitCode, 0, baseline.output);
  assert.deepEqual(decode(baseline).report.export.warnings, []);
  const checked = await runCliAsync(['validate', 'images.json', '--export', '--image-diagnostics', '--json'], f.options); assert.equal(checked.exitCode, 0, checked.output);
  const warnings = decode(checked).report.export.warnings;
  for (const code of ['png-inspection-incomplete', 'gif-inspection-incomplete', 'jpeg-exif-inspection-incomplete', 'jpeg-icc-inspection-incomplete']) assert.ok(warnings.some(item => item.code === code), code);
  const strict = await runCliAsync(['validate', 'images.json', '--export', '--image-diagnostics', '--fail-on-export-warning', '--json'], f.options); assert.equal(strict.exitCode, 1); assert.equal(decode(strict).report.valid, false);
  const rejected = await runCliAsync(['generate', 'images.json', 'rejected-images.pptx', '--image-diagnostics', '--fail-on-export-warning', '--json'], f.options); assert.equal(rejected.exitCode, 3); assert.equal(decode(rejected).error.code, 'export-warnings-rejected'); await assert.rejects(readFile(path.join(f.directory, 'rejected-images.pptx')), {code: 'ENOENT'});
  const generated = await runCliAsync(['generate', 'images.json', 'images.pptx', '--image-diagnostics', '--json'], f.options); assert.equal(generated.exitCode, 0, generated.output); assert.deepEqual(decode(generated).report.exportWarnings, warnings);
  for (const args of [['validate', 'images.json', '--image-diagnostics'], ['preview', 'images.json', '--slide', '1', '--json', '--image-diagnostics'], ['inspect', 'x', '--image-diagnostics'], ['generate', 'images.json', 'x', '--image-diagnostics', '--image-diagnostics']]) assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
});

test('image diagnostics report real animation and metadata once per asset without changing media bytes', async t => {
  const f = await fixture(t);
  const fixtures = [
    ['animated-png', await readFile(new URL('../../../tests/fixtures/two-frame-apng.png', import.meta.url))],
    ['png-icc', await readFile(new URL('../../../tests/fixtures/png-icc-srgb.png', import.meta.url))],
    ['jpeg-icc', await readFile(new URL('../../../tests/fixtures/jpeg-icc-srgb.jpg', import.meta.url))],
    ['jpeg-exif', await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url))],
    ['animated-gif', Buffer.from('R0lGODlhAgACAIEAAP8AAAAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwECAAAh+QQIAAAAACwAAAAAAgACAAAIBgABCAQQEAAh+QQIBwAAACwAAAAAAgACAIEAAP8AAAAAAAAAAAAIBgABCAQQEAA7', 'base64')],
  ];
  const assets = fixtures.map(([id, bytes]) => ({id, kind: 'image', source: {type: 'bytes', encoding: 'base64', value: bytes.toString('base64')}}));
  const elements = [...fixtures.map(([id], i) => ({type: 'image', assetId: id, box: {x: i * 35, y: 10, width: 30, height: 30}})), {type: 'image', assetId: 'animated-gif', box: {x: 10, y: 50, width: 30, height: 30}}];
  await writeFile(path.join(f.directory, 'real-images.json'), JSON.stringify({version: 'pptkit-authoring-v1', presentation: {size: {width: 200, height: 100}}, assets, slides: [{elements}]}));
  const args = ['--image-diagnostics', '--json'];
  const checked = await runCliAsync(['validate', 'real-images.json', '--export', ...args], f.options); assert.equal(checked.exitCode, 0, checked.output);
  const warnings = decode(checked).report.export.warnings;
  for (const [assetId, code] of [['animated-png','apng-animation-playback-unverified'], ['png-icc','png-icc-color-unverified'], ['jpeg-icc','jpeg-icc-color-unverified'], ['jpeg-exif','jpeg-exif-orientation-unverified'], ['animated-gif','gif-animation-playback-unverified']]) {
    assert.equal(warnings.filter(item => item.code === code && item.assetId === assetId).length, 1, `${assetId}: ${code}`);
  }
  const generated = await runCliAsync(['generate', 'real-images.json', 'real-images.pptx', ...args], f.options); assert.equal(generated.exitCode, 0, generated.output); assert.deepEqual(decode(generated).report.exportWarnings, warnings);
  const zip = readGeneratedZip(await readFile(path.join(f.directory, 'real-images.pptx')));
  const hashes = [...zip.entries()].filter(([name]) => name.startsWith('ppt/media/')).map(([, bytes]) => sha(bytes)).sort();
  assert.deepEqual(hashes, fixtures.map(([, bytes]) => sha(bytes)).sort());
  const strict = await runCliAsync(['generate', 'real-images.json', 'strict-images.pptx', '--image-diagnostics', '--fail-on-export-warning', '--json'], f.options); assert.equal(strict.exitCode, 3); await assert.rejects(readFile(path.join(f.directory, 'strict-images.pptx')), {code: 'ENOENT'});
  const cleanAssets = [['clean-png', png], ['clean-gif', validGif]].map(([id, bytes]) => ({id, kind: 'image', source: {type: 'bytes', encoding: 'base64', value: bytes.toString('base64')}}));
  await writeFile(path.join(f.directory, 'clean-images.json'), JSON.stringify({version: 'pptkit-authoring-v1', assets: cleanAssets, slides: [{elements: cleanAssets.map((asset, i) => ({type: 'image', assetId: asset.id, box: {x: i * 40, y: 10, width: 30, height: 30}}))}]}));
  const clean = await runCliAsync(['generate', 'clean-images.json', 'clean-images.pptx', '--image-diagnostics', '--fail-on-export-warning', '--json'], f.options); assert.equal(clean.exitCode, 0, clean.output); assert.deepEqual(decode(clean).report.exportWarnings, []);
});

test('CLI reports malformed PNG ICC envelope before strict publication', async t => {
  const f = await fixture(t), original = await readFile(new URL('../../../tests/fixtures/png-icc-srgb.png', import.meta.url));
  let offset = 8, start = -1, end = -1;
  while (offset < original.length) { const length = original.readUInt32BE(offset), name = original.toString('ascii', offset + 4, offset + 8), next = offset + length + 12; if (name === 'iCCP') { start = offset; end = next; break; } offset = next; }
  assert.ok(start > 0); const bytes = Buffer.concat([original.subarray(0, end), original.subarray(start, end), original.subarray(end)]);
  const spec = {version: 'pptkit-authoring-v1', assets: [{id: 'icc', kind: 'image', source: {type: 'bytes', encoding: 'base64', value: bytes.toString('base64')}}], slides: [{elements: [{type: 'image', assetId: 'icc', box: {x: 10, y: 10, width: 30, height: 30}}]}]};
  await writeFile(path.join(f.directory, 'bad-icc.json'), JSON.stringify(spec));
  const result = await runCliAsync(['validate', 'bad-icc.json', '--export', '--image-diagnostics', '--json'], f.options); assert.equal(result.exitCode, 0, result.output);
  assert.deepEqual(decode(result).report.export.warnings.map(item => item.code), ['png-icc-color-unverified', 'png-icc-inspection-incomplete']);
  const rejected = await runCliAsync(['generate', 'bad-icc.json', 'bad-icc.pptx', '--image-diagnostics', '--fail-on-export-warning', '--json'], f.options); assert.equal(rejected.exitCode, 3); await assert.rejects(readFile(path.join(f.directory, 'bad-icc.pptx')), {code: 'ENOENT'});
});

test('JSON round connector paths validate, preview, generate and reimport without losing closure', async t => {
  const f = await fixture(t);
  const closed = {type: 'connector', id: 'closed', start: {x: 20, y: 20}, route: [{x: 80, y: 20}, {x: 80, y: 60}], end: {x: 20, y: 20}, pathMode: 'closed-round', style: {paint: {type: 'solid', color: 'FF0000', opacity: .5}, width: 3}};
  const open = {type: 'connector', id: 'open', start: {x: 100, y: 20}, end: {x: 150, y: 50}, pathMode: 'open-round', style: {paint: {type: 'solid', color: '0000FF'}, width: 2}};
  const spec = {version: 'pptkit-authoring-v1', presentation: {size: {width: 200, height: 100}}, slides: [{elements: [closed, open]}]};
  await writeFile(path.join(f.directory, 'paths.json'), JSON.stringify(spec));
  const valid = await runCliAsync(['validate', 'paths.json', '--export', '--json'], f.options); assert.equal(valid.exitCode, 0, valid.output);
  const preview = await runCliAsync(['preview', 'paths.json', '--slide', '1', '--json'], f.options); assert.equal(preview.exitCode, 0, preview.output);
  assert.match(decode(preview).report.slide.svg, /<polygon /); assert.match(decode(preview).report.slide.svg, /<polyline /);
  const generated = await runCliAsync(['generate', 'paths.json', 'paths.pptx', '--json'], f.options); assert.equal(generated.exitCode, 0, generated.output);
  const bytes = await readFile(path.join(f.directory, 'paths.pptx')), xml = readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString();
  assert.match(xml, /<a:close\/>/); assert.equal((xml.match(/cap="rnd"/g) ?? []).length, 2); assert.match(xml, /<a:alpha val="50000"\/>/);
  const imported = await importPptx(bytes, {projection: 'content'});
  try { const elements = imported.document.slides[0].elements; assert.deepEqual(elements.map(item => item.pathMode), ['closed-round', 'open-round']); assert.deepEqual(elements[0].start, elements[0].end); assert.equal(elements[0].style.paint.opacity, .5); }
  finally { imported.dispose(); }
  spec.slides[0].elements[0].end = {x: 21, y: 20}; await writeFile(path.join(f.directory, 'invalid-path.json'), JSON.stringify(spec));
  const invalid = await runCliAsync(['validate', 'invalid-path.json', '--json'], f.options); assert.equal(invalid.exitCode, 1); assert.ok(decode(invalid).report.diagnostics.some(item => item.code === 'invalid-connector-path-mode'));
  const rejected = await runCliAsync(['generate', 'invalid-path.json', 'invalid-path.pptx', '--json'], f.options); assert.equal(rejected.exitCode, 3); await assert.rejects(readFile(path.join(f.directory, 'invalid-path.pptx')), {code: 'ENOENT'});
});

test('generate rejects invalid authoring, external assets, existing files and bounded output before commit', async t => {
  const f = await fixture(t); const target = path.join(f.directory, 'out.pptx');
  const cases = [
    ['bad JSON', '{'],
    ['unknown schema', JSON.stringify({ version: 'v2', slides: [] })],
    ['unknown field', JSON.stringify({ version: 'pptkit-authoring-v1', slides: [], forgotten: true })],
    ['unknown slide field', JSON.stringify({ version: 'pptkit-authoring-v1', slides: [{ typo: 1 }] })],
    ['external path', JSON.stringify({ version: 'pptkit-authoring-v1', slides: [], assets: [{ kind: 'image', source: { type: 'path', value: '/tmp/img.png' } }] })],
    ['invalid geometry', JSON.stringify({ version: 'pptkit-authoring-v1', slides: [{ elements: [{ type: 'text', box: { x: 0, y: 0, width: -1, height: 10 }, content: 'bad' }] }] })],
  ];
  for (const [name, input] of cases) {
    await writeFile(path.join(f.directory, 'bad.json'), input); const result = await runCliAsync(['generate', 'bad.json', 'out.pptx', '--json'], f.options);
    assert.equal(result.exitCode, 3, `${name}: ${result.output}`); await assert.rejects(readFile(target), { code: 'ENOENT' });
  }
  await writeFile(path.join(f.directory, 'good.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [{}] }));
  const name = 'z'.repeat(230) + '.pptx'; const limited = await runCliAsync(['generate', 'good.json', name, '--max-output-bytes', '512', '--json'], f.options);
  assert.equal(limited.exitCode, 4); await assert.rejects(readFile(path.join(f.directory, name)), { code: 'ENOENT' });
  const controller = new AbortController(); controller.abort(); const aborted = await runCliAsync(['generate', 'good.json', 'out.pptx', '--json'], { ...f.options, signal: controller.signal });
  assert.equal(aborted.exitCode, 130); await assert.rejects(readFile(target), { code: 'ENOENT' });
});

test('validate reports Core diagnostics without generating or writing a deck', async t => {
  const f = await fixture(t); const filename = path.join(f.directory, 'validate.json');
  const spec = { version: 'pptkit-authoring-v1', slides: [{ id: 's', elements: [{ id: 'bad', type: 'text', box: { x: 0, y: 0, width: -1, height: 10 }, content: 'bad' }] }] };
  await writeFile(filename, JSON.stringify(spec));
  const invalid = await runCliAsync(['validate', 'validate.json', '--json'], f.options), body = decode(invalid);
  assert.equal(invalid.exitCode, 1); assert.notEqual(invalid.channel, 'stderr'); assert.equal(body.status, 'invalid'); assert.equal(body.report.valid, false); assert.equal(body.report.scope, 'core-authoring');
  assert.ok(body.report.diagnostics.some(d => d.severity === 'error' && d.path.includes('box')));
  const generated = await runCliAsync(['generate', 'validate.json', 'out.pptx', '--json'], f.options);
  assert.equal(generated.exitCode, 3); assert.deepEqual(decode(generated).error.source.diagnostics, body.report.diagnostics);
  await assert.rejects(readFile(path.join(f.directory, 'out.pptx')), { code: 'ENOENT' });
  spec.slides[0].elements[0].box.width = 100; await writeFile(filename, JSON.stringify(spec));
  const valid = await runCliAsync(['validate', 'validate.json', '--json'], f.options); assert.equal(valid.exitCode, 0); assert.equal(decode(valid).report.valid, true);
  const human = await runCliAsync(['validate', 'validate.json'], f.options); assert.equal(human.exitCode, 0); assert.match(human.output, /Core authoring validation only/);
  await assert.rejects(readFile(path.join(f.directory, 'out.pptx')), { code: 'ENOENT' });
});

test('validate rejects malformed or unsupported input, cross-command flags, bounds and cancellation', async t => {
  const f = await fixture(t); await writeFile(path.join(f.directory, 'bad.json'), '{');
  const invalid = await runCliAsync(['validate', 'bad.json', '--json'], f.options); assert.equal(invalid.exitCode, 3); assert.equal(decode(invalid).error.code, 'authoring-invalid-json');
  await writeFile(path.join(f.directory, 'remote.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [], assets: [{ kind: 'image', source: { type: 'url', value: 'https://example.invalid/a.png' } }] }));
  const remote = await runCliAsync(['validate', 'remote.json', '--json'], f.options); assert.equal(remote.exitCode, 3); assert.equal(decode(remote).error.code, 'authoring-asset-unsupported');
  await writeFile(path.join(f.directory, 'null-box.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [{ elements: [{ type: 'text', box: null, content: 'x' }] }] }));
  for (const args of [['validate', 'null-box.json'], ['generate', 'null-box.json', 'null-box.pptx']]) {
    const malformed = await runCliAsync([...args, '--json'], f.options); assert.equal(malformed.exitCode, 3); assert.equal(decode(malformed).error.code, 'authoring-invalid');
  }
  for (const element of [{ type: 'connector', start: null, end: null }, { type: 'table', rows: null, columns: null }]) {
    await writeFile(path.join(f.directory, 'invalid-element.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [{ elements: [element] }] }));
    const malformed = await runCliAsync(['validate', 'invalid-element.json', '--json'], f.options); assert.equal(malformed.exitCode, 3); assert.equal(decode(malformed).error.code, 'authoring-invalid');
  }
  const tooMany = { version: 'pptkit-authoring-v1', slides: [{ elements: Array.from({ length: 100001 }, () => ({ type: 'text' })) }] };
  await writeFile(path.join(f.directory, 'many.json'), JSON.stringify(tooMany));
  const excessive = await runCliAsync(['validate', 'many.json', '--json'], f.options); assert.equal(excessive.exitCode, 4); assert.equal(decode(excessive).error.code, 'authoring-limit-exceeded');
  assert.equal((await runCliAsync(['validate', 'bad.json', '--mode', 'no-op'], f.options)).exitCode, 2);
  const limited = await runCliAsync(['validate', 'remote.json', '--max-input-bytes', '1', '--json'], f.options); assert.equal(limited.exitCode, 4);
  const abort = new AbortController(); abort.abort(); const cancelled = await runCliAsync(['validate', 'remote.json', '--json'], { ...f.options, signal: abort.signal }); assert.equal(cancelled.exitCode, 130);
});

test('validate --export checks generated package and reports exporter warnings without writing', async t => {
  const f = await fixture(t), name = 'export check.json';
  const spec = { version: 'pptkit-authoring-v1', assets: [{ id: 'img', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', encoding: 'base64', value: Buffer.from('not an image').toString('base64') }, width: 8, height: 4 }], slides: [{ elements: [{ type: 'image', assetId: 'img', box: { x: 0, y: 0, width: 80, height: 40 } }] }] };
  await writeFile(path.join(f.directory, name), JSON.stringify(spec));
  const before = (await readdir(f.directory)).sort();
  const plain = await runCliAsync(['validate', name, '--json'], f.options); assert.equal(plain.exitCode, 0); assert.equal(decode(plain).report.export, undefined);
  const checked = await runCliAsync(['validate', name, '--export', '--json'], f.options); assert.equal(checked.exitCode, 0, checked.output);
  const report = decode(checked).report; assert.equal(report.scope, 'core-and-export'); assert.equal(report.export.status, 'checked'); assert.equal(report.export.slideCount, 1); assert.ok(report.export.byteLength > 0); assert.ok(report.export.warnings.length > 0);
  assert.equal(decode(checked).nativeVerification, 'not-performed');
  const strict = await runCliAsync(['validate', name, '--export', '--fail-on-export-warning', '--json'], f.options); assert.equal(strict.exitCode, 1); assert.equal(decode(strict).report.valid, false); assert.equal(decode(strict).report.failOnExportWarning, true); assert.deepEqual(decode(strict).report.export.warnings, report.export.warnings);
  const rejectedOutput = path.join(f.directory, 'rejected.pptx');
  const rejectedWarnings = await runCliAsync(['generate', name, 'rejected.pptx', '--fail-on-export-warning', '--json'], f.options); assert.equal(rejectedWarnings.exitCode, 3); assert.equal(decode(rejectedWarnings).error.code, 'export-warnings-rejected'); assert.deepEqual(decode(rejectedWarnings).error.source.warnings, report.export.warnings); await assert.rejects(readFile(rejectedOutput), { code: 'ENOENT' });
  const manyWarnings = { ...spec, slides: [{ elements: Array.from({ length: 10 }, (_, i) => ({ ...spec.slides[0].elements[0], id: `img-${i}` })) }] };
  await writeFile(path.join(f.directory, 'many-warnings.json'), JSON.stringify(manyWarnings));
  const limitedWarning = await runCliAsync(['generate', 'many-warnings.json', 'limited.pptx', '--fail-on-export-warning', '--json', '--max-output-bytes', '512'], f.options); assert.equal(limitedWarning.exitCode, 4); assert.equal(decode(limitedWarning).error.code, 'output-limit-exceeded'); await assert.rejects(readFile(path.join(f.directory, 'limited.pptx')), { code: 'ENOENT' });
  await writeFile(rejectedOutput, 'existing'); const existing = await runCliAsync(['generate', name, 'rejected.pptx', '--fail-on-export-warning', '--json'], f.options); assert.equal(existing.exitCode, 3); assert.equal(await readFile(rejectedOutput, 'utf8'), 'existing');
  const limited = await runCliAsync(['validate', name, '--export', '--json', '--max-output-bytes', '512'], f.options); assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
  assert.deepEqual((await readdir(f.directory)).sort(), [...before, 'rejected.pptx', 'many-warnings.json'].sort()); assert.equal(await readFile(path.join(f.directory, name), 'utf8'), JSON.stringify(spec));
  const generated = await runCliAsync(['generate', name, 'output.pptx', '--json'], f.options); assert.equal(generated.exitCode, 0, generated.output); assert.deepEqual(report.export.warnings, decode(generated).report.exportWarnings);
  const timed = { version: 'pptkit-authoring-v1', slides: [{ id: 'p', elements: [{ id: 'a', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 100, height: 50 } }], eventTimeline: { stages: [{ id: 'stage', target: { pageId: 'p', elementId: 'a' }, trigger: { type: 'object-click', elementId: 'a' }, timing: { kind: 'span', id: 'time', durationMs: 10, fill: 'hold' } }] } }] };
  await writeFile(path.join(f.directory, 'clean.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [{}] }));
  const clean = await runCliAsync(['validate', 'clean.json', '--export', '--fail-on-export-warning', '--json'], f.options); assert.equal(clean.exitCode, 0); assert.deepEqual(decode(clean).report.export.warnings, []);
  assert.equal((await runCliAsync(['generate', 'clean.json', 'clean.pptx', '--fail-on-export-warning', '--json'], f.options)).exitCode, 0);
  await writeFile(path.join(f.directory, 'timed.json'), JSON.stringify(timed));
  assert.equal((await runCliAsync(['validate', 'timed.json', '--json'], f.options)).exitCode, 0);
  const rejected = await runCliAsync(['validate', 'timed.json', '--export', '--json'], f.options); assert.equal(rejected.exitCode, 3); assert.ok(decode(rejected).error.source.diagnostics.some(d => d.code === 'event-timeline-export-unsupported'));
  spec.slides[0].elements[0].box.width = -1; await writeFile(path.join(f.directory, name), JSON.stringify(spec));
  const invalid = await runCliAsync(['validate', name, '--export', '--json'], f.options); assert.equal(invalid.exitCode, 1); assert.equal(decode(invalid).report.export.status, 'skipped-invalid-core');
  for (const args of [['inspect', 'x', '--export'], ['generate', name, 'x', '--export'], ['preview', name, '--slide', '1', '--json', '--export'], ['validate', name, '--export', '--export'], ['validate', name, '--fail-on-export-warning'], ['inspect', 'x', '--fail-on-export-warning']]) assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
  const controller = new AbortController(); controller.abort(); assert.equal((await runCliAsync(['validate', name, '--export'], { ...f.options, signal: controller.signal })).exitCode, 130);
});

test('validate can report missing descriptions and optionally reject them', async t => {
  const f = await fixture(t);
  const table = { id: 'table', type: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ cells: [{ content: 'A' }] }] };
  const chart = { id: 'chart', type: 'chart', box: { x: 120, y: 0, width: 100, height: 50 }, chartType: 'bar', categories: ['A'], series: [{ name: 'S', values: [1] }], accessibility: { description: 'Orders' } };
  const spec = { version: 'pptkit-authoring-v1', slides: [{ id: 'page', elements: [table, chart] }] };
  await writeFile(path.join(f.directory, 'descriptions.json'), JSON.stringify(spec));
  const base = await runCliAsync(['validate', 'descriptions.json', '--json'], f.options);
  assert.equal(decode(base).report.descriptions, undefined);
  const inspected = await runCliAsync(['validate', 'descriptions.json', '--check-descriptions', '--export', '--json'], f.options);
  assert.equal(inspected.exitCode, 0, inspected.output);
  const report = decode(inspected).report;
  assert.equal(report.descriptions.status, 'checked'); assert.equal(report.descriptions.scope, 'authored-definitions');
  assert.deepEqual(report.descriptions.findings.map(finding => [finding.elementId, finding.slideId]), [['table', 'page']]);
  assert.equal(report.export.status, 'checked'); assert.equal(report.valid, true);
  const strict = await runCliAsync(['validate', 'descriptions.json', '--check-descriptions', '--fail-on-missing-description', '--json'], f.options);
  assert.equal(strict.exitCode, 1); assert.equal(decode(strict).report.valid, false); assert.equal(decode(strict).report.failOnMissingDescription, true);
  const human = await runCliAsync(['validate', 'descriptions.json', '--check-descriptions'], f.options);
  assert.equal(human.exitCode, 0); assert.match(human.output, /missing-description/);
  const many = { ...spec, slides: [{ id: 'page', elements: Array.from({ length: 12 }, (_, index) => ({ ...table, id: `table-${index}` })) }] };
  await writeFile(path.join(f.directory, 'many-descriptions.json'), JSON.stringify(many));
  const limited = await runCliAsync(['validate', 'many-descriptions.json', '--check-descriptions', '--max-output-bytes', '512', '--json'], f.options);
  assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
  spec.slides[0].elements[0].accessibility = { decorative: true }; await writeFile(path.join(f.directory, 'descriptions.json'), JSON.stringify(spec));
  const clean = await runCliAsync(['validate', 'descriptions.json', '--check-descriptions', '--fail-on-missing-description', '--json'], f.options);
  assert.equal(clean.exitCode, 0); assert.deepEqual(decode(clean).report.descriptions.findings, []);
  spec.slides[0].elements[0].box.width = -1; await writeFile(path.join(f.directory, 'descriptions.json'), JSON.stringify(spec));
  const invalid = await runCliAsync(['validate', 'descriptions.json', '--check-descriptions', '--export', '--json'], f.options);
  assert.equal(invalid.exitCode, 1); assert.equal(decode(invalid).report.descriptions.status, 'skipped-invalid-core'); assert.equal(decode(invalid).report.export.status, 'skipped-invalid-core');
  for (const args of [['validate', 'missing', '--fail-on-missing-description'], ['inspect', 'missing', '--check-descriptions'], ['generate', 'missing', 'out', '--check-descriptions'], ['validate', 'missing', '--check-descriptions', '--check-descriptions']]) assert.equal((await runCliAsync(args)).exitCode, 2);
  const controller = new AbortController(); controller.abort();
  assert.equal((await runCliAsync(['validate', 'descriptions.json', '--check-descriptions'], { ...f.options, signal: controller.signal })).exitCode, 130);
});

test('validate description occurrences report only modeled visible pages and bound expansion', async t => {
  const f = await fixture(t);
  const table = id => ({ id, type: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ cells: [{ content: 'A' }] }] });
  const spec = { version: 'pptkit-authoring-v1', presentation: { masters: [{ id: 'm', name: 'Master', elements: [table('master-table')] }] }, layouts: [{ id: 'l', name: 'Layout', masterId: 'm', elements: [table('layout-table')] }, { id: 'unused', name: 'Unused', elements: [table('unused-table')] }], slides: [{ id: 'visible', layoutId: 'l' }, { id: 'hidden-inherited', layoutId: 'l', showInheritedShapes: false }, { id: 'hidden-slide', layoutId: 'l', hidden: true }] };
  const file = path.join(f.directory, 'occurrences.json'); await writeFile(file, JSON.stringify(spec));
  const args = ['validate', 'occurrences.json', '--check-descriptions', '--description-occurrences', '--json'];
  const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0, result.output);
  const report = decode(result).report.descriptions; assert.equal(report.scope, 'modeled-slide-occurrences');
  assert.deepEqual(report.findings.map(finding => [finding.slideId, finding.elementId, finding.source.kind]), [['visible', 'master-table', 'master'], ['visible', 'layout-table', 'layout']]);
  const human = await runCliAsync(args.filter(arg => arg !== '--json'), f.options); assert.match(human.output, /slide=visible source=master:m/);
  const strict = await runCliAsync([...args, '--fail-on-missing-description'], f.options); assert.equal(strict.exitCode, 1);
  spec.slides[0].showInheritedShapes = false; await writeFile(file, JSON.stringify(spec));
  const clean = await runCliAsync([...args, '--fail-on-missing-description'], f.options); assert.equal(clean.exitCode, 0); assert.deepEqual(decode(clean).report.descriptions.findings, []);
  const definitions = await runCliAsync(['validate', 'occurrences.json', '--check-descriptions', '--fail-on-missing-description', '--json'], f.options); assert.equal(definitions.exitCode, 1);
  spec.layouts[0].elements[0].box.width = -1; await writeFile(file, JSON.stringify(spec));
  const invalidCore = await runCliAsync(args, f.options); assert.equal(invalidCore.exitCode, 1); assert.equal(decode(invalidCore).report.descriptions.status, 'skipped-invalid-core');
  spec.layouts[0].elements[0].box.width = 100;
  spec.slides = Array.from({ length: 30 }, (_, index) => ({ id: `page-${index}`, layoutId: 'l' })); await writeFile(file, JSON.stringify(spec));
  const limited = await runCliAsync([...args, '--max-output-bytes', '512'], f.options); assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
  for (const invalid of [['validate', 'missing', '--description-occurrences'], ['inspect', 'missing', '--check-descriptions', '--description-occurrences'], ['validate', 'missing', '--check-descriptions', '--description-occurrences', '--description-occurrences']]) assert.equal((await runCliAsync(invalid, f.options)).exitCode, 2);
});

test('validate contrast distinguishes numeric failures and unknown scenes in its strict policy', async t => {
  const f = await fixture(t);
  const makeText = (color = '777777') => ({ id: 't', type: 'text', box: { x: 0, y: 0, width: 200, height: 50 }, content: [{ runs: [{ text: 'Text', style: { color, fontSize: 12 } }] }] });
  const spec = { version: 'pptkit-authoring-v1', slides: [{ id: 'page', elements: [makeText()] }] };
  const file = path.join(f.directory, 'contrast.json'); await writeFile(file, JSON.stringify(spec));
  const check = await runCliAsync(['validate', 'contrast.json', '--check-contrast', '--check-descriptions', '--export', '--json'], f.options);
  assert.equal(check.exitCode, 0, check.output); const report = decode(check).report;
  assert.equal(report.contrast.status, 'checked'); assert.equal(report.contrast.findings[0].status, 'below-threshold'); assert.equal(report.descriptions.status, 'checked'); assert.equal(report.export.status, 'checked');
  const strict = await runCliAsync(['validate', 'contrast.json', '--check-contrast', '--fail-on-contrast', '--json'], f.options);
  assert.equal(strict.exitCode, 1); assert.equal(decode(strict).report.failOnContrast, true);
  spec.slides[0].elements = [makeText('000000')]; await writeFile(file, JSON.stringify(spec));
  const pass = await runCliAsync(['validate', 'contrast.json', '--check-contrast', '--fail-on-contrast', '--json'], f.options);
  assert.equal(pass.exitCode, 0); assert.equal(decode(pass).report.contrast.findings[0].ratio, 21);
  const sizeDependent = makeText(); sizeDependent.content[0].runs[0].style.fontSize = 18;
  spec.slides[0].elements = [sizeDependent]; await writeFile(file, JSON.stringify(spec));
  const sizeStrict = await runCliAsync(['validate', 'contrast.json', '--check-contrast', '--fail-on-contrast', '--json'], f.options);
  assert.equal(sizeStrict.exitCode, 1); assert.equal(decode(sizeStrict).report.contrast.findings[0].reason, 'output-size-dependent');
  const alphaDependent = makeText('000000'); alphaDependent.content[0].runs[0].style.colorOpacity = 0.534681;
  spec.slides[0].elements = [alphaDependent]; await writeFile(file, JSON.stringify(spec));
  const alphaStrict = await runCliAsync(['validate', 'contrast.json', '--check-contrast', '--fail-on-contrast', '--json'], f.options);
  assert.equal(alphaStrict.exitCode, 1); assert.equal(decode(alphaStrict).report.contrast.findings[0].reason, 'output-opacity-dependent');
  spec.slides[0].elements = [makeText('000000')];
  spec.slides[0].elements.push({ type: 'shape', id: 'overlay', shape: 'rect', box: { x: 300, y: 0, width: 50, height: 50 } }); await writeFile(file, JSON.stringify(spec));
  const unknown = await runCliAsync(['validate', 'contrast.json', '--check-contrast', '--fail-on-contrast', '--json'], f.options);
  assert.equal(unknown.exitCode, 1); assert.equal(decode(unknown).report.contrast.findings[0].status, 'indeterminate');
  const human = await runCliAsync(['validate', 'contrast.json', '--check-contrast'], f.options); assert.equal(human.exitCode, 0); assert.match(human.output, /indeterminate/);
  spec.slides[0].elements[0].box.width = -1; await writeFile(file, JSON.stringify(spec));
  const invalid = await runCliAsync(['validate', 'contrast.json', '--check-contrast', '--json'], f.options); assert.equal(invalid.exitCode, 1); assert.equal(decode(invalid).report.contrast.status, 'skipped-invalid-core');
  for (const args of [['validate', 'missing', '--fail-on-contrast'], ['inspect', 'missing', '--check-contrast'], ['generate', 'missing', 'out', '--check-contrast'], ['validate', 'missing', '--check-contrast', '--check-contrast']]) assert.equal((await runCliAsync(args)).exitCode, 2);
});

test('validate reports decorative actions as warnings without blocking generation', async t => {
  const f = await fixture(t), spec = { version: 'pptkit-authoring-v1', slides: [{ id: 'page', elements: [{ id: 'object', type: 'shape', shape: 'rect', box: { x: 0, y: 0, width: 100, height: 50 }, accessibility: { decorative: true }, action: { type: 'url', url: 'https://example.com' } }] }] };
  await writeFile(path.join(f.directory, 'decorative-action.json'), JSON.stringify(spec));
  const checked = await runCliAsync(['validate', 'decorative-action.json', '--json'], f.options);
  assert.equal(checked.exitCode, 0); assert.equal(decode(checked).report.valid, true);
  assert.ok(decode(checked).report.diagnostics.some(item => item.code === 'decorative-action' && item.severity === 'warning' && item.elementId === 'object'));
  const generated = await runCliAsync(['generate', 'decorative-action.json', 'decorative-action.pptx', '--json'], f.options);
  assert.equal(generated.exitCode, 0, generated.output);
});

test('explicit local assets are captured once per declaration for validation, preview and generation', async t => {
  const f = await fixture(t), image = (await import('../../pptx-parser/test/image-fixture.mjs')).png;
  await mkdir(path.join(f.directory, 'assets with spaces')); await writeFile(path.join(f.directory, 'assets with spaces', 'photo.png'), image);
  const spec = { version: 'pptkit-authoring-v1', assets: [{ id: 'img', kind: 'image', mimeType: 'image/png', source: { type: 'path', value: 'photo.png' }, width: 8, height: 4 }], slides: [{ elements: [{ type: 'image', assetId: 'img', box: { x: 0, y: 0, width: 80, height: 40 } }] }] };
  await writeFile(path.join(f.directory, 'local.json'), JSON.stringify(spec));
  const missingOptIn = await runCliAsync(['validate', 'local.json', '--json'], f.options); assert.equal(missingOptIn.exitCode, 3); assert.equal(decode(missingOptIn).error.code, 'authoring-asset-unsupported');
  const flags = ['--asset-base', 'assets with spaces', '--json'];
  const validated = await runCliAsync(['validate', 'local.json', '--export', ...flags], f.options); assert.equal(validated.exitCode, 0, validated.output); const input = decode(validated).assetInputs[0];
  assert.equal(input.path, 'assets with spaces/photo.png'); assert.equal(input.byteLength, image.length); assert.equal(input.sha256, sha(image)); assert.deepEqual(decode(validated).report.export.warnings, []);
  const preview = await runCliAsync(['preview', 'local.json', '--slide', '1', ...flags], f.options); assert.equal(preview.exitCode, 0, preview.output); assert.deepEqual(decode(preview).assetInputs, [input]); assert.match(decode(preview).report.slide.svg, /data:image\/png;base64,/);
  const generated = await runCliAsync(['generate', 'local.json', 'local.pptx', ...flags], f.options); assert.equal(generated.exitCode, 0, generated.output); assert.deepEqual(decode(generated).assetInputs, [input]);
  const session = await importPptx(await readFile(path.join(f.directory, 'local.pptx')), { projection: 'content' }); try { assert.equal(session.document.slides[0].elements[0].type, 'image'); } finally { session.dispose(); }
  await writeFile(path.join(f.directory, 'outside.png'), image); await symlink(path.join(f.directory, 'outside.png'), path.join(f.directory, 'assets with spaces', 'link.png'));
  const linked = { ...spec, assets: [{ ...spec.assets[0], source: { type: 'path', value: 'link.png' } }] }; await writeFile(path.join(f.directory, 'linked.json'), JSON.stringify(linked));
  const linkedResult = await runCliAsync(['validate', 'linked.json', ...flags], f.options); assert.equal(linkedResult.exitCode, 0); assert.equal(decode(linkedResult).assetInputs[0].sha256, sha(image));
  for (const bad of ['/absolute.png', '../photo.png', 'dir/../photo.png', 'dir\\photo.png', 'https://example.invalid/image.png', '']) {
    spec.assets[0].source.value = bad; await writeFile(path.join(f.directory, 'bad-local.json'), JSON.stringify(spec));
    const result = await runCliAsync(['generate', 'bad-local.json', 'bad-local.pptx', ...flags], f.options); assert.equal(result.exitCode, 3, bad); assert.equal(decode(result).error.code, 'authoring-asset-unsupported'); await assert.rejects(readFile(path.join(f.directory, 'bad-local.pptx')), { code: 'ENOENT' });
  }
  spec.assets[0].source.value = 'missing.png'; await writeFile(path.join(f.directory, 'missing-local.json'), JSON.stringify(spec));
  const missing = await runCliAsync(['generate', 'missing-local.json', 'missing.pptx', ...flags], f.options); assert.equal(missing.exitCode, 3); assert.equal(decode(missing).error.code, 'file-io'); await assert.rejects(readFile(path.join(f.directory, 'missing.pptx')), { code: 'ENOENT' });
  await writeFile(path.join(f.directory, 'assets with spaces', 'large.bin'), Buffer.alloc(700, 1));
  const largeAssets = { version: 'pptkit-authoring-v1', assets: ['a', 'b'].map(id => ({ id, kind: 'image', source: { type: 'path', value: 'large.bin' } })), slides: [{}] };
  await writeFile(path.join(f.directory, 'budget.json'), JSON.stringify(largeAssets));
  const budget = await runCliAsync(['generate', 'budget.json', 'budget.pptx', ...flags, '--max-input-bytes', '1200'], f.options); assert.equal(budget.exitCode, 4); assert.equal(decode(budget).error.code, 'input-limit-exceeded'); await assert.rejects(readFile(path.join(f.directory, 'budget.pptx')), { code: 'ENOENT' });
  const abort = new AbortController(); abort.abort(); assert.equal((await runCliAsync(['validate', 'local.json', ...flags], { ...f.options, signal: abort.signal })).exitCode, 130);
  for (const args of [['inspect', 'source.pptx', '--asset-base', '.'], ['preview', 'source.pptx', '--input-format', 'pptx', '--slide', '1', '--json', '--asset-base', '.']]) assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
});

test('preview returns selected SVG and renderer warnings without writing files', async t => {
  const f = await fixture(t), image = (await import('../../pptx-parser/test/image-fixture.mjs')).png;
  const spec = { version: 'pptkit-authoring-v1', presentation: { size: { width: 500, height: 250 } }, assets: [{ id: 'img', kind: 'image', source: { type: 'bytes', encoding: 'base64', value: Buffer.from(image).toString('base64') }, mimeType: 'image/png', width: 8, height: 4 }], slides: [{ id: 'one', elements: [{ type: 'text', box: { x: 10, y: 10, width: 200, height: 40 }, content: 'First page' }] }, { id: 'two', hidden: true, elements: [{ type: 'text', box: { x: 10, y: 10, width: 200, height: 40 }, content: 'Second page' }, { type: 'image', assetId: 'img', box: { x: 200, y: 10, width: 80, height: 40 } }] }] };
  const source = path.join(f.directory, 'preview.json'); await writeFile(source, JSON.stringify(spec));
  const result = await runCliAsync(['preview', 'preview.json', '--slide', '2', '--json'], f.options); assert.equal(result.exitCode, 0, result.output); const body = decode(result);
  assert.equal(body.report.scope, 'svg-static'); assert.equal(body.report.width, 500); assert.equal(body.report.height, 250); assert.equal(body.report.slide.slideId, 'two'); assert.equal(body.report.slide.index, 1); assert.equal(body.report.slide.hidden, true);
  assert.match(body.report.slide.svg, /Second page/); assert.doesNotMatch(body.report.slide.svg, /First page/); assert.match(body.report.slide.svg, /data:image\/png;base64,/); assert.ok(Array.isArray(body.report.warnings));
  assert.deepEqual(await readFile(source, 'utf8'), JSON.stringify(spec));
  const first = await runCliAsync(['preview', 'preview.json', '--slide', '1', '--json'], f.options); assert.equal(first.exitCode, 0); assert.match(decode(first).report.slide.svg, /First page/);
  const all = await runCliAsync(['preview', 'preview.json', '--slide', 'all', '--json'], f.options); assert.equal(all.exitCode, 0);
  const pages = decode(all).report.slides; assert.deepEqual(pages.map(s => [s.slideId, s.index, s.hidden]), [['one', 0, false], ['two', 1, true]]);
  assert.equal(pages[0].svg, decode(first).report.slide.svg); assert.equal(pages[1].svg, body.report.slide.svg);
  const cap = Math.max(Buffer.byteLength(first.output), Buffer.byteLength(result.output)) + 32;
  assert.ok(cap < Buffer.byteLength(all.output));
  const allLimited = await runCliAsync(['preview', 'preview.json', '--slide', 'all', '--max-output-bytes', String(cap), '--json'], f.options);
  assert.equal(allLimited.exitCode, 4); assert.equal(decode(allLimited).error.code, 'output-limit-exceeded');
  spec.assets[0].source.value = Buffer.from('not an image').toString('base64'); await writeFile(source, JSON.stringify(spec));
  const warned = await runCliAsync(['preview', 'preview.json', '--slide', '2', '--json'], f.options);
  assert.equal(warned.exitCode, 0); assert.ok(decode(warned).report.warnings.some(w => w.code === 'asset-bytes-unsupported'));
  const allWarnings = await runCliAsync(['preview', 'preview.json', '--slide', 'all', '--json'], f.options);
  assert.equal(allWarnings.exitCode, 0); assert.ok(decode(allWarnings).report.warnings.some(w => w.code === 'asset-bytes-unsupported'));
  const unselectedWarning = await runCliAsync(['preview', 'preview.json', '--slide', '1', '--json'], f.options);
  assert.equal(unselectedWarning.exitCode, 0); assert.ok(!decode(unselectedWarning).report.warnings.some(w => w.code === 'asset-bytes-unsupported'));
  const missing = await runCliAsync(['preview', 'preview.json', '--slide', '3', '--json'], f.options); assert.equal(missing.exitCode, 3); assert.equal(decode(missing).error.code, 'preview-slide-out-of-range');
  const limited = await runCliAsync(['preview', 'preview.json', '--slide', '2', '--max-output-bytes', '512', '--json'], f.options); assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
  for (const args of [['preview', 'preview.json'], ['preview', 'preview.json', '--slide', '0', '--json'], ['preview', 'preview.json', '--slide', '01', '--json'], ['preview', 'preview.json', '--slide', 'ALL', '--json'], ['preview', 'preview.json', '--slide', '1'], ['inspect', 'source.pptx', '--slide', '1', '--json']]) assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
  const controller = new AbortController(); controller.abort(); assert.equal((await runCliAsync(['preview', 'preview.json', '--slide', '1', '--json'], { ...f.options, signal: controller.signal })).exitCode, 130);
});

test('preview creates a single standalone SVG exclusively and keeps sources unchanged', async t => {
  const f = await fixture(t), spec = { version: 'pptkit-authoring-v1', slides: [{ elements: [{ type: 'text', content: 'Standalone preview', box: { x: 10, y: 10, width: 220, height: 60 } }] }] };
  const source = path.join(f.directory, 'source deck.json'), output = path.join(f.directory, 'preview slide.svg'); await writeFile(source, JSON.stringify(spec));
  const args = ['preview', 'source deck.json', '--slide', '1', '--json', '--output', 'preview slide.svg'];
  const inline = await runCliAsync(['preview', 'source deck.json', '--slide', '1', '--json'], f.options); assert.equal(inline.exitCode, 0);
  const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0, result.output); const body = decode(result);
  const svg = await readFile(output, 'utf8'); assert.match(svg, /Standalone preview/); assert.equal(svg, decode(inline).report.slide.svg); assert.equal(body.report.publication.status, 'committed'); assert.equal(body.report.publication.byteLength, Buffer.byteLength(svg)); assert.equal(body.report.slide.svg, undefined); assert.equal(body.nativeVerification, 'not-performed');
  assert.deepEqual(await readFile(source, 'utf8'), JSON.stringify(spec)); const again = await runCliAsync(args, f.options); assert.equal(again.exitCode, 3); assert.equal(decode(again).error.code, 'file-target-exists'); assert.equal(await readFile(output, 'utf8'), svg);
  await symlink(output, path.join(f.directory, 'symlink.svg')); const symlinkResult = await runCliAsync(['preview', 'source deck.json', '--slide', '1', '--json', '--output', 'symlink.svg'], f.options); assert.equal(symlinkResult.exitCode, 3); assert.equal(decode(symlinkResult).error.code, 'file-target-exists'); assert.equal(await readFile(output, 'utf8'), svg);
  assert.equal((await runCliAsync(['validate', 'source deck.json', '--output', 'bad.svg'], f.options)).exitCode, 2);
  const limited = await runCliAsync(['preview', 'source deck.json', '--slide', '1', '--json', '--output', 'limited.svg', '--max-output-bytes', '512'], f.options); assert.equal(limited.exitCode, 4); await assert.rejects(readFile(path.join(f.directory, 'limited.svg')), { code: 'ENOENT' });
  const abort = new AbortController(); abort.abort(); assert.equal((await runCliAsync(['preview', 'source deck.json', '--slide', '1', '--json', '--output', 'cancelled.svg'], { ...f.options, signal: abort.signal })).exitCode, 130); await assert.rejects(readFile(path.join(f.directory, 'cancelled.svg')), { code: 'ENOENT' });
  const absent = await runCliAsync(['preview', 'source deck.json', '--slide', '1', '--json', '--output', 'missing/slide.svg'], f.options); assert.equal(absent.exitCode, 3); assert.equal(decode(absent).error.code, 'file-io');
  const racingArgs = ['preview', 'source deck.json', '--slide', '1', '--json', '--output', 'race.svg'];
  const racing = await Promise.all([runCliAsync(racingArgs, f.options), runCliAsync(racingArgs, f.options)]); assert.deepEqual(racing.map(r => r.exitCode).sort(), [0, 3]); assert.equal(decode(racing.find(r => r.exitCode === 3)).error.code, 'file-target-exists'); assert.equal(await readFile(path.join(f.directory, 'race.svg'), 'utf8'), svg);
  const imported = await runCliAsync(['preview', 'source.pptx', '--input-format', 'pptx', '--slide', '1', '--json', '--output', 'imported.svg'], f.options); assert.equal(imported.exitCode, 0, imported.output); assert.equal(decode(imported).report.inputFormat, 'pptx'); assert.ok(decode(imported).report.importReport); assert.match(await readFile(path.join(f.directory, 'imported.svg'), 'utf8'), /<svg/);
  assert.ok(!(await readdir(f.directory)).some(name => name.startsWith('.pptkit-svg-')));
});

test('whole-deck preview creates ordered HTML with hidden slides and empty decks', async t => {
  const f = await fixture(t), spec = { version: 'pptkit-authoring-v1', slides: [{ id: 'a', elements: [{ type: 'text', content: 'First slide', box: { x: 10, y: 10, width: 200, height: 50 } }] }, { id: 'b', hidden: true, elements: [{ type: 'text', content: 'Second slide', box: { x: 10, y: 10, width: 200, height: 50 } }] }] };
  await writeFile(path.join(f.directory, 'deck.json'), JSON.stringify(spec));
  const inline = await runCliAsync(['preview', 'deck.json', '--slide', 'all', '--json'], f.options); assert.equal(inline.exitCode, 0, inline.output);
  const args = ['preview', 'deck.json', '--slide', 'all', '--output', 'deck preview.html', '--json'];
  const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0, result.output); const report = decode(result).report, html = await readFile(path.join(f.directory, 'deck preview.html'), 'utf8');
  assert.equal(report.publication.format, 'html'); assert.equal(report.publication.byteLength, Buffer.byteLength(html)); assert.deepEqual(report.slides.map(s => [s.slideId, s.index, s.hidden]), [['a', 0, false], ['b', 1, true]]); assert.ok(report.slides.every(s => s.svg === undefined));
  assert.match(html, /<!doctype html>/); assert.match(html, /data-slide="2" data-hidden="true"/); assert.ok(html.indexOf('First slide') < html.indexOf('Second slide')); assert.equal((html.match(/<iframe sandbox=""/g) ?? []).length, 2);
  for (const slide of decode(inline).report.slides) assert.ok(html.includes(slide.svg.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')));
  const repeated = await runCliAsync(args, f.options); assert.equal(repeated.exitCode, 3); assert.equal(decode(repeated).error.code, 'file-target-exists'); assert.equal(await readFile(path.join(f.directory, 'deck preview.html'), 'utf8'), html);
  const limited = await runCliAsync(['preview', 'deck.json', '--slide', 'all', '--output', 'limited.html', '--json', '--max-output-bytes', '512'], f.options); assert.equal(limited.exitCode, 4); await assert.rejects(readFile(path.join(f.directory, 'limited.html')), { code: 'ENOENT' });
  const rawSize = decode(inline).report.slides.reduce((sum, slide) => sum + Buffer.byteLength(slide.svg), 0), htmlBudget = rawSize + 1; assert.ok(Buffer.byteLength(html) > htmlBudget);
  const wrapperLimited = await runCliAsync(['preview', 'deck.json', '--slide', 'all', '--output', 'wrapper-limited.html', '--json', '--max-output-bytes', String(htmlBudget)], f.options); assert.equal(wrapperLimited.exitCode, 4); assert.equal(decode(wrapperLimited).error.code, 'output-limit-exceeded'); await assert.rejects(readFile(path.join(f.directory, 'wrapper-limited.html')), { code: 'ENOENT' });
  await writeFile(path.join(f.directory, 'empty.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [] }));
  const empty = await runCliAsync(['preview', 'empty.json', '--slide', 'all', '--output', 'empty.html', '--json'], f.options); assert.equal(empty.exitCode, 0, empty.output); assert.deepEqual(decode(empty).report.slides, []); assert.doesNotMatch(await readFile(path.join(f.directory, 'empty.html'), 'utf8'), /<section/);
  const pptx = await runCliAsync(['preview', 'source.pptx', '--input-format', 'pptx', '--slide', 'all', '--output', 'import.html', '--json'], f.options); assert.equal(pptx.exitCode, 0, pptx.output); assert.ok(decode(pptx).report.importReport); assert.match(await readFile(path.join(f.directory, 'import.html'), 'utf8'), /&lt;svg/);
  assert.ok(!(await readdir(f.directory)).some(name => name.startsWith('.pptkit-svg-')));
});

test('all-slide preview of an empty authoring document returns an empty array', async t => {
  const f = await fixture(t); await writeFile(path.join(f.directory, 'empty.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [] }));
  const all = await runCliAsync(['preview', 'empty.json', '--slide', 'all', '--json'], f.options);
  assert.equal(all.exitCode, 0, all.output); assert.deepEqual(decode(all).report.slides, []);
  const one = await runCliAsync(['preview', 'empty.json', '--slide', '1', '--json'], f.options);
  assert.equal(one.exitCode, 3); assert.equal(decode(one).error.code, 'preview-slide-out-of-range');
});

test('PPTX preview returns partial import report and selected static SVG without modifying source', async t => {
  for (const strict of [false, true]) {
    const f = await fixture(t, encodeStructure(contentFixture(undefined, strict)));
    const result = await runCliAsync(['preview', 'source.pptx', '--input-format', 'pptx', '--slide', '1', '--json'], f.options);
    assert.equal(result.exitCode, 0, result.output); const body = decode(result);
    assert.equal(body.report.inputFormat, 'pptx'); assert.equal(body.report.slide.index, 0); assert.match(body.report.slide.svg, /A &amp; B/);
    const expected = await importPptx(f.bytes, { projection: 'content' });
    try { assert.deepEqual(body.report.importReport, expected.report); assert.ok(body.report.importReport.diagnostics.some(d => d.code === 'content-projection-partial')); } finally { expected.dispose(); }
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), Buffer.from(f.bytes));
    const second = await runCliAsync(['preview', 'source.pptx', '--input-format', 'pptx', '--slide', '2', '--json'], f.options);
    assert.equal(second.exitCode, 0); assert.equal(decode(second).report.slide.index, 1); assert.doesNotMatch(decode(second).report.slide.svg, /A &amp; B/);
    const all = await runCliAsync(['preview', 'source.pptx', '--input-format', 'pptx', '--slide', 'all', '--json'], f.options);
    assert.equal(all.exitCode, 0); const allReport = decode(all).report; assert.deepEqual(allReport.slides.map(s => s.index), [0, 1]);
    assert.equal(allReport.slides[0].svg, body.report.slide.svg); assert.deepEqual(allReport.importReport, body.report.importReport);
    const limited = await runCliAsync(['preview', 'source.pptx', '--input-format', 'pptx', '--slide', '1', '--max-output-bytes', '512', '--json'], f.options);
    assert.equal(limited.exitCode, 4); assert.equal(decode(limited).error.code, 'output-limit-exceeded');
  }
  for (const args of [['preview', 'missing', '--input-format', 'guess', '--slide', '1', '--json'], ['inspect', 'missing', '--input-format', 'pptx'], ['preview', 'missing', '--input-format', 'pptx', '--input-format', 'pptx', '--slide', '1', '--json']]) assert.equal((await runCliAsync(args)).exitCode, 2);
});

test('PPTX preview leaves external-only pictures unprojected and never fetches them', async t => {
  const { imageFixture, encodeImages, picture } = await import('../../pptx-parser/test/image-fixture.mjs');
  const { relationships } = await import('../../pptx-parser/test/opc-fixture.mjs');
  const source = imageFixture(picture({ embed: null, link: 'outside' }));
  source.parts['deck/_rels/z.xml.rels'] = relationships([source.rel('layout', 'slideLayout', 'layout.xml'), { ...source.rel('outside', 'image', 'https://example.invalid/private'), TargetMode: 'External' }]);
  const f = await fixture(t, encodeImages(source)); let fetchCount = 0; const original = globalThis.fetch;
  globalThis.fetch = async () => { fetchCount++; throw new Error('unexpected request'); };
  try {
    const result = await runCliAsync(['preview', 'source.pptx', '--input-format', 'pptx', '--slide', '1', '--json'], f.options);
    assert.equal(result.exitCode, 0, result.output); assert.equal(fetchCount, 0);
    const report = decode(result).report; assert.doesNotMatch(report.slide.svg, /example\.invalid/);
    assert.ok(report.importReport.findings.some(finding => /external/i.test(finding.reason)));
  } finally { globalThis.fetch = original; }
});

test('description manifests combine with text, chart and numbering over three reopen rounds', async t => {
  for (const strict of [false, true]) {
    const ext = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="0"/></a:ext></a:extLst>';
    const textShape = shape({ id: '20', txBox: true, text: body('<a:p><a:r><a:t>Old</a:t></a:r></a:p>') }).replace('<p:cNvPr id="20" name="Object 20" descr="Description" />', `<p:cNvPr id="20" name="Object 20" descr="Description">${ext}</p:cNvPr>`);
    const source = encodeChart(chartFixture({ strict, xml: chartXml({ family: 'bar', data: series({ refs: true, values: [0, -2.5] }), external: true }), book: workbook({ values: [0, -2.5] }), content: frame() + textShape }));
    const f = await fixture(t, source); let current = source, input = 'source.pptx';
    for (const [round, number] of ['17', '0', 'clear'].entries()) {
      const inspection = await runCliAsync(['inspect', input, '--projection', 'content', '--edit-save-profile', 'text-runs-v3', '--chart-edit-save-profile', 'existing-values-v1', '--text-values', '--chart-values', '--description-values', '--name-values', '--decorative-values', '--visibility-values', '--json'], f.options);
      assert.equal(inspection.exitCode, 0, inspection.output); const data = decode(inspection), desc = data.descriptionEditCandidates, names = data.nameEditCandidates, decorative = data.decorativeEditCandidates, visibility = data.visibilityEditCandidates, text = data.textEditCandidates.targets[0], chart = data.chartEditCandidates.targets.find(x => x.pointIndex === 1);
      assert.equal(desc.sourceSha256, sha(current)); assert.equal(desc.targets.length, 2);
      const manifests = [
        ['visibility.json', { version: visibility.version, sourceSha256: visibility.sourceSha256, edits: [{ ...visibility.targets[0], replacement: !visibility.targets[0].expected }] }],
        ['decorative.json', { version: decorative.version, sourceSha256: decorative.sourceSha256, edits: decorative.targets.map(target => ({ ...target, replacement: round % 2 === 0 })) }],
        ['name.json', { version: names.version, sourceSha256: names.sourceSha256, edits: names.targets.map((target, i) => ({ ...target, replacement: `Name ${round} ${i}` })) }],
        ['desc.json', { version: desc.version, sourceSha256: desc.sourceSha256, edits: desc.targets.map((target, i) => ({ ...target, replacement: round === 2 && i === 0 ? null : `Alt ${round} & ${i}` })) }],
        ['text.json', { version: data.textEditCandidates.version, sourceSha256: sha(current), edits: [{ ...text, replacement: `Text ${round}` }] }],
        ['chart.json', { version: data.chartEditCandidates.version, sourceSha256: sha(current), edits: [{ ...chart, replacement: String(11 + round) }] }],
      ];
      for (const [name, manifest] of manifests) await writeFile(path.join(f.directory, name), JSON.stringify(manifest));
      const output = `round-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--description-edits', 'desc.json', '--name-edits', 'name.json', '--decorative-edits', 'decorative.json', '--visibility-edits', 'visibility.json', '--text-edits', 'text.json', '--chart-edits', 'chart.json', '--first-slide-number', number, '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); assert.equal(decode(dry).report.changes.length, 9); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).inputs.length, 7); assert.equal(decode(result).report.changes.length, 9);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(current), after = readGeneratedZip(next);
      for (const [name, bytes] of before) if (!['deck/main.data', 'deck/z.xml', 'deck/chart.xml', 'deck/book.xlsx'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
      const reopened = await importPptx(next, { projection: 'content', editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', chartEditSaveProfile: 'existing-values-v1', firstSlideNumberSaveProfile: 'attribute-v2' });
      try { assert.equal(reopened.document.firstSlideNumber, number === 'clear' ? undefined : Number(number)); assert.equal(reopened.document.slides[0].hidden, !visibility.targets[0].expected); assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], 11 + round); assert.equal(reopened.document.slides[0].elements[1].content[0].runs[0].text, `Text ${round}`); assert.equal(reopened.document.slides[0].elements[1].accessibility.decorative, round % 2 === 0); for (const [i, target] of desc.targets.entries()) { const object = reopened.document.slides[0].elements.find(e => e.id === target.elementId); assert.equal(object.accessibility?.description, round === 2 && i === 0 ? undefined : `Alt ${round} & ${i}`); assert.equal(object.name, `Name ${round} ${i}`); } } finally { reopened.dispose(); }
      current = next; input = output;
    }
    assert.deepEqual(await readFile(path.join(f.directory, 'source.pptx')), source);
  }
});

test('description manifests reject stale, duplicate and invalid targets without publishing', async t => {
  const f = await fixture(t), inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--description-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
  const c = decode(inspected).descriptionEditCandidates, target = c.targets[0], edit = { ...target, replacement: 'Alt' }, base = { version: c.version, sourceSha256: c.sourceSha256, edits: [edit] };
  const cases = [[{ ...base, sourceSha256: '0'.repeat(64) }, 'description-edits-source-mismatch'], [{ ...base, edits: [{ ...edit, expected: 'wrong' }] }, 'description-edits-baseline-mismatch'], [{ ...base, edits: [edit, edit] }, 'description-edits-invalid'], [{ ...base, edits: [{ ...edit, elementId: 'missing' }] }, 'description-edits-target-unbound'], [{ ...base, edits: [{ ...edit, replacement: 2 }] }, 'description-edits-invalid'], [{ ...base, extra: true }, 'description-edits-invalid']];
  for (const [manifest, code] of cases) { await writeFile(path.join(f.directory, 'desc.json'), JSON.stringify(manifest)); const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--description-edits', 'desc.json', '--json'], f.options); assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' }); }
  for (const args of [['inspect', 'missing', '--description-values', '--json'], ['inspect', 'missing', '--projection', 'content', '--description-values'], ['save', 'a', 'b', '--mode', 'no-op', '--description-edits', 'x'], ['save', 'a', 'b', '--mode', 'edit', '--description-edits', 'x', '--picture-edit', 'p', '--picture-bytes', 'i']]) assert.equal((await runCliAsync(args)).exitCode, 2);
});

test('description candidates distinguish absent, empty and removed values in nested groups', async t => {
  const content = group(shape({ id: '21', txBox: true, text: body() }) + shape({ id: '22', txBox: true, text: body() }), 10), source = encodeStructure(contentFixture(content)), f = await fixture(t, source);
  let input = 'source.pptx';
  for (const [round, replacement] of ['', null, 'New & 中文'].entries()) {
    const inspected = await runCliAsync(['inspect', input, '--projection', 'content', '--description-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).descriptionEditCandidates;
    const groupFind = decode(inspected).report.findings.filter(x => x.feature === 'drawing.description-attribute' && x.tier === 'editable');
    assert.equal(c.targets.length, groupFind.length); assert.ok(c.targets.length >= 2);
    const selected = c.targets.slice(-2); if (round === 1) assert.deepEqual(selected.map(x => x.expected), ['', '']); if (round === 2) assert.deepEqual(selected.map(x => x.expected), [null, null]);
    await writeFile(path.join(f.directory, 'desc.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: selected.map(target => ({ ...target, replacement })) }));
    const output = `nested-${round}.pptx`, result = await runCliAsync(['save', input, output, '--mode', 'edit', '--description-edits', 'desc.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).report.changes.length, 2); input = output;
  }
});

test('description edit requests reject incompatible text manifests and malformed JSON before publication', async t => {
  const f = await fixture(t), inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--description-values', '--json'], f.options), c = decode(inspected).descriptionEditCandidates;
  await writeFile(path.join(f.directory, 'desc.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...c.targets[0], replacement: 'Changed' }] }));
  const textInspect = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--edit-save-profile', 'text-runs-v1', '--text-values', '--json'], f.options); assert.equal(textInspect.exitCode, 0, textInspect.output); const text = decode(textInspect).textEditCandidates;
  await writeFile(path.join(f.directory, 'text.json'), JSON.stringify({ version: text.version, sourceSha256: text.sourceSha256, edits: [{ ...text.targets[0], replacement: 'Changed' }] }));
  const incompatible = await runCliAsync(['save', 'source.pptx', 'out.pptx', '--mode', 'edit', '--description-edits', 'desc.json', '--text-edits', 'text.json', '--json'], f.options); assert.equal(incompatible.exitCode, 3); assert.equal(decode(incompatible).error.code, 'description-edits-incompatible');
  await writeFile(path.join(f.directory, 'desc.json'), Buffer.from([0xff]));
  const malformed = await runCliAsync(['save', 'source.pptx', 'out.pptx', '--mode', 'edit', '--description-edits', 'desc.json', '--json'], f.options); assert.equal(malformed.exitCode, 3); assert.equal(decode(malformed).error.code, 'description-edits-invalid'); await assert.rejects(readFile(path.join(f.directory, 'out.pptx')), { code: 'ENOENT' });
});

test('name manifest validates source, baseline, duplicates, value types and missing bindings before publication', async t => {
  const f = await fixture(t), inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--name-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
  const c = decode(inspected).nameEditCandidates, target = c.targets[0], edit = { ...target, replacement: 'New name' }, base = { version: c.version, sourceSha256: c.sourceSha256, edits: [edit] };
  const cases = [[{ ...base, sourceSha256: '0'.repeat(64) }, 'name-edits-source-mismatch'], [{ ...base, edits: [{ ...edit, expected: 'stale' }] }, 'name-edits-baseline-mismatch'], [{ ...base, edits: [edit, edit] }, 'name-edits-invalid'], [{ ...base, edits: [{ ...edit, elementId: 'missing' }] }, 'name-edits-target-unbound'], [{ ...base, edits: [{ ...edit, replacement: null }] }, 'name-edits-invalid'], [{ ...base, unexpected: true }, 'name-edits-invalid']];
  for (const [manifest, code] of cases) { await writeFile(path.join(f.directory, 'names.json'), JSON.stringify(manifest)); const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--name-edits', 'names.json', '--json'], f.options); assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' }); }
  for (const args of [['inspect', 'missing', '--name-values', '--json'], ['inspect', 'missing', '--projection', 'content', '--name-values'], ['save', 'a', 'b', '--mode', 'no-op', '--name-edits', 'n'], ['save', 'a', 'b', '--mode', 'edit', '--name-edits', 'n', '--picture-edit', 'p', '--picture-bytes', 'i']]) assert.equal((await runCliAsync(args)).exitCode, 2);
});

test('name manifests preserve nested group edits and reject changes to missing source names', async t => {
  const inner = shape({ id: '21', txBox: true, text: body() }) + shape({ id: '22', txBox: true, text: body() }), source = encodeStructure(contentFixture(group(inner, 10))), f = await fixture(t, source);
  const inspect = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--name-values', '--json'], f.options); assert.equal(inspect.exitCode, 0, inspect.output); const c = decode(inspect).nameEditCandidates; assert.ok(c.targets.length >= 2);
  await writeFile(path.join(f.directory, 'names.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: c.targets.slice(-2).map((x, i) => ({ ...x, replacement: `Nested ${i} & 中文` })) }));
  const saved = await runCliAsync(['save', 'source.pptx', 'nested.pptx', '--mode', 'edit', '--name-edits', 'names.json', '--json'], f.options); assert.equal(saved.exitCode, 0, saved.output); assert.equal(decode(saved).report.changes.length, 2);
  const after = await runCliAsync(['inspect', 'nested.pptx', '--projection', 'content', '--name-values', '--json'], f.options); assert.equal(after.exitCode, 0, after.output); assert.deepEqual(decode(after).nameEditCandidates.targets.slice(-2).map(x => x.expected), ['Nested 0 & 中文', 'Nested 1 & 中文']);
  const absent = encodeStructure(contentFixture(shape({ id: '2' }).replace(' name="Object 2"', ''))), missing = await fixture(t, absent), result = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--name-values', '--json'], missing.options); assert.equal(result.exitCode, 0, result.output); assert.equal(decode(result).nameEditCandidates.targets.length, 0); assert.equal(decode(result).nameEditCandidates.unboundCount, 1);
});

test('decorative manifests validate boolean, baseline, source and binding without publication on errors', async t => {
  const ext = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>';
  const object = shape({ id: '2' }).replace('<p:cNvPr id="2" name="Object 2" descr="Description" />', `<p:cNvPr id="2" name="Object 2" descr="Description">${ext}</p:cNvPr>`), f = await fixture(t, encodeStructure(contentFixture(object)));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--decorative-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).decorativeEditCandidates, target = c.targets[0]; assert.equal(target.expected, true);
  const edit = { ...target, replacement: false }, base = { version: c.version, sourceSha256: c.sourceSha256, edits: [edit] };
  const cases = [[{ ...base, sourceSha256: '0'.repeat(64) }, 'decorative-edits-source-mismatch'], [{ ...base, edits: [{ ...edit, expected: false }] }, 'decorative-edits-baseline-mismatch'], [{ ...base, edits: [edit, edit] }, 'decorative-edits-invalid'], [{ ...base, edits: [{ ...edit, elementId: 'missing' }] }, 'decorative-edits-target-unbound'], [{ ...base, edits: [{ ...edit, replacement: 'false' }] }, 'decorative-edits-invalid'], [{ ...base, edits: [{ ...edit, replacement: null }] }, 'decorative-edits-invalid']];
  for (const [manifest, code] of cases) { await writeFile(path.join(f.directory, 'dec.json'), JSON.stringify(manifest)); const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--decorative-edits', 'dec.json', '--json'], f.options); assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' }); }
  await writeFile(path.join(f.directory, 'dec.json'), JSON.stringify(base)); const saved = await runCliAsync(['save', 'source.pptx', 'good.pptx', '--mode', 'edit', '--decorative-edits', 'dec.json', '--json'], f.options); assert.equal(saved.exitCode, 0, saved.output);
  const reopened = await runCliAsync(['inspect', 'good.pptx', '--projection', 'content', '--decorative-values', '--json'], f.options); assert.equal(reopened.exitCode, 0, reopened.output); assert.equal(decode(reopened).decorativeEditCandidates.targets[0].expected, false);
  for (const args of [['inspect', 'missing', '--decorative-values', '--json'], ['inspect', 'missing', '--projection', 'content', '--decorative-values'], ['save', 'a', 'b', '--mode', 'no-op', '--decorative-edits', 'x'], ['save', 'a', 'b', '--mode', 'edit', '--decorative-edits', 'x', '--picture-edit', 'p', '--picture-bytes', 'i']]) assert.equal((await runCliAsync(args)).exitCode, 2);
  const absent = await fixture(t, encodeStructure(contentFixture(shape({ id: '2' })))); const noBinding = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--decorative-values', '--json'], absent.options); assert.equal(noBinding.exitCode, 0, noBinding.output); assert.equal(decode(noBinding).decorativeEditCandidates.targets.length, 0); assert.equal(decode(noBinding).decorativeEditCandidates.unboundCount, 1);
});

test('decorative manifests update distinct children in one nested group atomically', async t => {
  const ext = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>';
  const make = id => shape({ id }).replace(`<p:cNvPr id="${id}" name="Object ${id}" descr="Description" />`, `<p:cNvPr id="${id}" name="Object ${id}" descr="Description">${ext}</p:cNvPr>`);
  const f = await fixture(t, encodeStructure(contentFixture(group(make('21') + make('22'), 10))));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--decorative-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).decorativeEditCandidates; assert.equal(c.targets.length, 2); assert.ok(c.targets.every(target => target.expected === true));
  await writeFile(path.join(f.directory, 'dec.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: c.targets.map(target => ({ ...target, replacement: false })) }));
  const saved = await runCliAsync(['save', 'source.pptx', 'nested.pptx', '--mode', 'edit', '--decorative-edits', 'dec.json', '--json'], f.options); assert.equal(saved.exitCode, 0, saved.output); assert.equal(decode(saved).report.changes.length, 2);
  const reopened = await runCliAsync(['inspect', 'nested.pptx', '--projection', 'content', '--decorative-values', '--json'], f.options); assert.equal(reopened.exitCode, 0, reopened.output); assert.ok(decode(reopened).decorativeEditCandidates.targets.every(target => target.expected === false));
});

test('visibility manifests save multiple slides and reject stale, duplicate, invalid or unbound edits', async t => {
  for (const strict of [false, true]) {
    const f = await fixture(t, encodeStructure(contentFixture(undefined, strict))), inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--visibility-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
    const c = decode(inspected).visibilityEditCandidates; assert.deepEqual(c.targets.map(x => x.expected), [true, false]); assert.equal(c.unboundCount, 0);
    const edits = c.targets.map(target => ({ ...target, replacement: !target.expected })), base = { version: c.version, sourceSha256: c.sourceSha256, edits };
    const cases = [[{ ...base, sourceSha256: '0'.repeat(64) }, 'visibility-edits-source-mismatch'], [{ ...base, edits: [{ ...edits[0], expected: false }] }, 'visibility-edits-baseline-mismatch'], [{ ...base, edits: [edits[0], edits[0]] }, 'visibility-edits-invalid'], [{ ...base, edits: [{ ...edits[0], slideId: 'missing' }] }, 'visibility-edits-target-unbound'], [{ ...base, edits: [{ ...edits[0], replacement: 'false' }] }, 'visibility-edits-invalid'], [{ ...base, extra: true }, 'visibility-edits-invalid']];
    for (const [manifest, code] of cases) { await writeFile(path.join(f.directory, 'visibility.json'), JSON.stringify(manifest)); const r = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--visibility-edits', 'visibility.json', '--json'], f.options); assert.equal(r.exitCode, 3, r.output); assert.equal(decode(r).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' }); }
    await writeFile(path.join(f.directory, 'visibility.json'), JSON.stringify(base)); const args = ['save', 'source.pptx', 'good.pptx', '--mode', 'edit', '--visibility-edits', 'visibility.json', '--json'];
    const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); assert.equal(decode(dry).report.changes.length, 2); await assert.rejects(readFile(path.join(f.directory, 'good.pptx')), { code: 'ENOENT' });
    const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output); const reopened = await runCliAsync(['inspect', 'good.pptx', '--projection', 'content', '--visibility-values', '--json'], f.options); assert.equal(reopened.exitCode, 0, reopened.output); assert.deepEqual(decode(reopened).visibilityEditCandidates.targets.map(x => x.expected), [false, true]);
  }
  for (const args of [['inspect', 'missing', '--visibility-values', '--json'], ['inspect', 'missing', '--projection', 'content', '--visibility-values'], ['save', 'a', 'b', '--mode', 'no-op', '--visibility-edits', 'x'], ['save', 'a', 'b', '--mode', 'edit', '--visibility-edits', 'x', '--picture-edit', 'p', '--picture-bytes', 'i']]) assert.equal((await runCliAsync(args)).exitCode, 2);
});

test('visibility candidates reject duplicate and inconsistent report or slide identities', () => {
  const slide = { id: 's', hidden: false, setHidden() {} }, binding = { slideId: 's', status: 'bound', source: { part: '/s.xml' } }, session = { document: { slides: [slide] }, report: { slideVisibilitySources: [binding] } };
  assert.deepEqual(visibilityCandidates(session, 'hash').targets, [{ slideId: 's', expected: false }]);
  session.report.slideVisibilitySources = [binding, binding]; assert.throws(() => visibilityCandidates(session, 'hash'), error => error.code === 'visibility-candidates-inconsistent');
  session.report.slideVisibilitySources = [binding]; session.document.slides = [slide, { ...slide }]; assert.throws(() => visibilityCandidates(session, 'hash'), error => error.code === 'visibility-candidates-inconsistent');
  session.document.slides = [{ ...slide, setHidden: undefined }]; assert.throws(() => visibilityCandidates(session, 'hash'), error => error.code === 'visibility-candidates-inconsistent');
  const controller = new AbortController(); controller.abort(); assert.throws(() => visibilityCandidates(session, 'hash', controller.signal), error => error.code === 'operation-aborted');
});

test('metadata manifests save bound fields over three rounds without rewriting other package parts', async t => {
  for (const strict of [false, true]) {
    const sourceFixture = metadataFixture(strict); sourceFixture.add('core', '<d:title>Original</d:title><d:creator>Alice</d:creator><d:subject/><c:revision>7</c:revision>');
    const f = await fixture(t, encodeMetadata(sourceFixture)); let input = 'source.pptx', source = f.bytes;
    for (let round = 0; round < 3; round++) {
      const inspected = await runCliAsync(['inspect', input, '--projection', 'content', '--metadata-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).metadataEditCandidates;
      assert.deepEqual(c.targets.map(x => x.field), ['title', 'author', 'subject']); assert.equal(c.unboundCount, 2); assert.equal(c.sourceSha256, sha(source));
      await writeFile(path.join(f.directory, 'metadata.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: c.targets.map(target => ({ ...target, replacement: `${target.field} ${round} & 中文` })) }));
      const output = `metadata-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--metadata-edits', 'metadata.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); assert.equal(decode(dry).report.changes.length, 3); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output); assert.equal(decode(saved).report.changes.length, 3);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next); for (const [name, bytes] of before) if (name !== 'inventory/core.data') assert.deepEqual(after.get(name), bytes, name);
      const reopened = await importPptx(next, { projection: 'content' }); try { for (const target of c.targets) assert.equal(reopened.document.metadata[target.field], `${target.field} ${round} & 中文`); assert.equal(reopened.document.metadata.revision, 7); } finally { reopened.dispose(); } source = next; input = output;
    }
  }
});

test('metadata manifests reject stale, duplicate, unsupported and unbound fields before publication', async t => {
  const fdata = metadataFixture(); fdata.add('core', '<d:title>Old</d:title><d:creator>Alice</d:creator><d:description>Mixed<!--opaque-->value</d:description>'); const f = await fixture(t, encodeMetadata(fdata));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--metadata-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).metadataEditCandidates, edit = { ...c.targets[0], replacement: 'New' }, base = { version: c.version, sourceSha256: c.sourceSha256, edits: [edit] };
  const cases = [[{ ...base, sourceSha256: '0'.repeat(64) }, 'metadata-edits-source-mismatch'], [{ ...base, edits: [{ ...edit, expected: 'stale' }] }, 'metadata-edits-baseline-mismatch'], [{ ...base, edits: [edit, edit] }, 'metadata-edits-invalid'], [{ ...base, edits: [{ field: 'description', expected: 'Mixedvalue', replacement: 'New' }] }, 'metadata-edits-target-unbound'], [{ ...base, edits: [{ field: 'revision', expected: '1', replacement: '2' }] }, 'metadata-edits-invalid'], [{ ...base, edits: [{ ...edit, replacement: null }] }, 'metadata-edits-invalid'], [{ ...base, extra: true }, 'metadata-edits-invalid']];
  for (const [manifest, code] of cases) { await writeFile(path.join(f.directory, 'metadata.json'), JSON.stringify(manifest)); const r = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--metadata-edits', 'metadata.json', '--json'], f.options); assert.equal(r.exitCode, 3, r.output); assert.equal(decode(r).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' }); }
  for (const args of [['inspect', 'missing', '--metadata-values', '--json'], ['inspect', 'missing', '--projection', 'content', '--metadata-values'], ['save', 'a', 'b', '--mode', 'no-op', '--metadata-edits', 'x'], ['save', 'a', 'b', '--mode', 'edit', '--metadata-edits', 'x', '--picture-edit', 'p', '--picture-bytes', 'i']]) assert.equal((await runCliAsync(args)).exitCode, 2);
});

test('metadata, visibility, text and object descriptions share a CLI transaction', async t => {
  for (const strict of [false, true]) {
    const data = metadataFixture(strict); data.add('core', '<d:title>Original</d:title><d:creator>Alice</d:creator><c:revision>7</c:revision>');
    data.parts['deck/z.xml'] = data.p('sld', `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="Root"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shape({ id: '2', txBox: true, text: body() })}</p:spTree></p:cSld>`, `xmlns:a="${data.A}" show="0"`);
    const f = await fixture(t, encodeMetadata(data)), inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--edit-save-profile', 'text-runs-v3', '--text-values', '--description-values', '--name-values', '--visibility-values', '--metadata-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
    const report = decode(inspected), manifest = (c, edits) => ({ version: c.version, sourceSha256: c.sourceSha256, edits });
    const metadata = report.metadataEditCandidates, visibility = report.visibilityEditCandidates, description = report.descriptionEditCandidates, names = report.nameEditCandidates, text = report.textEditCandidates;
    const files = [['metadata.json', manifest(metadata, [{ ...metadata.targets.find(x => x.field === 'title'), replacement: 'Mixed title' }])], ['visibility.json', manifest(visibility, [{ ...visibility.targets[0], replacement: false }])], ['description.json', manifest(description, [{ ...description.targets[0], replacement: 'Mixed alt' }])], ['name.json', manifest(names, [{ ...names.targets[0], replacement: 'Mixed name' }])], ['text.json', manifest(text, [{ ...text.targets[0], replacement: 'Mixed text' }])]];
    for (const [name, value] of files) await writeFile(path.join(f.directory, name), JSON.stringify(value));
    const args = ['save', 'source.pptx', 'mixed.pptx', '--mode', 'edit', '--metadata-edits', 'metadata.json', '--visibility-edits', 'visibility.json', '--description-edits', 'description.json', '--name-edits', 'name.json', '--text-edits', 'text.json', '--first-slide-number', '17', '--json'];
    const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output); assert.equal(decode(saved).report.changes.length, 6);
    const next = await readFile(path.join(f.directory, 'mixed.pptx')), before = readGeneratedZip(f.bytes), after = readGeneratedZip(next); for (const [name, bytes] of before) if (!['inventory/core.data', 'deck/main.data', 'deck/z.xml'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
    const reopened = await importPptx(next, { projection: 'content' }); try { assert.equal(reopened.document.metadata.title, 'Mixed title'); assert.equal(reopened.document.metadata.author, 'Alice'); assert.equal(reopened.document.metadata.revision, 7); assert.equal(reopened.document.firstSlideNumber, 17); assert.equal(reopened.document.slides[0].hidden, false); assert.equal(reopened.document.slides[0].elements[0].name, 'Mixed name'); assert.equal(reopened.document.slides[0].elements[0].accessibility.description, 'Mixed alt'); assert.equal(reopened.document.slides[0].elements[0].content[0].runs[0].text, 'Mixed text'); } finally { reopened.dispose(); }
  }
});

test('metadata candidates reject ambiguous source identity and unavailable values', () => {
  const binding = { field: 'title', status: 'bound', source: { part: '/core.xml' } }, session = { document: { metadata: { title: 'A' }, setMetadata() {} }, report: { coreMetadataSources: [binding] } };
  assert.deepEqual(metadataCandidates(session, 'hash').targets, [{ field: 'title', expected: 'A' }]);
  session.report.coreMetadataSources = [binding, binding]; assert.throws(() => metadataCandidates(session, 'hash'), error => error.code === 'metadata-candidates-inconsistent');
  session.report.coreMetadataSources = [{ ...binding, field: 'revision' }]; assert.throws(() => metadataCandidates(session, 'hash'), error => error.code === 'metadata-candidates-inconsistent');
  session.report.coreMetadataSources = [binding]; session.document.metadata = {}; assert.throws(() => metadataCandidates(session, 'hash'), error => error.code === 'metadata-candidates-inconsistent');
  const controller = new AbortController(); controller.abort(); assert.throws(() => metadataCandidates(session, 'hash', controller.signal), error => error.code === 'operation-aborted');
});


test('custom property manifests preserve source identity over three rounds and merge with core edits', async t => {
  for (const strict of [false, true]) {
    const sourceFixture = metadataFixture(strict);
    sourceFixture.add('core', '<d:title>Original</d:title>');
    sourceFixture.add('custom', custom(2, 'lpwstr', 'old', 'name="_x005F_x0041_"') + custom(3, 'bool', 'false', 'name="Flag"') + custom(4, 'i4', '1', 'name="Number"') + custom(5, 'r8', '-0', 'name="Zero"') + custom(6, 'filetime', '2020-01-02T03:04:05Z', 'name="Date"') + custom(7, 'vector', '<v:lpwstr>x</v:lpwstr>', 'name="Opaque"'));
    const f = await fixture(t, encodeMetadata(sourceFixture)); let input = 'source.pptx', source = f.bytes;
    for (let round = 0; round < 3; round++) {
      const inspect = await runCliAsync(['inspect', input, '--projection', 'content', '--custom-property-values', '--metadata-values', '--json'], f.options);
      assert.equal(inspect.exitCode, 0, inspect.output); const report = decode(inspect), candidates = report.customPropertyEditCandidates;
      assert.equal(candidates.targets.length, 5); assert.equal(candidates.unboundCount, 1); assert.equal(candidates.sourceSha256, sha(source));
      assert.equal(candidates.targets[0].expected.name, '_x0041_');
      const replacements = { string: `round ${round}`, boolean: round % 2 ? 'false' : 'true', integer: String(round + 5), real: String(round + 0.5), date: `2020-01-0${round + 3}T03:04:05.000Z` };
      await writeFile(path.join(f.directory, 'custom.json'), JSON.stringify({ version: candidates.version, sourceSha256: candidates.sourceSha256, edits: candidates.targets.map(x => ({ ...x, replacement: replacements[x.expected.type] })) }));
      const core = report.metadataEditCandidates;
      await writeFile(path.join(f.directory, 'core.json'), JSON.stringify({ version: core.version, sourceSha256: core.sourceSha256, edits: [{ ...core.targets[0], replacement: `title ${round}` }] }));
      const output = `custom-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--custom-property-edits', 'custom.json', '--metadata-edits', 'core.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const result = await runCliAsync(args, f.options); assert.equal(result.exitCode, 0, result.output);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next);
      for (const [name, bytes] of before) if (!['inventory/custom.data', 'inventory/core.data'].includes(name)) assert.deepEqual(after.get(name), bytes, name);
      const session = await importPptx(next, { projection: 'content' });
      try { assert.equal(session.document.metadata.title, `title ${round}`); assert.equal(session.document.metadata.customProperties[0].value, `round ${round}`); assert.equal(session.document.metadata.customProperties[0].name, '_x0041_'); } finally { session.dispose(); }
      input = output; source = next;
    }
  }
});

test('custom property manifests reject stale, duplicate, unbound and invalid requests before publication', async t => {
  const data = metadataFixture(); data.add('custom', custom(2, 'lpwstr', 'old', 'name="A"') + custom(3, 'i4', '1', 'name="N"'));
  const f = await fixture(t, encodeMetadata(data));
  const inspect = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--json'], f.options); assert.equal(inspect.exitCode, 0, inspect.output);
  const c = decode(inspect).customPropertyEditCandidates, target = { ...c.targets[0], replacement: 'new' }, number = { ...c.targets[1], replacement: '1.5' };
  const base = { version: c.version, sourceSha256: c.sourceSha256, edits: [target] };
  const cases = [[{ ...base, sourceSha256: '0'.repeat(64) }, 'custom-property-edits-source-mismatch'], [{ ...base, edits: [{ ...target, expected: { ...target.expected, value: 'stale' } }] }, 'custom-property-edits-baseline-mismatch'], [{ ...base, edits: [target, target] }, 'custom-property-edits-invalid'], [{ ...base, edits: [{ ...target, sourceObjectId: 'unknown' }] }, 'custom-property-edits-target-unbound'], [{ ...base, edits: [number] }, 'custom-property-edits-invalid'], [{ ...base, edits: [{ ...target, replacement: null }] }, 'custom-property-edits-invalid']];
  for (const [manifest, code] of cases) {
    await writeFile(path.join(f.directory, 'custom.json'), JSON.stringify(manifest));
    const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--custom-property-edits', 'custom.json', '--json'], f.options);
    assert.equal(result.exitCode, 3, result.output); assert.equal(decode(result).error.code, code); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
  for (const args of [['inspect', 'missing', '--custom-property-values', '--json'], ['inspect', 'missing', '--projection', 'content', '--custom-property-values'], ['save', 'a', 'b', '--mode', 'no-op', '--custom-property-edits', 'x'], ['save', 'a', 'b', '--mode', 'edit', '--custom-property-edits', 'x', '--picture-edit', 'p', '--picture-bytes', 'i']]) assert.equal((await runCliAsync(args)).exitCode, 2);
});


test('custom property v2 manifests rename and edit values through three guarded rounds', async t => {
  for (const strict of [false, true]) {
    const data = metadataFixture(strict); data.add('custom', custom(2, 'lpwstr', 'old', 'name="First"') + custom(3, 'bool', 'false', 'name="Second"') + custom(4, 'vector', '<v:lpwstr>opaque</v:lpwstr>', 'name="Opaque"'));
    const f = await fixture(t, encodeMetadata(data)); let input = 'source.pptx', source = f.bytes;
    for (let round = 0; round < 3; round++) {
      const inspected = await runCliAsync(['inspect', input, '--projection', 'content', '--custom-property-values', '--custom-property-names', '--json'], f.options);
      assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).customPropertyEditCandidates;
      assert.equal(c.version, 'pptkit-custom-property-edits-v2'); assert.equal(c.targets.length, 2);
      const edits = [{ ...c.targets[0], replacementName: `Name ${round} _x0041_`, replacement: `Text ${round}` }, { ...c.targets[1], replacementName: `Flag ${round}` }];
      await writeFile(path.join(f.directory, 'custom-v2.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits }));
      const output = `rename-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--custom-property-edits', 'custom-v2.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next);
      for (const [name, bytes] of before) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), bytes, name);
      const reopened = await importPptx(next, { projection: 'content' });
      try { assert.deepEqual(reopened.document.metadata.customProperties, [{ name: `Name ${round} _x0041_`, type: 'string', value: `Text ${round}` }, { name: `Flag ${round}`, type: 'boolean', value: false }]); } finally { reopened.dispose(); }
      source = next; input = output;
    }
  }
});

test('custom property v2 rejects malformed, stale, duplicate and colliding names before publication', async t => {
  const data = metadataFixture(); data.add('custom', custom(2, 'lpwstr', 'a', 'name="A"') + custom(3, 'lpwstr', 'b', 'name="B"') + custom(4, 'vector', '<v:lpwstr>opaque</v:lpwstr>', 'name="Opaque"'));
  const f = await fixture(t, encodeMetadata(data));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--custom-property-names', '--json'], f.options);
  assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).customPropertyEditCandidates, target = c.targets[0];
  const base = { version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...target, replacementName: 'New' }] };
  const cases = [
    { ...base, sourceSha256: '0'.repeat(64) }, { ...base, edits: [{ ...target }] },
    { ...base, edits: [{ ...target, replacementName: '' }] }, { ...base, edits: [{ ...target, replacementName: 'B' }] }, { ...base, edits: [{ ...target, replacementName: 'Opaque' }] },
    { ...base, edits: [{ ...target, replacementName: 'X' }, { ...target, replacementName: 'Y' }] },
    { ...base, edits: [{ ...target, expected: { ...target.expected, name: 'stale' }, replacementName: 'New' }] },
    { ...base, version: 'pptkit-custom-property-edits-v1' },
  ];
  for (const manifest of cases) {
    await writeFile(path.join(f.directory, 'custom-v2.json'), JSON.stringify(manifest));
    const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--custom-property-edits', 'custom-v2.json', '--json'], f.options);
    assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
  assert.equal((await runCliAsync(['inspect', 'missing', '--custom-property-names', '--json'])).exitCode, 2);
});


test('CLI authoring exports custom values and reports unsupported XML strings as input errors', async t => {
  const f = await fixture(t);
  const valid = { version: 'pptkit-authoring-v1', presentation: { metadata: { customProperties: [{ name: 'Text', type: 'string', value: '_x0041_ & 中文' }, { name: 'Flag', type: 'boolean', value: false }] } }, slides: [] };
  await writeFile(path.join(f.directory, 'custom-authoring.json'), JSON.stringify(valid));
  const result = await runCliAsync(['generate', 'custom-authoring.json', 'custom-authored.pptx', '--json'], f.options);
  assert.equal(result.exitCode, 0, result.output);
  const session = await importPptx(await readFile(path.join(f.directory, 'custom-authored.pptx')), { projection: 'structure' });
  try { assert.deepEqual(session.document.metadata.customProperties, valid.presentation.metadata.customProperties); } finally { session.dispose(); }
  for (const value of ['bad\u0000value', '_x005F_x0041_']) {
    const input = structuredClone(valid); input.presentation.metadata.customProperties[0].value = value;
    await writeFile(path.join(f.directory, 'bad-custom.json'), JSON.stringify(input));
    for (const args of [['generate', 'bad-custom.json', 'bad-custom.pptx', '--json'], ['validate', 'bad-custom.json', '--export', '--json']]) {
      const rejected = await runCliAsync(args, f.options); assert.equal(rejected.exitCode, 3, rejected.output); assert.equal(decode(rejected).error.code, 'authoring-export-invalid');
    }
    await assert.rejects(readFile(path.join(f.directory, 'bad-custom.pptx')), { code: 'ENOENT' });
  }
});


test('rename candidates exclude sources with undecidable opaque names while value edits remain available', async t => {
  const data = metadataFixture(); data.add('custom', custom(2, 'lpwstr', 'old', 'name="A"') + custom(3, 'vector', '<v:lpwstr>x</v:lpwstr>', 'name="_x0041_"') + custom(4, 'lpwstr', 'safe', 'name="B"'));
  const f = await fixture(t, encodeMetadata(data));
  const v2 = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--custom-property-names', '--json'], f.options);
  assert.equal(v2.exitCode, 0, v2.output); assert.deepEqual(decode(v2).customPropertyEditCandidates.targets, []);
  const v1 = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--json'], f.options);
  assert.equal(v1.exitCode, 0, v1.output); const c = decode(v1).customPropertyEditCandidates; assert.equal(c.targets.length, 1);
  await writeFile(path.join(f.directory, 'custom.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...c.targets[0], replacement: 'new' }] }));
  const result = await runCliAsync(['save', 'source.pptx', 'value.pptx', '--mode', 'edit', '--custom-property-edits', 'custom.json', '--json'], f.options);
  assert.equal(result.exitCode, 0, result.output);
  const saved = await importPptx(await readFile(path.join(f.directory, 'value.pptx')), { projection: 'content' });
  try { assert.deepEqual(saved.document.metadata.customProperties, [{ name: 'B', type: 'string', value: 'new' }]); } finally { saved.dispose(); }
});


test('custom property v3 deletes explicit source identities alongside rename and value changes', async t => {
  for (const strict of [false, true]) {
    const data = metadataFixture(strict); data.add('custom', custom(2, 'lpwstr', 'a', 'name="A"') + custom(3, 'bool', 'false', 'name="B"') + custom(4, 'i4', '7', 'name="C"') + custom(5, 'vector', '<v:lpwstr>opaque</v:lpwstr>', 'name="Opaque"'));
    const f = await fixture(t, encodeMetadata(data)); let input = 'source.pptx', source = f.bytes;
    for (let round = 0; round < 3; round++) {
      const inspected = await runCliAsync(['inspect', input, '--projection', 'content', '--custom-property-values', '--custom-property-deletions', '--json'], f.options);
      assert.equal(inspected.exitCode, 0, inspected.output); const candidates = decode(inspected).customPropertyEditCandidates;
      assert.equal(candidates.version, 'pptkit-custom-property-edits-v3'); assert.equal(candidates.targets.length, 3 - round);
      const edits = [{ ...candidates.targets[0], delete: true }];
      if (round === 0) edits.push({ ...candidates.targets[1], replacementName: 'Renamed', replacement: 'true' });
      await writeFile(path.join(f.directory, 'remove.json'), JSON.stringify({ version: candidates.version, sourceSha256: candidates.sourceSha256, edits }));
      const output = `remove-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--custom-property-edits', 'remove.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next);
      for (const [name, bytes] of before) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), bytes, name);
      assert.match(after.get('inventory/custom.data').toString(), /<v:vector><v:lpwstr>opaque<\/v:lpwstr><\/v:vector>/);
      input = output; source = next;
    }
    const s = await importPptx(source, { projection: 'content' }); try { assert.deepEqual(s.document.metadata.customProperties, []); assert.deepEqual(s.report.inspection.metadata.custom.properties.map(p => p.name), ['Opaque']); } finally { s.dispose(); }
  }
});

test('v3 deletion rejects stale, duplicate, unbound and mixed per-target operations before publication', async t => {
  const data = metadataFixture(); data.add('custom', custom(2, 'lpwstr', 'a', 'name="A"') + custom(3, 'vector', '<v:lpwstr>x</v:lpwstr>', 'name="Opaque"'));
  const f = await fixture(t, encodeMetadata(data));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--custom-property-deletions', '--json'], f.options);
  assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).customPropertyEditCandidates, target = c.targets[0], deletion = { ...target, delete: true };
  const base = { version: c.version, sourceSha256: c.sourceSha256, edits: [deletion] };
  const cases = [{ ...base, sourceSha256: '0'.repeat(64) }, { ...base, edits: [deletion, deletion] }, { ...base, edits: [{ ...deletion, expected: { ...target.expected, value: 'stale' } }] }, { ...base, edits: [{ ...deletion, sourceObjectId: 'unknown' }] }, { ...base, edits: [{ ...deletion, replacement: 'bad' }] }, { ...base, edits: [{ ...target, delete: false }] }, { ...base, version: 'pptkit-custom-property-edits-v2' }];
  for (const manifest of cases) {
    await writeFile(path.join(f.directory, 'remove.json'), JSON.stringify(manifest));
    const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--custom-property-edits', 'remove.json', '--json'], f.options);
    assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
  assert.equal((await runCliAsync(['inspect', 'missing', '--custom-property-deletions', '--json'])).exitCode, 2);
});


test('custom property v4 appends typed entries and combines deletion and rename in three rounds', async t => {
  for (const strict of [false, true]) {
    const data = metadataFixture(strict); data.add('custom', custom(2, 'lpwstr', 'a', 'name="A"') + custom(3, 'bool', 'false', 'name="B"') + custom(4, 'vector', '<v:lpwstr>opaque</v:lpwstr>', 'name="Opaque"'));
    const f = await fixture(t, encodeMetadata(data)); let input = 'source.pptx', source = f.bytes;
    for (let round = 0; round < 3; round++) {
      const inspected = await runCliAsync(['inspect', input, '--projection', 'content', '--custom-property-values', '--custom-property-additions', '--json'], f.options);
      assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).customPropertyEditCandidates;
      assert.equal(c.version, 'pptkit-custom-property-edits-v4'); assert.equal(c.insertion.status, 'available');
      const edits = round === 0 ? [{ ...c.targets[0], delete: true }, { ...c.targets[1], replacementName: 'Renamed', replacement: 'true' }] : [];
      const additions = round === 0 ? [{ name: 'Text', type: 'string', value: '中文 & _x0041_' }, { name: 'Number', type: 'integer', value: '42' }, { name: 'Real', type: 'real', value: '-0' }, { name: 'Date', type: 'date', value: '2020-01-02T03:04:05.000Z' }] : [{ name: `Later ${round}`, type: 'boolean', value: 'false' }];
      await writeFile(path.join(f.directory, 'append.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits, additions }));
      const output = `append-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--custom-property-edits', 'append.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next);
      for (const [name, bytes] of before) if (name !== 'inventory/custom.data') assert.deepEqual(after.get(name), bytes, name);
      const reopened = await importPptx(next, { projection: 'content' });
      try {
        const properties = reopened.report.inspection.metadata.custom.properties;
        assert.equal(properties.filter(p => p.name === 'Opaque').length, 1); assert.equal(new Set(properties.map(p => p.id)).size, properties.length);
        if (round === 0) { assert.equal(properties.find(p => p.name === 'Renamed').id, 3); assert.ok(properties.filter(p => ['Text', 'Number', 'Real', 'Date'].includes(p.name)).every(p => p.id > 4)); }
        assert.equal(reopened.document.metadata.customProperties.find(p => p.name === 'Text').value, '中文 & _x0041_');
      } finally { reopened.dispose(); }
      source = next; input = output;
    }
  }
});

test('custom property v4 rejects unavailable, stale, invalid and conflicting insertions before publication', async t => {
  const data = metadataFixture(); data.add('custom', custom(2, 'vector', '<v:lpwstr>opaque</v:lpwstr>', 'name="Opaque"'));
  const f = await fixture(t, encodeMetadata(data));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--custom-property-additions', '--json'], f.options);
  assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).customPropertyEditCandidates;
  const base = { version: c.version, sourceSha256: c.sourceSha256, edits: [], additions: [{ name: 'New', type: 'integer', value: '3' }] };
  const cases = [{ ...base, sourceSha256: '0'.repeat(64) }, { ...base, additions: [{ name: 'N', type: 'integer', value: '1.5' }] }, { ...base, additions: [{ name: 'N', type: 'unknown', value: 'x' }] }, { ...base, additions: [{ name: 'N', type: 'date', value: 'bad' }] }, { ...base, additions: [{ name: 'N', type: 'boolean', value: '1' }] }, { ...base, additions: [{ name: 'N', type: 'string', value: 'x', extra: true }] }];
  for (const manifest of cases) {
    await writeFile(path.join(f.directory, 'append.json'), JSON.stringify(manifest));
    const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--custom-property-edits', 'append.json', '--json'], f.options);
    assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
  await writeFile(path.join(f.directory, 'append.json'), JSON.stringify({ ...base, additions: [{ name: 'Opaque', type: 'string', value: 'x' }] }));
  const collision = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--custom-property-edits', 'append.json', '--json'], f.options);
  assert.equal(collision.exitCode, 3, collision.output); assert.equal(decode(collision).error.code, 'save-edit-unsupported'); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  assert.equal((await runCliAsync(['inspect', 'missing', '--custom-property-additions', '--json'])).exitCode, 2);
  const absentData = metadataFixture(), absent = await fixture(t, encodeMetadata(absentData));
  const noPart = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--custom-property-additions', '--json'], absent.options);
  assert.equal(noPart.exitCode, 0, noPart.output); const noPartCandidates = decode(noPart).customPropertyEditCandidates; assert.equal(noPartCandidates.insertion.status, 'unavailable');
  await writeFile(path.join(absent.directory, 'append.json'), JSON.stringify({ version: noPartCandidates.version, sourceSha256: noPartCandidates.sourceSha256, edits: [], additions: [{ name: 'New', type: 'string', value: 'x' }] }));
  const rejected = await runCliAsync(['save', 'source.pptx', 'missing-part.pptx', '--mode', 'edit', '--custom-property-edits', 'append.json', '--json'], absent.options);
  assert.equal(rejected.exitCode, 3, rejected.output); assert.equal(decode(rejected).error.code, 'custom-property-edits-target-unbound'); await assert.rejects(readFile(path.join(absent.directory, 'missing-part.pptx')), { code: 'ENOENT' });
});

test('custom property v5 creates a missing part, then appends across three reopens', async t => {
  for (const strict of [false, true]) {
    const f = await fixture(t, encodeMetadata(metadataFixture(strict))); let input = 'source.pptx', source = f.bytes;
    for (let round = 0; round < 3; round++) {
      const inspected = await runCliAsync(['inspect', input, '--projection', 'content', '--custom-property-values', '--custom-property-additions', '--custom-property-create', '--json'], f.options);
      assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).customPropertyEditCandidates;
      assert.equal(c.version, 'pptkit-custom-property-edits-v5'); assert.equal(c.insertion.status, 'available');
      const additions = round === 0 ? [{ name: 'Text', type: 'string', value: '中文 & _x0041_' }, { name: 'Number', type: 'integer', value: '42' }] : [{ name: `Flag ${round}`, type: 'boolean', value: 'true' }];
      await writeFile(path.join(f.directory, 'create.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: [], additions }));
      const output = `create-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--custom-property-edits', 'create.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next);
      for (const [name, value] of before) if (!(round === 0 ? ['_rels/.rels', '[Content_Types].xml'] : ['pptkit-custom-properties-1.xml']).includes(name)) assert.deepEqual(after.get(name), value, name);
      const reopened = await importPptx(next, { projection: 'content' });
      try { assert.equal(reopened.document.metadata.customProperties.length, round + 2); assert.equal(reopened.document.metadata.customProperties[0].value, '中文 & _x0041_'); }
      finally { reopened.dispose(); }
      source = next; input = output;
    }
  }
});

test('v5 creation rejects stale, invalid and v4 manifests before publication', async t => {
  const f = await fixture(t, encodeMetadata(metadataFixture()));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--custom-property-additions', '--custom-property-create', '--json'], f.options);
  assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).customPropertyEditCandidates;
  const base = { version: c.version, sourceSha256: c.sourceSha256, edits: [], additions: [{ name: 'New', type: 'string', value: 'x' }] };
  for (const manifest of [{ ...base, sourceSha256: '0'.repeat(64) }, { ...base, version: 'pptkit-custom-property-edits-v4' }, { ...base, additions: [{ name: 'New', type: 'integer', value: '01' }] }, { ...base, additions: [{ name: 'New', type: 'string', value: '_x005F_x0041_' }] }]) {
    await writeFile(path.join(f.directory, 'create.json'), JSON.stringify(manifest));
    const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--custom-property-edits', 'create.json', '--json'], f.options);
    assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
  assert.equal((await runCliAsync(['inspect', 'missing', '--custom-property-create', '--json'])).exitCode, 2);
});

test('v5 retains v4 rename, deletion and append semantics on existing parts', async t => {
  const data = metadataFixture(); data.add('custom', custom(2, 'lpwstr', 'a', 'name="A"') + custom(3, 'bool', 'false', 'name="B"'));
  const f = await fixture(t, encodeMetadata(data));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--custom-property-additions', '--custom-property-create', '--json'], f.options);
  assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).customPropertyEditCandidates;
  await writeFile(path.join(f.directory, 'mixed.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...c.targets[0], delete: true }, { ...c.targets[1], replacementName: 'Renamed', replacement: 'true' }], additions: [{ name: 'New', type: 'integer', value: '42' }] }));
  const result = await runCliAsync(['save', 'source.pptx', 'mixed.pptx', '--mode', 'edit', '--custom-property-edits', 'mixed.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output);
  const reopened = await importPptx(await readFile(path.join(f.directory, 'mixed.pptx')), { projection: 'content' });
  try { assert.deepEqual(reopened.document.metadata.customProperties, [{ name: 'Renamed', type: 'boolean', value: true }, { name: 'New', type: 'integer', value: 42 }]); }
  finally { reopened.dispose(); }
});

test('metadata v2 explicitly saves revision with title across three rounds', async t => {
  for (const strict of [false, true]) {
    const data = metadataFixture(strict); data.add('core', '<d:title>Original</d:title><c:revision>0</c:revision>');
    const f = await fixture(t, encodeMetadata(data)); let input = 'source.pptx', source = f.bytes;
    for (const [round, value] of [1, 42, Number.MAX_SAFE_INTEGER].entries()) {
      const result = await runCliAsync(['inspect', input, '--projection', 'content', '--metadata-values', '--metadata-revision', '--json'], f.options);
      assert.equal(result.exitCode, 0, result.output); const c = decode(result).metadataEditCandidates; assert.equal(c.version, 'pptkit-metadata-edits-v2');
      const revision = c.targets.find(x => x.field === 'revision'), title = c.targets.find(x => x.field === 'title'); assert.equal(typeof revision.expected, 'number');
      await writeFile(path.join(f.directory, 'meta-v2.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...revision, replacement: value }, { ...title, replacement: `title ${round}` }] }));
      const output = `revision-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--metadata-edits', 'meta-v2.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output);
      const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next);
      for (const [name, payload] of before) if (name !== 'inventory/core.data') assert.deepEqual(after.get(name), payload, name);
      const reopened = await importPptx(next, { projection: 'content' }); try { assert.equal(reopened.document.metadata.revision, value); assert.equal(reopened.document.metadata.title, `title ${round}`); } finally { reopened.dispose(); }
      input = output; source = next;
    }
  }
});

test('metadata v2 rejects stale, unbound and invalid revision values without publication', async t => {
  const data = metadataFixture(); data.add('core', '<c:revision>7</c:revision>'); const f = await fixture(t, encodeMetadata(data));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--metadata-values', '--metadata-revision', '--json'], f.options);
  assert.equal(inspected.exitCode, 0, inspected.output); const c = decode(inspected).metadataEditCandidates, target = c.targets.find(x => x.field === 'revision');
  const base = { version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...target, replacement: 8 }] };
  const cases = [{ ...base, sourceSha256: '0'.repeat(64) }, { ...base, edits: [{ ...target, expected: 6, replacement: 8 }] }, { ...base, version: 'pptkit-metadata-edits-v1' }, ...[-1, 1.5, 9007199254740992, '8', null].map(replacement => ({ ...base, edits: [{ ...target, replacement }] }))];
  for (const manifest of cases) { await writeFile(path.join(f.directory, 'bad.json'), JSON.stringify(manifest)); const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--metadata-edits', 'bad.json', '--json'], f.options); assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' }); }
  assert.equal((await runCliAsync(['inspect', 'missing', '--metadata-revision', '--json'])).exitCode, 2);
  const absent = await fixture(t, encodeMetadata(metadataFixture())); const report = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--metadata-values', '--metadata-revision', '--json'], absent.options);
  assert.equal(report.exitCode, 0, report.output); const absentCandidates = decode(report).metadataEditCandidates; assert.ok(!absentCandidates.targets.some(x => x.field === 'revision'));
  await writeFile(path.join(absent.directory, 'bad.json'), JSON.stringify({ version: absentCandidates.version, sourceSha256: absentCandidates.sourceSha256, edits: [{ field: 'revision', expected: 1, replacement: 2 }] }));
  const rejected = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--metadata-edits', 'bad.json', '--json'], absent.options); assert.equal(rejected.exitCode, 3, rejected.output); await assert.rejects(readFile(path.join(absent.directory, 'bad.pptx')), { code: 'ENOENT' });
});

test('metadata v3 saves aggregate keywords and revision over three reopens', async t => {
  for (const strict of [false, true]) {
    const data = metadataFixture(strict); data.add('core', '<c:keywords>alpha, beta</c:keywords><c:revision>7</c:revision>');
    const f = await fixture(t, encodeMetadata(data)); let input = 'source.pptx', source = f.bytes;
    for (const [round, value] of ['中文 & test', '', 'last'].entries()) {
      const result = await runCliAsync(['inspect', input, '--projection', 'content', '--metadata-values', '--metadata-keywords', '--json'], f.options); assert.equal(result.exitCode, 0, result.output);
      const c = decode(result).metadataEditCandidates; assert.equal(c.version, 'pptkit-metadata-edits-v3'); const keyword = c.targets.find(x => x.field === 'keywords'), revision = c.targets.find(x => x.field === 'revision');
      await writeFile(path.join(f.directory, 'meta-v3.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...keyword, replacement: value }, { ...revision, replacement: round + 8 }] }));
      const output = `keywords-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--metadata-edits', 'meta-v3.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output); const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next);
      for (const [name, payload] of before) if (name !== 'inventory/core.data') assert.deepEqual(after.get(name), payload, name);
      const reopened = await importPptx(next, { projection: 'content' }); try { assert.deepEqual(reopened.document.metadata.keywords, [value]); assert.equal(reopened.document.metadata.revision, round + 8); } finally { reopened.dispose(); }
      source = next; input = output;
    }
  }
});

test('metadata v3 blocks stale, old-version, invalid and segmented keyword edits', async t => {
  const data = metadataFixture(); data.add('core', '<c:keywords>alpha</c:keywords>'); const f = await fixture(t, encodeMetadata(data));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--metadata-values', '--metadata-keywords', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
  const c = decode(inspected).metadataEditCandidates, target = c.targets.find(x => x.field === 'keywords'), base = { version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...target, replacement: 'new' }] };
  for (const manifest of [{ ...base, sourceSha256: '0'.repeat(64) }, { ...base, version: 'pptkit-metadata-edits-v2' }, { ...base, edits: [{ ...target, expected: 'stale', replacement: 'new' }] }, { ...base, edits: [{ ...target, replacement: ['new'] }] }, { ...base, edits: [base.edits[0], base.edits[0]] }]) {
    await writeFile(path.join(f.directory, 'bad.json'), JSON.stringify(manifest)); const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--metadata-edits', 'bad.json', '--json'], f.options); assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
  assert.equal((await runCliAsync(['inspect', 'missing', '--metadata-keywords', '--json'])).exitCode, 2);
  const segmented = metadataFixture(); segmented.add('core', '<c:keywords><c:value>alpha</c:value></c:keywords>'); const g = await fixture(t, encodeMetadata(segmented));
  const report = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--metadata-values', '--metadata-keywords', '--json'], g.options); assert.equal(report.exitCode, 0, report.output); const unbound = decode(report).metadataEditCandidates; assert.ok(!unbound.targets.some(x => x.field === 'keywords'));
  await writeFile(path.join(g.directory, 'bad.json'), JSON.stringify({ version: unbound.version, sourceSha256: unbound.sourceSha256, edits: [{ field: 'keywords', expected: 'alpha', replacement: 'new' }] }));
  const rejected = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--metadata-edits', 'bad.json', '--json'], g.options); assert.equal(rejected.exitCode, 3, rejected.output); await assert.rejects(readFile(path.join(g.directory, 'bad.pptx')), { code: 'ENOENT' });
});

test('metadata v4 saves company and core fields over three reopens', async t => {
  for (const strict of [false, true]) {
    const data = metadataFixture(strict); data.add('core', '<d:title>Old</d:title><c:revision>7</c:revision>'); data.add('extended', '<Company>Old Company</Company><Manager>Sentinel</Manager>');
    const f = await fixture(t, encodeMetadata(data)); let input = 'source.pptx', source = f.bytes;
    for (const [round, value] of ['中文 & company', '', 'Final'].entries()) {
      const inspected = await runCliAsync(['inspect', input, '--projection', 'content', '--metadata-values', '--metadata-company', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
      const c = decode(inspected).metadataEditCandidates; assert.equal(c.version, 'pptkit-metadata-edits-v4'); const company = c.targets.find(x => x.field === 'company'), title = c.targets.find(x => x.field === 'title');
      await writeFile(path.join(f.directory, 'meta-v4.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...company, replacement: value }, { ...title, replacement: `title ${round}` }] }));
      const output = `company-${round}.pptx`, args = ['save', input, output, '--mode', 'edit', '--metadata-edits', 'meta-v4.json', '--json'];
      const dry = await runCliAsync([...args, '--dry-run'], f.options); assert.equal(dry.exitCode, 0, dry.output); await assert.rejects(readFile(path.join(f.directory, output)), { code: 'ENOENT' });
      const saved = await runCliAsync(args, f.options); assert.equal(saved.exitCode, 0, saved.output); const next = await readFile(path.join(f.directory, output)), before = readGeneratedZip(source), after = readGeneratedZip(next);
      for (const [name, payload] of before) if (!['inventory/core.data', 'inventory/extended.data'].includes(name)) assert.deepEqual(after.get(name), payload, name);
      assert.match(after.get('inventory/extended.data').toString(), /<Manager>Sentinel<\/Manager>/);
      const reopened = await importPptx(next, { projection: 'content' }); try { assert.equal(reopened.document.metadata.company, value); assert.equal(reopened.document.metadata.title, `title ${round}`); } finally { reopened.dispose(); }
      source = next; input = output;
    }
  }
});

test('metadata v4 rejects stale, old-version, invalid and unbound company edits', async t => {
  const data = metadataFixture(); data.add('extended', '<Company>Old</Company>'); const f = await fixture(t, encodeMetadata(data));
  const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--metadata-values', '--metadata-company', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
  const c = decode(inspected).metadataEditCandidates, target = c.targets.find(x => x.field === 'company'), base = { version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...target, replacement: 'New' }] };
  for (const manifest of [{ ...base, sourceSha256: '0'.repeat(64) }, { ...base, version: 'pptkit-metadata-edits-v3' }, { ...base, edits: [{ ...target, expected: 'stale', replacement: 'New' }] }, { ...base, edits: [{ ...target, replacement: null }] }, { ...base, edits: [base.edits[0], base.edits[0]] }]) {
    await writeFile(path.join(f.directory, 'bad.json'), JSON.stringify(manifest)); const result = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--metadata-edits', 'bad.json', '--json'], f.options); assert.equal(result.exitCode, 3, result.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
  assert.equal((await runCliAsync(['inspect', 'missing', '--metadata-company', '--json'])).exitCode, 2);
  const absent = await fixture(t, encodeMetadata(metadataFixture())), report = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--metadata-values', '--metadata-company', '--json'], absent.options); assert.equal(report.exitCode, 0, report.output); const c2 = decode(report).metadataEditCandidates; assert.ok(!c2.targets.some(x => x.field === 'company'));
  await writeFile(path.join(absent.directory, 'bad.json'), JSON.stringify({ version: c2.version, sourceSha256: c2.sourceSha256, edits: [{ field: 'company', expected: '', replacement: 'New' }] }));
  const rejected = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--metadata-edits', 'bad.json', '--json'], absent.options); assert.equal(rejected.exitCode, 3, rejected.output); await assert.rejects(readFile(path.join(absent.directory, 'bad.pptx')), { code: 'ENOENT' });
});

test('custom integer variants retain source types and reject narrower-range overflow in CLI', async t => {
  for (const strict of [false, true]) {
    const data = metadataFixture(strict); data.add('custom', custom(2, 'i1', '-128', 'name="Small"') + custom(3, 'ui2', '65535', 'name="Unsigned"') + custom(4, 'i8', '2147483647', 'name="WideType"'));
    const f = await fixture(t, encodeMetadata(data));
    const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
    const c = decode(inspected).customPropertyEditCandidates, edits = c.targets.map(item => ({ ...item, replacement: item.expected.name === 'Small' ? '127' : item.expected.name === 'Unsigned' ? '0' : '-2147483648' }));
    await writeFile(path.join(f.directory, 'integer.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits }));
    const result = await runCliAsync(['save', 'source.pptx', 'out.pptx', '--mode', 'edit', '--custom-property-edits', 'integer.json', '--json'], f.options); assert.equal(result.exitCode, 0, result.output);
    const output = readGeneratedZip(await readFile(path.join(f.directory, 'out.pptx'))).get('inventory/custom.data').toString(); assert.match(output, /<v:i1>127<\/v:i1>/); assert.match(output, /<v:ui2>0<\/v:ui2>/); assert.match(output, /<v:i8>-2147483648<\/v:i8>/);
    await writeFile(path.join(f.directory, 'bad.json'), JSON.stringify({ version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...c.targets.find(item => item.expected.name === 'Small'), replacement: '128' }] }));
    const blocked = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--custom-property-edits', 'bad.json', '--json'], f.options); assert.equal(blocked.exitCode, 3, blocked.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
});

test('custom r4 manifest retains binary32 tag and rejects inexact replacements', async t => {
  for (const strict of [false,true]) {
    const data = metadataFixture(strict); data.add('custom', custom(2, 'r4', '1.25', 'name="Single"')); const f = await fixture(t, encodeMetadata(data));
    const inspected = await runCliAsync(['inspect', 'source.pptx', '--projection', 'content', '--custom-property-values', '--json'], f.options); assert.equal(inspected.exitCode, 0, inspected.output);
    const c = decode(inspected).customPropertyEditCandidates, target = c.targets[0], base = { version: c.version, sourceSha256: c.sourceSha256, edits: [{ ...target, replacement: '2.5' }] };
    await writeFile(path.join(f.directory, 'r4.json'), JSON.stringify(base)); const saved = await runCliAsync(['save', 'source.pptx', 'out.pptx', '--mode', 'edit', '--custom-property-edits', 'r4.json', '--json'], f.options); assert.equal(saved.exitCode, 0, saved.output);
    assert.match(readGeneratedZip(await readFile(path.join(f.directory, 'out.pptx'))).get('inventory/custom.data').toString(), /<v:r4>2.5<\/v:r4>/);
    await writeFile(path.join(f.directory, 'r4.json'), JSON.stringify({ ...base, edits: [{ ...target, replacement: '0.1' }] })); const blocked = await runCliAsync(['save', 'source.pptx', 'bad.pptx', '--mode', 'edit', '--custom-property-edits', 'r4.json', '--json'], f.options); assert.equal(blocked.exitCode, 3, blocked.output); await assert.rejects(readFile(path.join(f.directory, 'bad.pptx')), { code: 'ENOENT' });
  }
});

test('custom bstr manifest preserves type and rejects ambiguous escapes', async t => {
  for (const strict of [false,true]) {
    const data=metadataFixture(strict); data.add('custom',custom(2,'bstr','original','name="Basic"')); const f=await fixture(t,encodeMetadata(data));
    const inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--custom-property-values','--json'],f.options); assert.equal(inspected.exitCode,0,inspected.output);
    const c=decode(inspected).customPropertyEditCandidates, target=c.targets[0], base={version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:'中文 & _x0041_'}]};
    await writeFile(path.join(f.directory,'bstr.json'),JSON.stringify(base)); const saved=await runCliAsync(['save','source.pptx','out.pptx','--mode','edit','--custom-property-edits','bstr.json','--json'],f.options); assert.equal(saved.exitCode,0,saved.output);
    const output=await readFile(path.join(f.directory,'out.pptx')); assert.match(readGeneratedZip(output).get('inventory/custom.data').toString(),/<v:bstr>/);
    const reopened=await importPptx(output,{projection:'content'}); try { assert.equal(reopened.document.metadata.customProperties[0].value,'中文 & _x0041_'); } finally { reopened.dispose(); }
    await writeFile(path.join(f.directory,'bstr.json'),JSON.stringify({...base,edits:[{...target,replacement:'_x005F_x0041_'}]})); const blocked=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--custom-property-edits','bstr.json','--json'],f.options); assert.equal(blocked.exitCode,3,blocked.output); await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});
  }
});

test('custom lpstr manifest preserves type and rejects ambiguous escapes', async t => {
  for (const strict of [false,true]) {
    const data=metadataFixture(strict); data.add('custom',custom(2,'lpstr','original','name="Basic"')); const f=await fixture(t,encodeMetadata(data));
    const inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--custom-property-values','--json'],f.options); assert.equal(inspected.exitCode,0,inspected.output);
    const c=decode(inspected).customPropertyEditCandidates, target=c.targets[0], base={version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:'中文 & _x0041_'}]};
    await writeFile(path.join(f.directory,'lpstr.json'),JSON.stringify(base)); const saved=await runCliAsync(['save','source.pptx','out.pptx','--mode','edit','--custom-property-edits','lpstr.json','--json'],f.options); assert.equal(saved.exitCode,0,saved.output);
    const output=await readFile(path.join(f.directory,'out.pptx')); assert.match(readGeneratedZip(output).get('inventory/custom.data').toString(),/<v:lpstr>/);
    const reopened=await importPptx(output,{projection:'content'}); try { assert.equal(reopened.document.metadata.customProperties[0].value,'中文 & _x0041_'); } finally { reopened.dispose(); }
    await writeFile(path.join(f.directory,'lpstr.json'),JSON.stringify({...base,edits:[{...target,replacement:'_x005F_x0041_'}]})); const blocked=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--custom-property-edits','lpstr.json','--json'],f.options); assert.equal(blocked.exitCode,3,blocked.output); await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});
  }
});

test('custom date variant manifest keeps source type and rejects invalid dates', async t => {
  for(const strict of [false,true]) {
    const data=metadataFixture(strict); data.add('custom',custom(2,'date','2020-02-29T00:00:00.120000+08:00','name="Date"')); const f=await fixture(t,encodeMetadata(data));
    const inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--custom-property-values','--json'],f.options); assert.equal(inspected.exitCode,0,inspected.output);
    const c=decode(inspected).customPropertyEditCandidates, target=c.targets[0]; assert.equal(target.expected.value,'2020-02-28T16:00:00.120Z'); const base={version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:'2024-02-29T00:00:00.123Z'}]};
    await writeFile(path.join(f.directory,'date.json'),JSON.stringify(base)); const saved=await runCliAsync(['save','source.pptx','out.pptx','--mode','edit','--custom-property-edits','date.json','--json'],f.options); assert.equal(saved.exitCode,0,saved.output);
    assert.match(readGeneratedZip(await readFile(path.join(f.directory,'out.pptx'))).get('inventory/custom.data').toString(),/<v:date>2024-02-29T00:00:00.123Z<\/v:date>/);
    await writeFile(path.join(f.directory,'date.json'),JSON.stringify({...base,edits:[{...target,replacement:'2024-02-30T00:00:00.000Z'}]})); const rejected=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--custom-property-edits','date.json','--json'],f.options); assert.equal(rejected.exitCode,3,rejected.output); await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});
  }
});


test('metadata v5 inserts missing fields, including normalized defaults, across three saves', async t => {
  for (const strict of [false, true]) {
    const data=metadataFixture(strict); data.add('core','<c:revision>7</c:revision><d:creator>Existing</d:creator>');
    const f=await fixture(t,encodeMetadata(data)); let input='source.pptx', source=f.bytes;
    for (let round=0;round<3;round++) {
      const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-insert','--json'],f.options); assert.equal(inspected.exitCode,0,inspected.output);
      const c=decode(inspected).metadataEditCandidates; assert.equal(c.version,'pptkit-metadata-edits-v5');
      const title=c.targets.find(x=>x.field==='title'), subject=c.targets.find(x=>x.field==='subject'), author=c.targets.find(x=>x.field==='author');
      if(round===0) {assert.equal(title.expected,null); assert.equal(subject.expected,null);}
      const edits=[{...title,replacement:round===0?'PPTKit Presentation':`Title ${round}`},{...subject,replacement:round===0?'':`Subject ${round}`},{...author,replacement:`Author ${round}`}];
      await writeFile(path.join(f.directory,'v5.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits}));
      const output=`insert-${round}.pptx`, args=['save',input,output,'--mode','edit','--metadata-edits','v5.json','--json'];
      const dry=await runCliAsync([...args,'--dry-run'],f.options); assert.equal(dry.exitCode,0,dry.output); await assert.rejects(readFile(path.join(f.directory,output)),{code:'ENOENT'});
      const saved=await runCliAsync(args,f.options); assert.equal(saved.exitCode,0,saved.output);
      const next=await readFile(path.join(f.directory,output)), before=readGeneratedZip(source), after=readGeneratedZip(next);
      for(const [name,payload] of before) if(name!=='inventory/core.data') assert.deepEqual(after.get(name),payload,name);
      const reopened=await importPptx(next,{projection:'content'}); try {assert.equal(reopened.document.metadata.title,edits[0].replacement);assert.equal(reopened.document.metadata.subject,edits[1].replacement);assert.equal(reopened.document.metadata.author,edits[2].replacement);} finally {reopened.dispose();}
      source=next;input=output;
    }
  }
});

test('metadata v5 rejects stale, unbound, duplicate and malformed insertion plans',async t=>{
  const data=metadataFixture(); data.add('core','<d:creator>Existing</d:creator>'); const f=await fixture(t,encodeMetadata(data));
  const inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-insert','--json'],f.options); assert.equal(inspected.exitCode,0,inspected.output);
  const c=decode(inspected).metadataEditCandidates, title=c.targets.find(x=>x.field==='title'), author=c.targets.find(x=>x.field==='author'), edit={...title,replacement:'New'}, base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
  for(const manifest of [{...base,sourceSha256:'0'.repeat(64)},{...base,version:'pptkit-metadata-edits-v4'},{...base,edits:[{...edit,expected:''}]},{...base,edits:[{...author,expected:null,replacement:'New'}]},{...base,edits:[edit,edit]},{...base,edits:[{...edit,replacement:null}]}]) {
    await writeFile(path.join(f.directory,'invalid-v5.json'),JSON.stringify(manifest)); const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','invalid-v5.json','--json'],f.options); assert.equal(result.exitCode,3,result.output); await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});
  }
  assert.equal((await runCliAsync(['inspect','missing','--metadata-insert','--json'])).exitCode,2);
  const absent=await fixture(t,encodeMetadata(metadataFixture()));
  const report=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-insert','--json'],absent.options); assert.equal(report.exitCode,0,report.output); assert.ok(!decode(report).metadataEditCandidates.targets.some(x=>x.field==='title'));
});


test('metadata v5 can explicitly insert only a normalized default',async t=>{
  const data=metadataFixture(); data.add('core','<c:revision>7</c:revision>'); const f=await fixture(t,encodeMetadata(data));
  const inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-insert','--json'],f.options); assert.equal(inspected.exitCode,0,inspected.output);
  const c=decode(inspected).metadataEditCandidates, target=c.targets.find(x=>x.field==='title'); assert.equal(target.expected,null);
  await writeFile(path.join(f.directory,'default.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:'PPTKit Presentation'}]}));
  const saved=await runCliAsync(['save','source.pptx','out.pptx','--mode','edit','--metadata-edits','default.json','--json'],f.options); assert.equal(saved.exitCode,0,saved.output);
  const reopened=await importPptx(await readFile(path.join(f.directory,'out.pptx')),{projection:'content',metadataSaveProfile:'core-insert-v5'}); try {assert.equal(reopened.document.metadata.title,'PPTKit Presentation');assert.equal(reopened.report.coreMetadataSources.find(x=>x.field==='title').status,'bound');} finally {reopened.dispose();}
});


test('custom manifest can replace a property identity without changing its projected value', async t => {
  const data=metadataFixture();data.add('custom',custom(2,'lpwstr','same','name="A"'));const f=await fixture(t,encodeMetadata(data));
  const report=await runCliAsync(['inspect','source.pptx','--projection','content','--custom-property-values','--custom-property-additions','--json'],f.options);assert.equal(report.exitCode,0,report.output);
  const c=decode(report).customPropertyEditCandidates;
  await writeFile(path.join(f.directory,'replace.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...c.targets[0],delete:true}],additions:[{name:'A',type:'string',value:'same'}]}));
  const result=await runCliAsync(['save','source.pptx','out.pptx','--mode','edit','--custom-property-edits','replace.json','--json'],f.options);assert.equal(result.exitCode,0,result.output);
  const reopened=await importPptx(await readFile(path.join(f.directory,'out.pptx')),{projection:'content'});
  try {assert.deepEqual(reopened.document.metadata.customProperties,[{name:'A',type:'string',value:'same'}]);assert.deepEqual(reopened.report.inspection.metadata.custom.properties.map(p=>p.id),[3]);} finally {reopened.dispose();}
});


test('custom decimal CLI preserves exact values and rejects inexact replacements',async t=>{
 for(const strict of [false,true]) {
  const data=metadataFixture(strict);data.add('custom',custom(2,'decimal','1.25','name="Decimal"'));const f=await fixture(t,encodeMetadata(data));
  const report=await runCliAsync(['inspect','source.pptx','--projection','content','--custom-property-values','--json'],f.options);assert.equal(report.exitCode,0,report.output);
  const c=decode(report).customPropertyEditCandidates,target=c.targets[0],base={version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:'2.5'}]};
  await writeFile(path.join(f.directory,'decimal.json'),JSON.stringify(base));const saved=await runCliAsync(['save','source.pptx','out.pptx','--mode','edit','--custom-property-edits','decimal.json','--json'],f.options);assert.equal(saved.exitCode,0,saved.output);assert.match(readGeneratedZip(await readFile(path.join(f.directory,'out.pptx'))).get('inventory/custom.data').toString(),/<v:decimal>2.5<\/v:decimal>/);
  await writeFile(path.join(f.directory,'decimal.json'),JSON.stringify({...base,edits:[{...target,replacement:'0.1'}]}));const bad=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--custom-property-edits','decimal.json','--json'],f.options);assert.equal(bad.exitCode,3,bad.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});
 }
});


test('metadata v6 saves dates and missing title over three reopens in both dialects',async t=>{
 for(const strict of [false,true]) {
  const data=metadataFixture(strict);data.add('core','<t:created xsi:type="t:W3CDTF">2020-02-29T10:00:00.1200+08:00</t:created><t:modified xsi:type="t:W3CDTF">2021-01-02T03:04:05Z</t:modified><c:revision>7</c:revision>');
  const f=await fixture(t,encodeMetadata(data));let input='source.pptx',source=f.bytes;
  for(let round=0;round<3;round++){
   const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-dates','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);
   const c=decode(inspected).metadataEditCandidates;assert.equal(c.version,'pptkit-metadata-edits-v6');const created=c.targets.find(x=>x.field==='createdAt'),title=c.targets.find(x=>x.field==='title');if(round===0){assert.equal(created.expected,'2020-02-29T02:00:00.120Z');assert.equal(title.expected,null);}
   const replacement=`2022-01-0${round+1}T00:00:00.000Z`;await writeFile(path.join(f.directory,'v6.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...created,replacement},{...title,replacement:`title ${round}` }]}));
   const output=`date-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','v6.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],f.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(f.directory,output)),{code:'ENOENT'});
   const saved=await runCliAsync(args,f.options);assert.equal(saved.exitCode,0,saved.output);const next=await readFile(path.join(f.directory,output)),before=readGeneratedZip(source),after=readGeneratedZip(next);for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);
   const reopened=await importPptx(next,{projection:'content'});try{assert.equal(reopened.document.metadata.createdAt,replacement);assert.equal(reopened.document.metadata.modifiedAt,'2021-01-02T03:04:05.000Z');assert.equal(reopened.document.metadata.title,`title ${round}`);}finally{reopened.dispose();}input=output;source=next;
  }
 }
});

test('metadata v6 rejects noncanonical dates, stale source and unbound date insertions',async t=>{
 const data=metadataFixture();data.add('core','<t:created xsi:type="t:W3CDTF">2020-01-01T00:00:00Z</t:created>');const f=await fixture(t,encodeMetadata(data));
 const report=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-dates','--json'],f.options);assert.equal(report.exitCode,0,report.output);const c=decode(report).metadataEditCandidates,target=c.targets.find(x=>x.field==='createdAt'),edit={...target,replacement:'2022-01-01T00:00:00.000Z'},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,sourceSha256:'0'.repeat(64)},{...base,version:'pptkit-metadata-edits-v5'},{...base,edits:[{...edit,expected:'stale'}]},{...base,edits:[{...edit,replacement:'2021-02-29T00:00:00.000Z'}]},{...base,edits:[{...edit,replacement:'2022-01-01T00:00:00Z'}]},{...base,edits:[{field:'modifiedAt',expected:null,replacement:'2022-01-01T00:00:00.000Z'}]},{...base,edits:[edit,edit]}]){
  await writeFile(path.join(f.directory,'invalid-v6.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','invalid-v6.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});
 }
 assert.equal((await runCliAsync(['inspect','missing','--metadata-dates','--json'])).exitCode,2);
});


test('metadata v7 inserts missing dates with typed source elements and three reopens',async t=>{
 for(const strict of [false,true]){
  const data=metadataFixture(strict);data.add('core','<c:revision>7</c:revision>');const f=await fixture(t,encodeMetadata(data));let input='source.pptx',source=f.bytes;
  for(let round=0;round<3;round++){
   const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-date-insert','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates;assert.equal(c.version,'pptkit-metadata-edits-v7');const created=c.targets.find(x=>x.field==='createdAt'),modified=c.targets.find(x=>x.field==='modifiedAt');if(round===0){assert.equal(created.expected,null);assert.equal(modified.expected,null);}
   const date=`2022-01-0${round+1}T00:00:00.000Z`;await writeFile(path.join(f.directory,'v7.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...created,replacement:date},{...modified,replacement:`2023-01-0${round+1}T00:00:00.000Z`}]}));
   const output=`date-insert-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','v7.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],f.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(f.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,f.options);assert.equal(saved.exitCode,0,saved.output);
   const next=await readFile(path.join(f.directory,output)),before=readGeneratedZip(source),after=readGeneratedZip(next);for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);const reopened=await importPptx(next,{projection:'content'});try{assert.equal(reopened.document.metadata.createdAt,date);}finally{reopened.dispose();}source=next;input=output;
  }
 }
});

test('metadata v7 rejects stale and invalid insertion manifests without publication',async t=>{
 const data=metadataFixture();data.add('core','');const f=await fixture(t,encodeMetadata(data)),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-date-insert','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);
 const c=decode(inspected).metadataEditCandidates,target=c.targets.find(x=>x.field==='createdAt'),edit={...target,replacement:'2020-01-01T00:00:00.000Z'},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,sourceSha256:'0'.repeat(64)},{...base,version:'pptkit-metadata-edits-v6'},{...base,edits:[{...edit,expected:'2020-01-01T00:00:00.000Z'}]},{...base,edits:[{...edit,replacement:'2021-02-29T00:00:00.000Z'}]},{...base,edits:[edit,edit]}]){await writeFile(path.join(f.directory,'bad-v7.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-v7.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});}
 assert.equal((await runCliAsync(['inspect','missing','--metadata-date-insert','--json'])).exitCode,2);
});


test('metadata v8 creates core and custom parts atomically across three CLI reopens',async t=>{
 for(const strict of [false,true]){
  const f=await fixture(t,encodeMetadata(metadataFixture(strict)));let input='source.pptx',source=f.bytes;
  for(let round=0;round<3;round++){
   const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-create','--custom-property-values','--custom-property-additions','--custom-property-create','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);
   const outputReport=decode(inspected),c=outputReport.metadataEditCandidates,customCandidates=outputReport.customPropertyEditCandidates;assert.equal(c.version,'pptkit-metadata-edits-v8');const title=c.targets.find(x=>x.field==='title'),created=c.targets.find(x=>x.field==='createdAt');if(round===0){assert.equal(title.expected,null);assert.equal(created.expected,null);}
   const replacement=`2020-01-0${round+1}T00:00:00.000Z`;await writeFile(path.join(f.directory,'v8.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...title,replacement:round===0?'PPTKit Presentation':`Title ${round}`},{...created,replacement}]}));
   const customArgs=[];if(round===0){await writeFile(path.join(f.directory,'v8-custom.json'),JSON.stringify({version:customCandidates.version,sourceSha256:customCandidates.sourceSha256,edits:[],additions:[{name:'Count',type:'integer',value:'3'}]}));customArgs.push('--custom-property-edits','v8-custom.json');}
   const output=`created-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','v8.json',...customArgs,'--json'];const dry=await runCliAsync([...args,'--dry-run'],f.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(f.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,f.options);assert.equal(saved.exitCode,0,saved.output);
   const next=await readFile(path.join(f.directory,output)),before=readGeneratedZip(source),after=readGeneratedZip(next);for(const [name,payload] of before)if(!['_rels/.rels','[Content_Types].xml','pptkit-core-properties-1.xml','pptkit-custom-properties-1.xml'].includes(name))assert.deepEqual(after.get(name),payload,name);
   const reopened=await importPptx(next,{projection:'content'});try{assert.equal(reopened.document.metadata.title,round===0?'PPTKit Presentation':`Title ${round}`);assert.equal(reopened.document.metadata.createdAt,replacement);assert.equal(reopened.document.metadata.customProperties[0].value,3);}finally{reopened.dispose();}source=next;input=output;
  }
 }
});

test('metadata v8 rejects stale, invalid and incompatible manifests without publication',async t=>{
 const f=await fixture(t,encodeMetadata(metadataFixture())),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-create','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates,title=c.targets.find(x=>x.field==='title'),date=c.targets.find(x=>x.field==='createdAt'),edit={...title,replacement:'New'},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,sourceSha256:'0'.repeat(64)},{...base,version:'pptkit-metadata-edits-v7'},{...base,edits:[{...edit,expected:'stale'}]},{...base,edits:[{...date,replacement:'2021-02-29T00:00:00.000Z'}]},{...base,edits:[edit,edit]}]){await writeFile(path.join(f.directory,'bad-v8.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-v8.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});}
 assert.equal((await runCliAsync(['inspect','missing','--metadata-create','--json'])).exitCode,2);
});


test('metadata v8 and custom creation fail together without publishing partial output',async t=>{
 const f=await fixture(t,encodeMetadata(metadataFixture())),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-create','--custom-property-values','--custom-property-additions','--custom-property-create','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);
 const c=decode(inspected),meta=c.metadataEditCandidates,custom=c.customPropertyEditCandidates,title=meta.targets.find(x=>x.field==='title');await writeFile(path.join(f.directory,'meta.json'),JSON.stringify({version:meta.version,sourceSha256:meta.sourceSha256,edits:[{...title,replacement:'New'}]}));await writeFile(path.join(f.directory,'custom.json'),JSON.stringify({version:custom.version,sourceSha256:custom.sourceSha256,edits:[],additions:[{name:'Invalid',type:'string',value:'_x005F_x0041_'}]}));
 const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','meta.json','--custom-property-edits','custom.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});
});


test('metadata v9 inserts default revision and saves further edits across three reopens',async t=>{
 for(const strict of [false,true])for(const existing of [false,true]){
  const data=metadataFixture(strict);if(existing)data.add('core','<d:title>Existing</d:title>');const f=await fixture(t,encodeMetadata(data));let input='source.pptx',source=f.bytes;
  for(let round=0;round<3;round++){
   const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-revision-insert','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates;assert.equal(c.version,'pptkit-metadata-edits-v9');const target=c.targets.find(x=>x.field==='revision');if(round===0)assert.equal(target.expected,null);
   const replacement=[1,7,0][round];await writeFile(path.join(f.directory,'v9.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement}]}));const output=`revision-insert-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','v9.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],f.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(f.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,f.options);assert.equal(saved.exitCode,0,saved.output);
   const next=await readFile(path.join(f.directory,output)),before=readGeneratedZip(source),after=readGeneratedZip(next);for(const [name,payload] of before)if(!['_rels/.rels','[Content_Types].xml','pptkit-core-properties-1.xml','inventory/core.data'].includes(name))assert.deepEqual(after.get(name),payload,name);const reopened=await importPptx(next,{projection:'content'});try{assert.equal(reopened.document.metadata.revision,replacement);}finally{reopened.dispose();}input=output;source=next;
  }
 }
});

test('metadata v9 rejects stale or malformed revision insertion without publishing',async t=>{
 const f=await fixture(t,encodeMetadata(metadataFixture())),report=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-revision-insert','--json'],f.options);assert.equal(report.exitCode,0,report.output);const c=decode(report).metadataEditCandidates,target=c.targets.find(x=>x.field==='revision'),edit={...target,replacement:1},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,sourceSha256:'0'.repeat(64)},{...base,version:'pptkit-metadata-edits-v8'},{...base,edits:[{...edit,replacement:-1}]},{...base,edits:[{...edit,replacement:1.5}]},{...base,edits:[{...edit,expected:1}]},{...base,edits:[edit,edit]}]){await writeFile(path.join(f.directory,'bad-v9.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-v9.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});}
 assert.equal((await runCliAsync(['inspect','missing','--metadata-revision-insert','--json'])).exitCode,2);
});


test('metadata v10 inserts one keyword across three reopens with or without core part',async t=>{
 for(const strict of [false,true])for(const present of [false,true]){
  const data=metadataFixture(strict);if(present)data.add('core','<d:title>Existing</d:title>');const f=await fixture(t,encodeMetadata(data));let input='source.pptx',source=f.bytes;
  for(const [round,value] of ['', '中文 & text','final'].entries()){
   const report=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-keywords-insert','--json'],f.options);assert.equal(report.exitCode,0,report.output);const c=decode(report).metadataEditCandidates,target=c.targets.find(x=>x.field==='keywords');assert.equal(c.version,'pptkit-metadata-edits-v10');if(round===0)assert.equal(target.expected,null);
   await writeFile(path.join(f.directory,'v10.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:value}]}));const output=`keywords-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','v10.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],f.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(f.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,f.options);assert.equal(saved.exitCode,0,saved.output);
   const next=await readFile(path.join(f.directory,output)),before=readGeneratedZip(source),after=readGeneratedZip(next);for(const [name,payload] of before)if(!['_rels/.rels','[Content_Types].xml','pptkit-core-properties-1.xml','inventory/core.data'].includes(name))assert.deepEqual(after.get(name),payload,name);const reopened=await importPptx(next,{projection:'content'});try{assert.deepEqual(reopened.document.metadata.keywords,[value]);}finally{reopened.dispose();}source=next;input=output;
  }
 }
});

test('metadata v10 rejects stale, invalid or duplicate keyword plans without publishing',async t=>{
 const f=await fixture(t,encodeMetadata(metadataFixture())),report=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-keywords-insert','--json'],f.options);assert.equal(report.exitCode,0,report.output);const c=decode(report).metadataEditCandidates,target=c.targets.find(x=>x.field==='keywords'),edit={...target,replacement:'new'},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,sourceSha256:'0'.repeat(64)},{...base,version:'pptkit-metadata-edits-v9'},{...base,edits:[{...edit,expected:''}]},{...base,edits:[{...edit,replacement:['new']}]},{...base,edits:[edit,edit]}]){await writeFile(path.join(f.directory,'bad-v10.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-v10.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});}
 assert.equal((await runCliAsync(['inspect','missing','--metadata-keywords-insert','--json'])).exitCode,2);
});


test('metadata v11 inserts company and preserves adjacent extended values over three rounds',async t=>{
 for(const strict of [false,true]){const data=metadataFixture(strict);data.add('extended','<Manager>Sentinel</Manager>');const f=await fixture(t,encodeMetadata(data));let input='source.pptx',source=f.bytes;
  for(const [round,value] of ['', '中文 & Co','Final'].entries()){
   const report=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-company-insert','--json'],f.options);assert.equal(report.exitCode,0,report.output);const c=decode(report).metadataEditCandidates,target=c.targets.find(x=>x.field==='company');assert.equal(c.version,'pptkit-metadata-edits-v11');if(round===0)assert.equal(target.expected,null);
   await writeFile(path.join(f.directory,'v11.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:value}]}));const output=`company-insert-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','v11.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],f.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(f.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,f.options);assert.equal(saved.exitCode,0,saved.output);const next=await readFile(path.join(f.directory,output)),before=readGeneratedZip(source),after=readGeneratedZip(next);for(const [name,payload] of before)if(name!=='inventory/extended.data')assert.deepEqual(after.get(name),payload,name);assert.match(after.get('inventory/extended.data').toString(),/<Manager>Sentinel<\/Manager>/);const reopened=await importPptx(next,{projection:'content'});try{assert.equal(reopened.document.metadata.company,value);}finally{reopened.dispose();}input=output;source=next;
  }
 }
});

test('metadata v11 refuses invalid or unsupported company insertions',async t=>{
 const data=metadataFixture();data.add('extended','<Manager>Sentinel</Manager>');const f=await fixture(t,encodeMetadata(data)),report=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-company-insert','--json'],f.options);assert.equal(report.exitCode,0,report.output);const c=decode(report).metadataEditCandidates,target=c.targets.find(x=>x.field==='company'),edit={...target,replacement:'New'},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,sourceSha256:'0'.repeat(64)},{...base,version:'pptkit-metadata-edits-v10'},{...base,edits:[{...edit,expected:''}]},{...base,edits:[{...edit,replacement:42}]},{...base,edits:[edit,edit]}]){await writeFile(path.join(f.directory,'bad-v11.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-v11.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});}
 const absent=await fixture(t,encodeMetadata(metadataFixture())),r=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-company-insert','--json'],absent.options);assert.equal(r.exitCode,0,r.output);assert.ok(!decode(r).metadataEditCandidates.targets.some(x=>x.field==='company'));assert.equal((await runCliAsync(['inspect','missing','--metadata-company-insert','--json'])).exitCode,2);
});

test('metadata v12 creates three property parts and survives three reopens', async t => {
 for (const strict of [false,true]) {
  const f=await fixture(t,encodeMetadata(metadataFixture(strict))); let input='source.pptx',source=f.bytes;
  for (let round=0;round<3;round++) {
   const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-extended-create','--custom-property-values','--custom-property-additions','--custom-property-create','--json'],f.options); assert.equal(inspected.exitCode,0,inspected.output);
   const report=decode(inspected),meta=report.metadataEditCandidates,custom=report.customPropertyEditCandidates; assert.equal(meta.version,'pptkit-metadata-edits-v12');
   const company=meta.targets.find(x=>x.field==='company'),title=meta.targets.find(x=>x.field==='title'); if(round===0){assert.equal(company.expected,null);assert.equal(title.expected,null);}
   await writeFile(path.join(f.directory,'v12.json'),JSON.stringify({version:meta.version,sourceSha256:meta.sourceSha256,edits:[{...company,replacement:['','中文 & Co','Final'][round]},{...title,replacement:`Title ${round}`}]}));
   const extra=[]; if(round===0){await writeFile(path.join(f.directory,'v12-custom.json'),JSON.stringify({version:custom.version,sourceSha256:custom.sourceSha256,edits:[],additions:[{name:'Count',type:'integer',value:'3'}]}));extra.push('--custom-property-edits','v12-custom.json');}
   const output=`v12-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','v12.json',...extra,'--json']; const dry=await runCliAsync([...args,'--dry-run'],f.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(f.directory,output)),{code:'ENOENT'});
   const result=await runCliAsync(args,f.options);assert.equal(result.exitCode,0,result.output);const next=await readFile(path.join(f.directory,output)),before=readGeneratedZip(source),after=readGeneratedZip(next);
   for(const [name,payload] of before)if(!['_rels/.rels','[Content_Types].xml','pptkit-core-properties-1.xml','pptkit-extended-properties-1.xml','pptkit-custom-properties-1.xml'].includes(name))assert.deepEqual(after.get(name),payload,name);
   for(const kind of ['core','extended','custom'])assert.ok(after.has(`pptkit-${kind}-properties-1.xml`)); const reopened=await importPptx(next,{projection:'content'});try{assert.equal(reopened.document.metadata.company,['','中文 & Co','Final'][round]);assert.equal(reopened.document.metadata.title,`Title ${round}`);assert.equal(reopened.document.metadata.customProperties[0].value,3);}finally{reopened.dispose();} input=output;source=next;
  }
 }
});

test('metadata v12 rejects stale, malformed and old manifests without publication', async t => {
 const f=await fixture(t,encodeMetadata(metadataFixture())),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-extended-create','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);
 const meta=decode(inspected).metadataEditCandidates,company=meta.targets.find(x=>x.field==='company'),edit={...company,replacement:'Company'},base={version:meta.version,sourceSha256:meta.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,sourceSha256:'0'.repeat(64)},{...base,version:'pptkit-metadata-edits-v11'},{...base,edits:[{...edit,expected:'stale'}]},{...base,edits:[{...edit,replacement:3}]},{...base,edits:[{...edit,replacement:'bad\u0000'}]},{...base,edits:[edit,edit]}]){await writeFile(path.join(f.directory,'bad-v12.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-v12.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});}
 assert.equal((await runCliAsync(['inspect','missing','--metadata-extended-create','--json'])).exitCode,2);
});

test('metadata v12 and invalid custom creation never publish a partial transaction', async t => {
 const f=await fixture(t,encodeMetadata(metadataFixture())),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-extended-create','--custom-property-values','--custom-property-additions','--custom-property-create','--json'],f.options);assert.equal(inspected.exitCode,0,inspected.output);
 const report=decode(inspected),meta=report.metadataEditCandidates,custom=report.customPropertyEditCandidates;
 await writeFile(path.join(f.directory,'meta-v12.json'),JSON.stringify({version:meta.version,sourceSha256:meta.sourceSha256,edits:[{...meta.targets.find(x=>x.field==='company'),replacement:'Company'},{...meta.targets.find(x=>x.field==='title'),replacement:'Title'}]}));
 await writeFile(path.join(f.directory,'custom-v12.json'),JSON.stringify({version:custom.version,sourceSha256:custom.sourceSha256,edits:[],additions:[{name:'Invalid',type:'string',value:'_x005F_x0041_'}]}));
 const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','meta-v12.json','--custom-property-edits','custom-v12.json','--json'],f.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(f.directory,'bad.pptx')),{code:'ENOENT'});
});

test('metadata v13 removes, reinserts and removes optional fields across three reopens',async t=>{
 for(const strict of [false,true]){const f=metadataFixture(strict);f.add('core','<d:subject>S</d:subject><d:description>D</d:description><c:revision>7</c:revision>');f.add('extended','<Manager>M</Manager><Company>C</Company>');const fixtureData=await fixture(t,encodeMetadata(f));let input='source.pptx';
  for(let round=0;round<3;round++){
   const report=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-remove','--json'],fixtureData.options);assert.equal(report.exitCode,0,report.output);const candidates=decode(report).metadataEditCandidates;assert.equal(candidates.version,'pptkit-metadata-edits-v13');
   if(round!==1)assert.deepEqual([...candidates.removableFields].sort(),['company','description','subject']);else assert.deepEqual(candidates.removableFields,[]);
   const edits=['subject','description','company'].map(field=>({...candidates.targets.find(item=>item.field===field),replacement:round===1?`New ${field}`:null}));await writeFile(path.join(fixtureData.directory,'remove.json'),JSON.stringify({version:candidates.version,sourceSha256:candidates.sourceSha256,edits}));
   const output=`removed-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','remove.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],fixtureData.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(fixtureData.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,fixtureData.options);assert.equal(saved.exitCode,0,saved.output);
   const bytes=await readFile(path.join(fixtureData.directory,output)),zip=readGeneratedZip(bytes);assert.match(zip.get('inventory/extended.data').toString(),/<Manager>M<\/Manager>/);const reopened=await importPptx(bytes,{projection:'content'});try{for(const field of ['subject','description','company'])assert.equal(reopened.document.metadata[field],round===1?`New ${field}`:undefined);assert.equal(reopened.document.metadata.revision,7);}finally{reopened.dispose();}input=output;
  }
 }
});

test('metadata v13 rejects unsafe removals, stale values and old manifests without publication',async t=>{
 const f=metadataFixture();f.add('core','<d:subject>S</d:subject><d:title>Title</d:title>');f.add('extended','<Company>C</Company>');const fixtureData=await fixture(t,encodeMetadata(f)),r=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-remove','--json'],fixtureData.options);assert.equal(r.exitCode,0,r.output);const c=decode(r).metadataEditCandidates,subject=c.targets.find(x=>x.field==='subject'),title=c.targets.find(x=>x.field==='title'),company=c.targets.find(x=>x.field==='company'),base={version:c.version,sourceSha256:c.sourceSha256,edits:[{...subject,replacement:null}]};
 for(const manifest of [{...base,version:'pptkit-metadata-edits-v12'},{...base,sourceSha256:'0'.repeat(64)},{...base,edits:[{...subject,expected:'wrong',replacement:null}]},{...base,edits:[{...title,replacement:null}]},{...base,edits:[{...company,expected:null,replacement:null}]},{...base,edits:[{...subject,replacement:null},{...subject,replacement:null}]},{...base,edits:[{...subject,replacement:null},{...company,expected:'wrong',replacement:'change'}]}]){await writeFile(path.join(fixtureData.directory,'bad-remove.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-remove.json','--json'],fixtureData.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(fixtureData.directory,'bad.pptx')),{code:'ENOENT'});}
 assert.equal((await runCliAsync(['inspect','missing','--metadata-remove','--json'])).exitCode,2);
});

test('metadata v14 removes and reinserts dates across three reopens in both dialects',async t=>{
 for(const strict of [false,true]){const f=metadataFixture(strict);f.add('core','<t:created xsi:type="t:W3CDTF">2020-02-29T10:00:00+08:00</t:created><t:modified xsi:type="t:W3CDTF">2021-01-02T03:04:05Z</t:modified><d:subject>S</d:subject>');const data=await fixture(t,encodeMetadata(f));let input='source.pptx';
  for(let round=0;round<3;round++){const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-date-remove','--json'],data.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates;assert.equal(c.version,'pptkit-metadata-edits-v14');if(round!==1)for(const field of ['createdAt','modifiedAt'])assert.ok(c.removableFields.includes(field));const edits=['createdAt','modifiedAt'].map(field=>({...c.targets.find(x=>x.field===field),replacement:round===1?'2022-01-01T00:00:00.123Z':null}));await writeFile(path.join(data.directory,'date-remove.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits}));const output=`date-removal-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','date-remove.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],data.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(data.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,data.options);assert.equal(saved.exitCode,0,saved.output);const r=await importPptx(await readFile(path.join(data.directory,output)),{projection:'content'});try{for(const field of ['createdAt','modifiedAt'])assert.equal(r.document.metadata[field],round===1?'2022-01-01T00:00:00.123Z':undefined);assert.equal(r.document.metadata.subject,'S');}finally{r.dispose();}input=output;}
 }
});

test('metadata v14 rejects stale and incompatible date removals without publication',async t=>{
 const f=metadataFixture();f.add('core','<t:created xsi:type="t:W3CDTF">2020-01-01T00:00:00Z</t:created><d:subject>S</d:subject>');const data=await fixture(t,encodeMetadata(f)),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-date-remove','--json'],data.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates,date=c.targets.find(x=>x.field==='createdAt'),subject=c.targets.find(x=>x.field==='subject'),edit={...date,replacement:null},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,version:'pptkit-metadata-edits-v13'},{...base,sourceSha256:'0'.repeat(64)},{...base,edits:[{...edit,expected:'2021-01-01T00:00:00.000Z'}]},{...base,edits:[{...edit,expected:null}]},{...base,edits:[edit,edit]},{...base,edits:[{...subject,replacement:null},{...edit,expected:'2021-01-01T00:00:00.000Z'}]}]){await writeFile(path.join(data.directory,'bad-date-remove.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-date-remove.json','--json'],data.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(data.directory,'bad.pptx')),{code:'ENOENT'});}assert.equal((await runCliAsync(['inspect','missing','--metadata-date-remove','--json'])).exitCode,2);
});

test('metadata v15 removes and reinserts a keyword over three reopens',async t=>{
 for(const strict of [false,true])for(const initial of ['', 'one, two']){const f=metadataFixture(strict);f.add('core',`<c:keywords>${initial}</c:keywords><d:subject>S</d:subject>`);const data=await fixture(t,encodeMetadata(f));let input='source.pptx';
  for(let round=0;round<3;round++){const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-keywords-remove','--json'],data.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates;assert.equal(c.version,'pptkit-metadata-edits-v15');if(round!==1)assert.ok(c.removableFields.includes('keywords'));const target=c.targets.find(x=>x.field==='keywords');await writeFile(path.join(data.directory,'keyword-remove.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:round===1?'中文 & key':null}]}));const output=`keyword-remove-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','keyword-remove.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],data.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(data.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,data.options);assert.equal(saved.exitCode,0,saved.output);const r=await importPptx(await readFile(path.join(data.directory,output)),{projection:'content'});try{assert.deepEqual(r.document.metadata.keywords??[],round===1?['中文 & key']:[]);assert.equal(r.document.metadata.subject,'S');}finally{r.dispose();}input=output;}
 }
});

test('metadata v15 rejects stale, unsupported and old keyword deletion manifests atomically',async t=>{
 const f=metadataFixture();f.add('core','<c:keywords>one</c:keywords><d:subject>S</d:subject>');const data=await fixture(t,encodeMetadata(f)),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-keywords-remove','--json'],data.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates,keyword=c.targets.find(x=>x.field==='keywords'),subject=c.targets.find(x=>x.field==='subject'),edit={...keyword,replacement:null},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,version:'pptkit-metadata-edits-v14'},{...base,sourceSha256:'0'.repeat(64)},{...base,edits:[{...edit,expected:'wrong'}]},{...base,edits:[{...edit,expected:null}]},{...base,edits:[edit,edit]},{...base,edits:[{...subject,replacement:null},{...edit,expected:'wrong'}]}]){await writeFile(path.join(data.directory,'bad-keyword-remove.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-keyword-remove.json','--json'],data.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(data.directory,'bad.pptx')),{code:'ENOENT'});}assert.equal((await runCliAsync(['inspect','missing','--metadata-keywords-remove','--json'])).exitCode,2);
});

test('metadata v16 removes and reinserts revisions including the default over three reopens',async t=>{
 for(const strict of [false,true])for(const initial of [1,7]){const f=metadataFixture(strict);f.add('core',`<c:revision>${initial}</c:revision><d:subject>S</d:subject>`);const data=await fixture(t,encodeMetadata(f));let input='source.pptx';
  for(let round=0;round<3;round++){const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-revision-remove','--json'],data.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates;assert.equal(c.version,'pptkit-metadata-edits-v16');if(round!==1)assert.ok(c.removableFields.includes('revision'));const target=c.targets.find(x=>x.field==='revision');if(round===1)assert.equal(target.expected,null);await writeFile(path.join(data.directory,'revision-remove.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits:[{...target,replacement:round===1?1:null}]}));const output=`revision-remove-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','revision-remove.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],data.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(data.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,data.options);assert.equal(saved.exitCode,0,saved.output);const r=await importPptx(await readFile(path.join(data.directory,output)),{projection:'content'});try{assert.equal(r.document.metadata.revision,round===1?1:undefined);assert.equal(r.document.metadata.subject,'S');}finally{r.dispose();}input=output;}
 }
});

test('metadata v16 refuses stale, unbound and older revision removal plans atomically',async t=>{
 const f=metadataFixture();f.add('core','<c:revision>7</c:revision><d:subject>S</d:subject>');const data=await fixture(t,encodeMetadata(f)),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-revision-remove','--json'],data.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates,revision=c.targets.find(x=>x.field==='revision'),subject=c.targets.find(x=>x.field==='subject'),edit={...revision,replacement:null},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,version:'pptkit-metadata-edits-v15'},{...base,sourceSha256:'0'.repeat(64)},{...base,edits:[{...edit,expected:8}]},{...base,edits:[{...edit,expected:null}]},{...base,edits:[edit,edit]},{...base,edits:[{...subject,replacement:null},{...edit,expected:8}]}]){await writeFile(path.join(data.directory,'bad-revision-remove.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-revision-remove.json','--json'],data.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(data.directory,'bad.pptx')),{code:'ENOENT'});}assert.equal((await runCliAsync(['inspect','missing','--metadata-revision-remove','--json'])).exitCode,2);
});

test('metadata v17 removes and reinserts defaulted fields through three reopens',async t=>{
 for(const strict of [false,true])for(const initial of [{title:'PPTKit Presentation',author:'PPTKit',language:'en-US',revision:1},{title:'Title',author:'Author',language:'fr-FR',revision:7}]){const f=metadataFixture(strict);f.add('core',`<d:title>${initial.title}</d:title><d:creator>${initial.author}</d:creator><d:language>${initial.language}</d:language><c:revision>${initial.revision}</c:revision><d:subject>S</d:subject>`);const data=await fixture(t,encodeMetadata(f));let input='source.pptx';const fields=['title','author','language','revision'],defaults={title:'PPTKit Presentation',author:'PPTKit',language:'en-US',revision:1};
  for(let round=0;round<3;round++){const inspected=await runCliAsync(['inspect',input,'--projection','content','--metadata-values','--metadata-default-remove','--json'],data.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates;assert.equal(c.version,'pptkit-metadata-edits-v17');if(round!==1)for(const field of fields)assert.ok(c.removableFields.includes(field));const edits=fields.map(field=>({...c.targets.find(x=>x.field===field),replacement:round===1?defaults[field]:null}));await writeFile(path.join(data.directory,'default-remove.json'),JSON.stringify({version:c.version,sourceSha256:c.sourceSha256,edits}));const output=`default-remove-${round}.pptx`,args=['save',input,output,'--mode','edit','--metadata-edits','default-remove.json','--json'];const dry=await runCliAsync([...args,'--dry-run'],data.options);assert.equal(dry.exitCode,0,dry.output);await assert.rejects(readFile(path.join(data.directory,output)),{code:'ENOENT'});const saved=await runCliAsync(args,data.options);assert.equal(saved.exitCode,0,saved.output);const r=await importPptx(await readFile(path.join(data.directory,output)),{projection:'content'});try{for(const field of fields)assert.equal(r.document.metadata[field],round===1?defaults[field]:undefined);assert.equal(r.document.metadata.subject,'S');}finally{r.dispose();}input=output;}
 }
});

test('metadata v17 rejects stale, duplicate and old deletion manifests without output',async t=>{
 const f=metadataFixture();f.add('core','<d:title>Title</d:title><d:creator>Author</d:creator>');const data=await fixture(t,encodeMetadata(f)),inspected=await runCliAsync(['inspect','source.pptx','--projection','content','--metadata-values','--metadata-default-remove','--json'],data.options);assert.equal(inspected.exitCode,0,inspected.output);const c=decode(inspected).metadataEditCandidates,title=c.targets.find(x=>x.field==='title'),author=c.targets.find(x=>x.field==='author'),edit={...title,replacement:null},base={version:c.version,sourceSha256:c.sourceSha256,edits:[edit]};
 for(const manifest of [{...base,version:'pptkit-metadata-edits-v16'},{...base,sourceSha256:'0'.repeat(64)},{...base,edits:[{...edit,expected:'wrong'}]},{...base,edits:[{...edit,expected:null}]},{...base,edits:[edit,edit]},{...base,edits:[{...author,replacement:null},{...edit,expected:'wrong'}]}]){await writeFile(path.join(data.directory,'bad-default-remove.json'),JSON.stringify(manifest));const result=await runCliAsync(['save','source.pptx','bad.pptx','--mode','edit','--metadata-edits','bad-default-remove.json','--json'],data.options);assert.equal(result.exitCode,3,result.output);await assert.rejects(readFile(path.join(data.directory,'bad.pptx')),{code:'ENOENT'});}assert.equal((await runCliAsync(['inspect','missing','--metadata-default-remove','--json'])).exitCode,2);
});

test('CLI image data validation omits bad media and can prevent publication', async t => {
  const f = await fixture(t);
  const original = await readFile(new URL('../../../tests/fixtures/exif-orientation6.jpg', import.meta.url));
  const bad = [
    ['png', Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGMVBAAAHQAXjsBI9wAAAABJRU5ErkJggg==', 'base64')],
    ['gif', Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64')],
    ['jpeg', original.subarray(0, -2)],
  ];
  await writeFile(path.join(f.directory, 'control.jpg'), original);
  const assets = [
    { id: 'control', kind: 'image', source: { type: 'path', value: 'control.jpg' } },
    ...bad.map(([id, bytes]) => ({ id, kind: 'image', source: { type: 'bytes', encoding: 'base64', value: bytes.toString('base64') } })),
  ];
  const spec = { version: 'pptkit-authoring-v1', assets, slides: [{ elements: assets.map(({ id }, i) => ({
    type: 'image', assetId: id, box: { x: i * 40, y: 10, width: 30, height: 30 },
  })) }] };
  await writeFile(path.join(f.directory, 'image-data.json'), JSON.stringify(spec));
  const common = ['--asset-base', '.', '--json'], check = '--validate-image-data';
  const baseline = await runCliAsync(['validate', 'image-data.json', '--export', ...common], f.options);
  assert.equal(baseline.exitCode, 0, baseline.output);
  assert.deepEqual(decode(baseline).report.export.warnings, []);

  const validated = await runCliAsync(['validate', 'image-data.json', '--export', check, ...common], f.options);
  assert.equal(validated.exitCode, 0, validated.output);
  const warnings = decode(validated).report.export.warnings;
  assert.deepEqual(warnings.filter(w => w.code === 'asset-read-failed').map(w => w.assetId), ['png', 'gif', 'jpeg']);
  for (const [id, reason] of [['png', /row filter/], ['gif', /missing end code/], ['jpeg', /missing EOI/]]) {
    assert.match(warnings.find(w => w.assetId === id && w.code === 'asset-read-failed').message, reason);
    assert.ok(warnings.some(w => w.code === 'image-omitted' && w.assetId === id));
  }
  const combined = await runCliAsync(['validate', 'image-data.json', '--export', check, '--image-diagnostics', ...common], f.options);
  assert.equal(combined.exitCode, 0, combined.output);
  assert.deepEqual(decode(combined).report.export.warnings.filter(w => w.assetId === 'control').map(w => w.code), ['jpeg-exif-orientation-unverified']);

  const validatedStrict = await runCliAsync(['validate', 'image-data.json', '--export', check, '--fail-on-export-warning', ...common], f.options);
  assert.equal(validatedStrict.exitCode, 1, validatedStrict.output);
  const rejected = await runCliAsync(['generate', 'image-data.json', 'rejected-data.pptx', check, '--fail-on-export-warning', ...common], f.options);
  assert.equal(rejected.exitCode, 3, rejected.output);
  assert.equal(decode(rejected).error.code, 'export-warnings-rejected');
  await assert.rejects(readFile(path.join(f.directory, 'rejected-data.pptx')), { code: 'ENOENT' });

  const generated = await runCliAsync(['generate', 'image-data.json', 'filtered-data.pptx', check, ...common], f.options);
  assert.equal(generated.exitCode, 0, generated.output);
  assert.deepEqual(decode(generated).report.exportWarnings, warnings);
  const media = [...readGeneratedZip(await readFile(path.join(f.directory, 'filtered-data.pptx')))].filter(([part]) => part.startsWith('ppt/media/'));
  assert.deepEqual(media, [['ppt/media/control.jpg', original]]);
  for (const args of [['validate', 'image-data.json', check], ['inspect', 'x', check], ['generate', 'image-data.json', 'x', check, check]]) {
    assert.equal((await runCliAsync(args, f.options)).exitCode, 2);
  }
});

test('chart manifest v4 discovers and saves independent bubble coordinates', async t => {
  const { createPresentation } = await import('@pptkit/core');
  const { generatePptx } = await import('@pptkit/pptx-exporter');
  const deck = createPresentation();
  deck.addSlide({ elements: [{ type: 'chart', chartType: 'bubble', box: { x: 20, y: 20, width: 400, height: 300 },
    series: [{ name: 'Series', xValues: [1, 2], yValues: [3, 4], sizes: [5, 6] }] }] });
  const source = (await generatePptx(deck)).bytes, f = await fixture(t, source);
  const inspect = (file, version) => runCliAsync(['inspect', file, '--projection', 'content', '--chart-edit-save-profile', version, '--chart-values', '--json'], f.options);
  const old = await inspect('source.pptx', 'existing-values-v2');
  assert.equal(old.exitCode, 0, old.output);
  assert.deepEqual(decode(old).chartEditCandidatesV3.targets, []);
  const current = await inspect('source.pptx', 'existing-values-v3');
  assert.equal(current.exitCode, 0, current.output);
  const report = decode(current), candidates = report.chartEditCandidatesV4;
  assert.equal(report.chartEditCandidatesV3, undefined);
  assert.equal(candidates.version, 'pptkit-chart-edits-v4');
  assert.equal(candidates.sourceSha256, sha(source));
  assert.equal(candidates.targets.length, 6);
  const selected = candidates.targets.filter(target => target.pointIndex === 1);
  assert.deepEqual(selected.map(({ coordinate, expected }) => [coordinate, expected]), [['x', '2'], ['y', '4'], ['size', '6']]);
  const edits = selected.map(target => ({ ...target, replacement: { x: '-0', y: '7.5', size: '16' }[target.coordinate] }));
  const manifest = { version: candidates.version, sourceSha256: candidates.sourceSha256, edits };
  const save = async (plan, file) => {
    await writeFile(path.join(f.directory, 'bubble-edits.json'), JSON.stringify(plan));
    return runCliAsync(['save', 'source.pptx', file, '--mode', 'edit', '--chart-edits', 'bubble-edits.json', '--json'], f.options);
  };
  for (const [plan, code] of [
    [{ ...manifest, version: 'pptkit-chart-edits-v3' }, 'chart-edits-target-unbound'],
    [{ ...manifest, sourceSha256: '0'.repeat(64) }, 'chart-edits-source-mismatch'],
    [{ ...manifest, edits: edits.map(edit => edit.coordinate === 'size' ? { ...edit, replacement: '0' } : edit) }, 'chart-edits-invalid'],
  ]) {
    const rejected = await save(plan, 'rejected.pptx');
    assert.equal(rejected.exitCode, 3, rejected.output);
    assert.equal(decode(rejected).error.code, code);
    await assert.rejects(readFile(path.join(f.directory, 'rejected.pptx')), { code: 'ENOENT' });
  }
  const saved = await save(manifest, 'bubble.pptx');
  assert.equal(saved.exitCode, 0, saved.output);
  assert.equal(decode(saved).report.assurance, 'scoped-preservation');
  const next = await inspect('bubble.pptx', 'existing-values-v3');
  assert.equal(next.exitCode, 0, next.output);
  assert.deepEqual(decode(next).chartEditCandidatesV4.targets.filter(target => target.pointIndex === 1).map(target => target.expected), ['-0', '7.5', '16']);
  const bytes = await readFile(path.join(f.directory, 'bubble.pptx')), parts = readGeneratedZip(bytes);
  const worksheet = readGeneratedZip(parts.get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')).get('xl/worksheets/sheet1.xml').toString();
  assert.match(worksheet, /<c r="A3"><v>-0<\/v><\/c><c r="B3"><v>7\.5<\/v><\/c><c r="C3"><v>16<\/v><\/c>/);
  const native = await readFile(new URL('../../../tests/fixtures/bubble-libreoffice-resave.pptx', import.meta.url));
  await writeFile(path.join(f.directory, 'native.pptx'), native);
  const readOnly = await inspect('native.pptx', 'existing-values-v3');
  assert.equal(readOnly.exitCode, 0, readOnly.output);
  assert.deepEqual(decode(readOnly).chartEditCandidatesV4.targets, []);
  assert.equal(decode(readOnly).chartEditCandidatesV4.unboundCount, 9);
});

test('chart manifest v4 saves independent column and line values in the same combination chart', async t => {
  const { createPresentation } = await import('@pptkit/core');
  const { generatePptx } = await import('@pptkit/pptx-exporter');
  const deck = createPresentation();
  deck.addSlide({ elements: [{ type: 'chart', chartType: 'combo', box: { x: 20, y: 20, width: 400, height: 300 },
    categories: ['A', 'B'], series: [{ kind: 'bar', name: 'Units', values: [2, 4] }, { kind: 'line', name: 'Rate', values: [80, 60] }] }] });
  const f = await fixture(t, (await generatePptx(deck)).bytes);
  const inspect = (file, profile) => runCliAsync(['inspect', file, '--projection', 'content', '--chart-edit-save-profile', profile, '--chart-values', '--json'], f.options);
  const current = await inspect('source.pptx', 'existing-values-v3');
  assert.equal(current.exitCode, 0, current.output);
  const { version, sourceSha256, targets } = decode(current).chartEditCandidatesV4;
  assert.equal(targets.length, 4);
  const edits = targets.filter(target => target.pointIndex === 1).map(target => ({ ...target, replacement: ['3.5', '75'][target.seriesIndex] }));
  assert.deepEqual(edits.map(target => [target.seriesIndex, target.expected, target.coordinate]), [[0, '4', undefined], [1, '60', undefined]]);
  const save = async (v, output) => {
    await writeFile(path.join(f.directory, 'combo-edits.json'), JSON.stringify({ version: v, sourceSha256, edits }));
    return runCliAsync(['save', 'source.pptx', output, '--mode', 'edit', '--chart-edits', 'combo-edits.json', '--json'], f.options);
  };
  const rejected = await save('pptkit-chart-edits-v3', 'blocked.pptx');
  assert.equal(rejected.exitCode, 3, rejected.output);
  assert.equal(decode(rejected).error.code, 'chart-edits-target-unbound');
  await assert.rejects(readFile(path.join(f.directory, 'blocked.pptx')), { code: 'ENOENT' });
  const saved = await save(version, 'edited.pptx');
  assert.equal(saved.exitCode, 0, saved.output);
  assert.equal(decode(saved).report.assurance, 'scoped-preservation');
  const reopened = await inspect('edited.pptx', 'existing-values-v3');
  assert.equal(reopened.exitCode, 0, reopened.output);
  assert.deepEqual(decode(reopened).chartEditCandidatesV4.targets.filter(target => target.pointIndex === 1).map(target => target.expected), ['3.5', '75']);
});

test('authoring JSON reading order survives SVG preview and warns on PPTX omission', async t => {
  const f = await fixture(t), box = { x: 10, y: 10, width: 50, height: 30 };
  const spec = { version: 'pptkit-authoring-v1', slides: [{
    id: 'ordered', elements: ['a', 'b'].map(id => ({ id, type: 'shape', shape: 'rect', box })),
    readingOrder: [{ elementIds: ['b', 'a'] }],
  }] };
  await writeFile(path.join(f.directory, 'ordered.json'), JSON.stringify(spec));
  const preview = await runCliAsync(['preview', 'ordered.json', '--slide', '1', '--json'], f.options);
  assert.equal(preview.exitCode, 0, preview.output);
  assert.match(decode(preview).report.slide.svg, /aria-owns=/);
  const generated = await runCliAsync(['generate', 'ordered.json', 'ordered.pptx', '--json'], f.options);
  assert.equal(generated.exitCode, 0, generated.output);
  assert.ok(decode(generated).report.exportWarnings.some(w => w.code === 'reading-order-not-exported' && w.slideId === 'ordered'));
  spec.slides[0].readingOrder[0].elementIds = ['b', 'b'];
  await writeFile(path.join(f.directory, 'ordered.json'), JSON.stringify(spec));
  const invalid = await runCliAsync(['generate', 'ordered.json', 'invalid.pptx', '--json'], f.options);
  assert.equal(invalid.exitCode, 3, invalid.output);
  assert.equal(decode(invalid).error.code, 'authoring-invalid');
  await assert.rejects(readFile(path.join(f.directory, 'invalid.pptx')), { code: 'ENOENT' });
});

test('contrast report distinguishes shared text occurrences and enforces an output bound', async t => {
  const f = await fixture(t), box = { x: 0, y: 0, width: 100, height: 50 };
  const spec = { version: 'pptkit-authoring-v1', presentation: { masters: [{
    id: 'm', name: 'Master', elements: [{ id: 'shared-text', type: 'text', box, content: [{ runs: [{ text: 'Label', style: { color: '000000' } }] }] }],
  }] }, layouts: [{ id: 'l', name: 'Layout', masterId: 'm' }], slides: [
    { id: 'white', layoutId: 'l', background: { type: 'solid', color: 'FFFFFF' } },
    { id: 'black', layoutId: 'l', background: { type: 'solid', color: '000000' } },
  ] };
  await writeFile(path.join(f.directory, 'shared-contrast.json'), JSON.stringify(spec));
  const args = ['validate', 'shared-contrast.json', '--check-contrast'];
  const result = await runCliAsync([...args, '--fail-on-contrast', '--json'], f.options);
  assert.equal(result.exitCode, 1, result.output);
  assert.deepEqual(decode(result).report.contrast.findings.map(finding => [finding.slideId, finding.source, finding.status]),
    [['white', { kind: 'master', id: 'm' }, 'meets-threshold'], ['black', { kind: 'master', id: 'm' }, 'below-threshold']]);
  const human = await runCliAsync(args, f.options);
  assert.match(human.output, /slide=black source=master:m/);
  spec.slides = Array.from({ length: 64 }, (_, i) => ({ ...spec.slides[0], id: `page-${i}` }));
  await writeFile(path.join(f.directory, 'shared-contrast.json'), JSON.stringify(spec));
  const bounded = await runCliAsync([...args, '--max-output-bytes', '1024', '--json'], f.options);
  assert.equal(bounded.exitCode, 4, bounded.output);
  assert.equal(decode(bounded).error.code, 'output-limit-exceeded');
  assert.match(decode(bounded).error.message, /Contrast findings/);
});
