import assert from "node:assert/strict";
import test from "node:test";
import { mkdtemp, writeFile, readFile, rm, readdir, stat, realpath } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, compareOpcPackages } from "../packages/pptx-parser/dist/index.js";
import { importPptxFile, writePptxFile } from "../packages/pptx-parser/dist/node.js";
import { createTableSaveFixture } from "../tests/fixtures/table-save.mjs";
import { png } from "../packages/pptx-parser/test/image-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { diagramDataFixture } from "../packages/pptx-parser/test/diagram-data-fixture.mjs";
import { shape, body } from "../packages/pptx-parser/test/content-fixture.mjs";

const profile = { projection: "content", editSaveProfile: "text-runs-v3" };
const hash = bytes => createHash("sha256").update(bytes).digest("hex");

test("file imports capture custom graph budgets before I/O and retain them for candidate reprojection", async t => {
  const directory = await realpath(await mkdtemp(path.join(tmpdir(), "pptkit-file-diagram-budget-")));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, "source.pptx"), f = diagramDataFixture();
  // More findings than the default: losing the custom budget during file capture
  // or actual-output reprojection must fail, rather than silently pass this test.
  f.setModel("<d:pt/>".repeat(10001), "");
  f.setTree(f.frame() + shape({ id: "10", txBox: true, text: body() }));
  const bytes = f.bytes(); await writeFile(file, bytes);
  const options = { ...profile, diagramLimits: { maxFindings: 10002 } };
  const pending = importPptxFile(file, options); options.diagramLimits.maxFindings = 1;
  const imported = await pending;
  try {
    assert.equal(imported.session.report.inspection.diagramInventory.dataModels[0].findings.length, 10001);
    const slide = imported.session.document.slides[0], text = structuredClone(slide.elements[0]);
    text.content[0].runs[0].text = "Custom budget survives candidate validation";
    slide.removeElement(text.id); slide.insertElement(0, text);
    const prepared = await imported.session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
    const saved = await imported.session.save(prepared.plan);
    assert.deepEqual(readGeneratedZip(saved.bytes).get("deck/diagrams/data.xml"), readGeneratedZip(bytes).get("deck/diagrams/data.xml"));
    assert.deepEqual(await readFile(file), bytes);
  } finally { imported.session.dispose(); }
  await assert.rejects(importPptxFile(file, { ...profile, diagramLimits: { maxFindings: 1 } }), { code: "import-limit-exceeded" });
});

test("three real in-place file saves preserve mixed-deck resources and equal byte-session candidates without rebasing", async t => {
  const directory = await realpath(await mkdtemp(path.join(tmpdir(), "pptkit-file-parity-")));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, "source.pptx");
  const original = (await generatePptx(createTableSaveFixture(createPresentation, png))).bytes;
  await writeFile(file, original, { mode: 0o640 });
  const sourceParts = readGeneratedZip(original); let previousSize;
  for (let round = 1; round <= 3; round++) {
    const source = await readFile(file), before = await stat(file), imported = await importPptxFile(file, profile);
    try {
      const slide = imported.session.document.slides[0], originalTable = slide.elements.find(e => e.type === "table" && e.columns.length === 3), index = slide.elements.indexOf(originalTable), table = structuredClone(originalTable);
      table.rows[1].cells[0].content[0].runs[0].text = `Saved round ${round}`;
      slide.removeElement(table.id); slide.insertElement(index, table);
      const prepared = await imported.session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
      const expected = await imported.session.save(prepared.plan);
      const written = await writePptxFile(imported.session, prepared.plan, file, { mode: "replace", expected: imported.file });
      const bytes = await readFile(file), output = readGeneratedZip(bytes), after = await stat(file);
      assert.equal(written.status, "committed"); assert.equal(written.directorySynced, true); assert.deepEqual(written.warnings, []);
      assert.equal(hash(bytes), hash(expected.bytes)); assert.equal(hash(bytes), written.file.sha256);
      assert.notEqual(after.ino, before.ino); assert.equal(after.mode & 0o777, 0o640);
      assert.deepEqual((await compareOpcPackages(source, bytes)).differences.map(d => [d.code, d.part]), [["payload-changed", "/ppt/slides/slide1.xml"]]);
      for (const [part, payload] of sourceParts) if (part !== "ppt/slides/slide1.xml") assert.equal(hash(output.get(part)), hash(payload), part);
      const reopened = await importPptx(bytes, profile);
      assert.deepEqual(normalizePresentation(reopened.document), normalizePresentation(imported.session.document)); reopened.dispose();
      assert.equal((await imported.session.prepareSave({ mode: "no-op" })).status, "blocked");
      assert.equal(imported.session.assertSavePlanCurrent(prepared.plan), undefined);
      if (previousSize) assert.equal(bytes.length, previousSize);
      previousSize = bytes.length;
      assert.deepEqual(await readdir(directory), ["source.pptx"]);
    } finally { imported.session.dispose(); }
  }
});

test("file output errors retain usable source sessions/plans and never alter the original byte baseline", async t => {
  const directory = await realpath(await mkdtemp(path.join(tmpdir(), "pptkit-file-retry-")));
  t.after(() => rm(directory, { recursive: true, force: true }));
  const source = path.join(directory, "source.pptx"), copy = path.join(directory, "copy.pptx");
  const input = (await generatePptx(createTableSaveFixture(createPresentation))).bytes; await writeFile(source, input);
  const imported = await importPptxFile(source, profile);
  try {
    const plan = await imported.session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
    await assert.rejects(writePptxFile(imported.session, plan.plan, source, { mode: "create" }), { code: "file-target-exists" });
    const saved = await writePptxFile(imported.session, plan.plan, copy, { mode: "create" });
    assert.equal(hash(await readFile(source)), hash(input)); assert.equal(hash(await readFile(copy)), hash(input));
    assert.equal(saved.report.assurance, "original-bytes");
    const replaced = await writePptxFile(imported.session, plan.plan, copy, { mode: "replace", expected: saved.file });
    assert.equal(replaced.file.sha256, saved.file.sha256); assert.equal(hash((await imported.session.save(plan.plan)).bytes), hash(input));
    assert.deepEqual((await readdir(directory)).sort(), ["copy.pptx", "source.pptx"]);
  } finally { imported.session.dispose(); }
});
