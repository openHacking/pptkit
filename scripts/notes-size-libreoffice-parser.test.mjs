import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";
import test from "node:test";
import { chromium } from "@playwright/test";
import { importPptx, inspectPptxNotes } from "../packages/pptx-parser/dist/index.js";

const moduleUrl = text => "data:text/javascript;base64," + Buffer.from(text).toString("base64");
const frame = (report, partName) => {
  const part = report.parts.find(p => p.part === `/${partName}`), drawing = part.drawings.find(d => d.role === "slide-image" && d.view === "effective");
  assert.ok(drawing); const transform = part.geometries.find(g => g.kind === "transform" && g.ownerDrawingElementIndex === drawing.source.elementIndex && g.view === "effective");
  if (!transform) return null;
  const attrs = name => Object.fromEntries(transform.components.find(c => c.name.local === name).attributes.map(a => [a.local, Number(a.value)]));
  const off = attrs("off"), ext = attrs("ext"); return { x: off.x, y: off.y, width: ext.cx, height: ext.cy };
};
const notesText = slide => slide.notes.flatMap(p => p.runs.map(r => r.text)).join("");

test("actual LibreOffice boundary files retain raw notes geometry and no-op bytes in Node, ESM and global", async t => {
  execFileSync("python3", ["scripts/research/verify-notes-size-libreoffice.py"], { stdio: "pipe" });
  const evidence = JSON.parse(await readFile("docs/testing/notes-size-libreoffice-evidence.json", "utf8"));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const core = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8"));
  const parser = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const record of evidence.records) {
    const bytes = await readFile(record.file); assert.equal(createHash("sha256").update(bytes).digest("hex"), record.sha256);
    const report = await inspectPptxNotes(bytes); assert.equal(report.geometryPolicy, "unresolved-source-only"); assert.equal(report.authoredMapping, "not-performed");
    assert.equal(report.pages.length, 2); assert.deepEqual(frame(report, record.inventory.slides[0].master), record.inventory.masterImage.frame);
    for (const [index, row] of record.inventory.slides.entries()) {
      const actual = report.pages[index]; assert.equal(actual.slidePart, `/${row.slide}`); assert.equal(actual.notesPart, `/${row.notes}`); assert.equal(actual.masterPart, `/${row.master}`); assert.equal(actual.body.status, "unique-direct");
      assert.deepEqual(frame(report, row.notes), row.localImage.frame);
      const image = report.parts.find(p => p.part === `/${row.notes}`).drawings.find(d => d.role === "slide-image");
      assert.ok(actual.placeholderMatches.some(match => match.source.elementIndex === image.placeholder.source.elementIndex && match.status === "matched-direct-type"));
    }
    for (const projection of ["structure", "content"]) {
      const session = await importPptx(bytes, { projection });
      try {
        if (projection === "content") assert.deepEqual(session.document.slides.map(notesText), record.inventory.slides.map(s => s.notesText));
        const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
      } finally { session.dispose(); }
    }
    const actual = await page.evaluate(async ({ bytes, parser }) => {
      const outputs = []; for (const api of [await import(parser), PPTKit]) {
        const input = Uint8Array.from(bytes), report = await api.inspectPptxNotes(input), session = await api.importPptx(input, { projection: "content" });
        try { const plan = await session.prepareSave({ mode: "no-op" }); if (plan.status !== "ready") throw Error(JSON.stringify(plan.report)); outputs.push({ report, notes: session.document.slides.map(s => s.notes.flatMap(p => p.runs.map(r => r.text)).join("")), bytes: [...(await session.save(plan.plan)).bytes] }); }
        finally { session.dispose(); }
      } return outputs;
    }, { bytes: [...bytes], parser });
    for (const result of actual) { assert.deepEqual(result.report, report); assert.deepEqual(result.notes, record.inventory.slides.map(s => s.notesText)); assert.deepEqual(result.bytes, [...bytes]); }
  }
  assert.deepEqual(requests, []);
});

test("notes text edits preserve LibreOffice local frames and all unrelated payloads over three rounds", async t => {
  execFileSync("python3", ["scripts/research/verify-notes-size-libreoffice.py"], { stdio: "pipe" });
  const evidence = JSON.parse(await readFile("docs/testing/notes-size-libreoffice-evidence.json", "utf8"));
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-notes-boundary-edits-")); t.after(() => rm(dir, { recursive: true, force: true }));
  for (const shape of ["portrait", "landscape", "rounded"]) {
    const record = evidence.records.find(r => r.scenario === shape && r.round === 3), frameBefore = record.inventory.slides[0].localImage.frame;
    let bytes = await readFile(record.file), beforeText = record.inventory.slides[0].notesText;
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v2" });
      try {
        const slide = session.document.slides[0], notes = structuredClone(slide.notes), afterText = `EDITED-${shape}-${round}`; notes[0].runs[0].text = afterText; slide.setNotes(notes);
        const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report));
        const saved = (await session.save(prepared.plan)).bytes, beforeFile = path.join(dir, `${shape}-${round}-before.pptx`), afterFile = path.join(dir, `${shape}-${round}-after.pptx`);
        await Promise.all([writeFile(beforeFile, bytes), writeFile(afterFile, saved)]);
        assert.equal(execFileSync("python3", ["scripts/research/verify-notes-size-edit.py", beforeFile, afterFile, record.inventory.slides[0].notes, beforeText, afterText], { encoding: "utf8" }).trim(), "verified");
        if (shape === "portrait" && round === 1) {
          const mutate = 'import sys,zipfile\nsrc,dst,name,old,new=sys.argv[1:]\nwith zipfile.ZipFile(src) as a,zipfile.ZipFile(dst,"w") as b:\n for entry in a.namelist():\n  data=a.read(entry)\n  if entry==name:\n   assert old.encode() in data\n   data=data.replace(old.encode(),new.encode(),1)\n  b.writestr(entry,data)';
          for (const [part, old, replacement] of [[record.inventory.slides[0].notes, `cx="${frameBefore.width}"`, `cx="${frameBefore.width + 1}"`], [record.inventory.slides[1].notes, record.inventory.slides[1].notesText, "UNAPPROVED"]]) {
            const damaged = path.join(dir, "damaged.pptx"); execFileSync("python3", ["-c", mutate, afterFile, damaged, part, old, replacement]);
            assert.throws(() => execFileSync("python3", ["scripts/research/verify-notes-size-edit.py", beforeFile, damaged, record.inventory.slides[0].notes, beforeText, afterText], { stdio: "pipe" }));
          }
        }
        const report = await inspectPptxNotes(saved); assert.deepEqual(frame(report, record.inventory.slides[0].notes), frameBefore);
        assert.deepEqual(frame(report, record.inventory.slides[1].notes), record.inventory.slides[1].localImage.frame);
        const reopened = await importPptx(saved, { projection: "content", editSaveProfile: "text-runs-v2" });
        try { assert.deepEqual(reopened.document.slides.map(notesText), [afterText, record.inventory.slides[1].notesText]); }
        finally { reopened.dispose(); }
        bytes = saved; beforeText = afterText;
      } finally { session.dispose(); }
    }
  }
});
