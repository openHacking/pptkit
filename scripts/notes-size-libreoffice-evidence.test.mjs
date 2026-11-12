import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";

const verify = file => execFileSync("python3", ["scripts/research/verify-notes-size-libreoffice.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] });
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
test("boundary notes evidence checks actual files and rejects rehashed text, relationship and geometry damage", async t => {
  const result = JSON.parse(verify()); assert.equal(result.records, 12); assert.equal(result.slides, 24); assert.equal(result.nativeRounds, 9);
  const parent = await mkdtemp(path.resolve(".tmp/notes-size-libreoffice/check-")); t.after(() => rm(parent, { recursive: true, force: true }));
  const evidence = JSON.parse(await readFile("docs/testing/notes-size-libreoffice-evidence.json", "utf8"));
  const file = path.join(parent, "manifest.json");
  for (const alter of [e => { e.records[1].inventory.slides[0].notesText = "forged"; }, e => { e.records[2].sha256 = "0".repeat(64); }, e => { e.records.pop(); }, e => { e.logs[0].exitCode = 1; }]) {
    const changed = structuredClone(evidence); alter(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => verify(file));
  }
  for (const [name, pattern, replacement] of [
    ["ppt/notesSlides/notesSlide1.xml", "NOTES-portrait-FIRST", "NOTES-portrait-WRONG"],
    ["ppt/notesSlides/_rels/notesSlide1.xml.rels", "../slides/slide1.xml", "../slides/slide2.xml"],
    ["ppt/notesSlides/notesSlide1.xml", 'cx="64800"', 'cx="64801"'],
  ]) {
    const changed = structuredClone(evidence), damaged = path.join(parent, "damaged.pptx");
    const script = 'import sys,zipfile\nsrc,dst,name,old,new=sys.argv[1:]\nwith zipfile.ZipFile(src) as a,zipfile.ZipFile(dst,"w") as b:\n for entry in a.namelist():\n  data=a.read(entry)\n  if entry==name:\n   assert old.encode() in data\n   data=data.replace(old.encode(),new.encode(),1)\n  b.writestr(entry,data)';
    execFileSync("python3", ["-c", script, evidence.records[1].file, damaged, name, pattern, replacement]); const bytes = await readFile(damaged);
    changed.records[1].file = path.relative(process.cwd(), damaged); changed.records[1].sha256 = sha(bytes);
    await writeFile(file, JSON.stringify(changed)); assert.throws(() => verify(file));
  }
});
