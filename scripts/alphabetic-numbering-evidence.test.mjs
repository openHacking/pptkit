import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";
import test from "node:test";

const verify = file => execFileSync("python3", ["scripts/research/verify-alphabetic-numbering.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" });
test("retained native numbering observation rejects altered evidence", async t => {
  assert.deepEqual(JSON.parse(verify()), { verified: true, markers: 6, pages: 1 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-alpha-evidence-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, "bad.json"), original = JSON.parse(await readFile("docs/testing/alphabetic-numbering-libreoffice-evidence.json", "utf8"));
  for (const alter of [e => { e.words[8].text = "aaa."; }, e => { e.sourceSha256 = "0".repeat(64); }, e => { e.rasterSha256 = "0".repeat(64); }, e => { e.sourceBullets[4].startAt = 54; }]) {
    const changed = structuredClone(original); alter(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => verify(file));
  }
});
