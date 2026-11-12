import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";

test("PowerPoint elbow evidence binds saved packages and both PDF observations", async t => {
  const verify = file => JSON.parse(execFileSync("python3", ["scripts/research/verify-native-elbows.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" }));
  assert.deepEqual(verify(), { verified: true, pages: 4, samples: 16 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-elbow-evidence-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/native-elbow-powerpoint-evidence.json", "utf8"));
  const file = path.join(dir, "changed.json");
  for (const mutate of [
    e => { e.sourceSha256 = "0".repeat(64); },
    e => { e.savedSha256 = "0".repeat(64); },
    e => { e.reopenedPdfSha256 = "0".repeat(64); },
    e => { e.pages[3].samples[0].outline[0]++; },
    e => { e.pages[2].axSha256 = "0".repeat(64); },
  ]) {
    const changed = structuredClone(original);
    mutate(changed);
    await writeFile(file, JSON.stringify(changed));
    assert.throws(() => verify(file));
  }
});
