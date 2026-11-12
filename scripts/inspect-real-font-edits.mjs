import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { normalizePresentation } from "../packages/core/dist/index.js";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-real-font-edits.mjs <native-artifacts-directory>");
const directory = path.resolve(process.argv[2]);
const expected = JSON.parse(await readFile(path.join(directory, "expected.json"), "utf8"));
const exec = promisify(execFile);
for (const file of ["real-font-edit.py", "real-font-edit-checks.py"]) {
  const result = await exec("python3", [path.resolve("scripts/research/oracles", file), directory], { timeout: 60000, maxBuffer: 1024 * 1024 });
  console.log(result.stdout.trim());
}
const evidence = JSON.parse(await readFile(path.join(directory, "edit-inspection.json"), "utf8"));
const projections = [];
for (const [stage, [bold, italic]] of Object.entries({ edit: [false, false], bold: [true, false], italic: [false, true], theme: [false, false] })) {
  for (const suffix of ["", "-round1", "-round2", "-round3"]) {
    const name = `fresh-${stage}-sc${suffix}`, bytes = await readFile(path.join(directory, name + ".pptx"));
    const session = await importPptx(bytes, { projection: "content", textFontProfile: "script-families-v3" });
    try {
      const document = normalizePresentation(session.document), before = structuredClone(document);
      assert.equal(document.slides.length, expected.controls.length);
      const controls = [];
      for (const [index, control] of expected.controls.entries()) {
        const slide = document.slides[index], label = slide.elements.find(e => e.name === "Label"), probe = slide.elements.find(e => e.name === "Probe");
        assert.equal(label.content.flatMap(p => p.runs).map(r => r.text).join(""), control.id);
        const runs = probe.content.flatMap(p => p.runs), edited = control.id === "map-sc";
        const wanted = edited ? [{ text: "骨直令返", language: "zh-CN" }, { text: "Ω", language: "el-GR" }] : [{ text: control.text, language: control.language }];
        assert.equal(runs.length, wanted.length);
        for (const [i, run] of runs.entries()) {
          assert.equal(run.text, wanted[i].text); assert.equal(run.style.language, wanted[i].language);
          assert.deepEqual([run.style.fontFamily, run.style.eastAsianFontFamily, run.style.complexScriptFontFamily], control.expectedSlots);
          assert.equal(run.style.fontSize, 48);
          assert.equal(run.style.bold, edited && bold); assert.equal(run.style.italic, edited && italic);
          assert.equal(Object.hasOwn(run.style, "scriptFontFamilies"), false, "unused author map must not be reconstructed from materialized slots");
        }
        controls.push({ id: control.id, runs });
      }
      const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
      const saved = await session.save(prepared.plan); assert.deepEqual(Buffer.from(saved.bytes), bytes);
      assert.deepEqual(normalizePresentation(session.document), before);
      projections.push({ name, sourceSha256: createHash("sha256").update(bytes).digest("hex"), noOpExact: true, authorMapReconstructed: false, controls });
    } finally { session.dispose(); }
  }
}
evidence.parser = { profile: "script-families-v3", artifacts: projections };
evidence.negativeControls = ["text", "language", "restored-Greek-map", "merged-run", "missing-bold", "missing-italic", "unchanged-theme", "wrong-actual-bold-face", "italic-without-shear", "missing-appended-glyph"];
evidence.recordedAt = new Date().toISOString();
await writeFile("docs/testing/real-font-edit-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ stages: 4, roundsPerStage: 3, nativePptxArtifacts: projections.length, parserNoOpExact: projections.length, manifest: "docs/testing/real-font-edit-evidence.json" }));
