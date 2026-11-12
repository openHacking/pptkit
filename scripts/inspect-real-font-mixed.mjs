import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { normalizePresentation } from "../packages/core/dist/index.js";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-real-font-mixed.mjs <native-artifacts-directory>");
const directory = path.resolve(process.argv[2]), exec = promisify(execFile);
const expected = JSON.parse(await readFile(path.join(directory, "expected.json"), "utf8"));
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
let negativeControls;
for (const file of ["real-font-mixed.py", "real-font-mixed-checks.py"]) {
  const result = await exec("python3", [path.resolve("scripts/research/oracles", file), directory], { timeout: 60000, maxBuffer: 1024 * 1024 });
  console.log(result.stdout.trim());
  if (file.endsWith("-checks.py")) negativeControls = JSON.parse(result.stdout);
}
const evidence = JSON.parse(await readFile(path.join(directory, "mixed-inspection.json"), "utf8"));
const stages = { edit: "mixed-edit", bold: "mixed-bold-final", italic: "mixed-italic", theme: "mixed-theme" };
const files = [{ name: "mixed-fonts", stage: "source" }, ...Object.entries(stages).flatMap(([stage, stem]) => ["", "-round1", "-round2", "-round3"].map(suffix => ({ name: stem + suffix, stage })))];
const projections = [];
for (const { name, stage } of files) {
  const bytes = await readFile(path.join(directory, name + ".pptx"));
  const session = await importPptx(bytes, { projection: "content", textFontProfile: "script-families-v3" });
  try {
    const document = normalizePresentation(session.document), before = structuredClone(document), controls = [];
    assert.equal(document.slides.length, 6);
    for (const [index, control] of expected.controls.entries()) {
      const slide = document.slides[index], label = slide.elements.find(e => e.name === "Label"), probe = slide.elements.find(e => e.name === "Probe");
      assert.equal(label.content.flatMap(p => p.runs).map(r => r.text).join(""), control.id);
      const runs = probe.content.flatMap(p => p.runs), language = control.language;
      const wanted = stage === "source" ? control.expectedRuns.map(r => ({ text: r.text, language, family: r.family })) : [
        { text: "Latin", language, family: "Arial" }, { text: "X", language: "en", family: "Arial" },
        { text: " ", language, family: "Arial" }, { text: control.nativeText + " ", language, family: control.expectedRuns[1].family },
        { text: "Ω", language, family: "Georgia" },
      ];
      assert.equal(runs.length, wanted.length);
      for (const [i, run] of runs.entries()) {
        assert.equal(run.text, wanted[i].text); assert.equal(run.style.language, wanted[i].language);
        assert.deepEqual([run.style.fontFamily, run.style.eastAsianFontFamily, run.style.complexScriptFontFamily], Array(3).fill(wanted[i].family));
        assert.equal(run.style.fontSize, 48);
        assert.equal(run.style.bold, stage === "bold" && i < 2); assert.equal(run.style.italic, stage === "italic" && i < 2);
        assert.equal(Object.hasOwn(run.style, "scriptFontFamilies"), false);
      }
      controls.push({ id: control.id, runs });
    }
    const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
    assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), bytes);
    assert.deepEqual(normalizePresentation(session.document), before);
    projections.push({ name, sourceSha256: hash(bytes), noOpExact: true, authorMapReconstructed: false, controls });
  } finally { session.dispose(); }
}
evidence.parser = { profile: "script-families-v3", artifacts: projections };
evidence.negativeControls = negativeControls;
evidence.fonts = expected.fonts;
evidence.authorRunsPerProbe = 1;
evidence.recordedAt = new Date().toISOString();
await writeFile("docs/testing/real-font-mixed-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ sourceAndNativeArtifacts: files.length, pagesPerArtifact: 6, negativeControls: negativeControls.rejectedCount, parserNoOpExact: projections.length, manifest: "docs/testing/real-font-mixed-evidence.json" }));
