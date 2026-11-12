import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { normalizePresentation } from "../packages/core/dist/index.js";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-real-font-combined.mjs <native-artifacts-directory>");
const directory = path.resolve(process.argv[2]), exec = promisify(execFile);
const expected = JSON.parse(await readFile(path.join(directory, "expected.json"), "utf8"));
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
let negativeControls;
for (const file of ["real-font-combined.py", "real-font-combined-checks.py"]) {
  const result = await exec("python3", [path.resolve("scripts/research/oracles", file), directory], { timeout: 60000, maxBuffer: 1024 * 1024 });
  console.log(result.stdout.trim());
  if (file.endsWith("-checks.py")) negativeControls = JSON.parse(result.stdout);
}
const evidence = JSON.parse(await readFile(path.join(directory, "combined-inspection.json"), "utf8"));
const files = [{ name: "combined-fonts", stage: "source", round: 0 }, ...["edit", "bold", "italic", "theme"].flatMap(stage => [0, 1, 2, 3].map(round => ({ name: `combined-${stage}${round ? "-round" + round : ""}`, stage, round })))];
const projections = [];
for (const { name, stage, round } of files) {
  const bytes = await readFile(path.join(directory, name + ".pptx"));
  const session = await importPptx(bytes, { projection: "content", textFontProfile: "script-families-v3" });
  try {
    const document = normalizePresentation(session.document), before = structuredClone(document);
    assert.equal(document.slides.length, 1);
    const slide = document.slides[0], label = slide.elements.find(e => e.name === "Label"), probe = slide.elements.find(e => e.name === "Probe");
    assert.equal(label.content.flatMap(p => p.runs).map(r => r.text).join(""), "combined-scripts");
    assert.equal(probe.content.length, 1);
    const first = expected.expectedRuns[0];
    const wanted = stage === "source" ? expected.expectedRuns : stage === "edit" && round === 0 ? [{ ...first, text: "LatinX " }, ...expected.expectedRuns.slice(1)] : [{ ...first, text: "LatinX" }, { ...first, text: " " }, ...expected.expectedRuns.slice(1)];
    const runs = probe.content[0].runs; assert.equal(runs.length, wanted.length);
    for (const [i, run] of runs.entries()) {
      assert.equal(run.text, wanted[i].text); assert.equal(run.style.language, wanted[i].language);
      assert.deepEqual([run.style.fontFamily, run.style.eastAsianFontFamily, run.style.complexScriptFontFamily], Array(3).fill(wanted[i].family));
      assert.equal(run.style.fontSize, 48);
      assert.equal(run.style.bold, stage === "bold" && i === 0); assert.equal(run.style.italic, stage === "italic" && i === 0);
      assert.equal(Object.hasOwn(run.style, "scriptFontFamilies"), false);
    }
    const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
    assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), bytes);
    assert.deepEqual(normalizePresentation(session.document), before);
    projections.push({ name, sourceSha256: hash(bytes), noOpExact: true, authorMapReconstructed: false, runs });
  } finally { session.dispose(); }
}
evidence.parser = { profile: "script-families-v3", artifacts: projections };
evidence.sourceRunDefinitions = expected.expectedRuns;
evidence.fonts = expected.fonts;
evidence.negativeControls = negativeControls;
evidence.recordedAt = new Date().toISOString();
await writeFile("docs/testing/real-font-combined-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ sourceAndNativeArtifacts: files.length, negativeControls: negativeControls.rejectedCount, parserNoOpExact: projections.length, manifest: "docs/testing/real-font-combined-evidence.json" }));
