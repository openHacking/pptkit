import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { normalizePresentation } from "../packages/core/dist/index.js";

if (process.argv.length !== 3) throw Error("Usage: node scripts/inspect-real-font-multi.mjs <native-artifacts-directory>");
const directory = path.resolve(process.argv[2]), exec = promisify(execFile);
const expected = JSON.parse(await readFile(path.join(directory, "expected.json"), "utf8"));
const hash = data => createHash("sha256").update(data).digest("hex");
let negativeControls;
for (const file of ["real-font-multi.py", "real-font-multi-checks.py"]) {
  const result = await exec("python3", [path.resolve("scripts/research/oracles", file), directory], { timeout: 60000, maxBuffer: 1024 * 1024 });
  console.log(result.stdout.trim());
  if (file.endsWith("-checks.py")) negativeControls = JSON.parse(result.stdout);
}
const evidence = JSON.parse(await readFile(path.join(directory, "multi-inspection.json"), "utf8"));
const candidates = [], hints = [];
const sourceFonts = Object.fromEntries(expected.fonts.map(font => [font.id, font]));
const officeFonts = "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts";
const files = Object.fromEntries(["tc", "ja", "ko", "ar", "he"].map(id => [id, sourceFonts[id].file]));
const desiredFaces = Object.entries(evidence.cases).flatMap(([id, c]) => [c.regular, c.bold].map(postscriptName => ({ postscriptName, file: files[id.slice(4)] })));
desiredFaces.push(...[["Cambria", "Cambria.ttc"], ["Cambria-Bold", "Cambriab.ttf"], ["Cambria-Italic", "Cambriai.ttf"]].map(([postscriptName, file]) => ({ postscriptName, file: path.join(officeFonts, file) })));
for (const { postscriptName, file } of desiredFaces) {
  const lines = (await exec("fc-query", ["-f", "%{family[0]}\t%{postscriptname}\t%{index}\t%{fontversion}\n", file], { timeout: 30000 })).stdout.trim().split("\n");
  const fields = lines.map(line => line.split("\t")).filter(row => row[1] === postscriptName);
  assert.equal(fields.length, 1, `Unique candidate source face: ${postscriptName}`);
  const [family, , index, fontVersion] = fields[0];
  const omega = JSON.parse((await exec("hb-shape", [file, `--face-index=${index}`, "--output-format=json", "--no-glyph-names", "Ω"], { timeout: 30000 })).stdout);
  assert.equal(omega.length, 1);
  assert.equal(omega[0].g === 0, ["GeezaPro", "GeezaPro-Bold", "ArialHebrew", "ArialHebrew-Bold"].includes(postscriptName));
  candidates.push({ family, postscriptName, file, index: Number(index), fontVersion: Number(fontVersion), sha256: hash(await readFile(file)), omegaHarfBuzz: omega, evidenceScope: "local same-PostScript-name candidate and glyph coverage; not a byte-identity claim for the PDF subset" });
}
for (const family of [...new Set(candidates.map(c => c.family))]) for (const style of ["Regular", "Bold", "Italic"]) {
  const fields = (await exec("fc-match", ["-f", "%{family[0]}\t%{postscriptname}\t%{file}\t%{index}", `${family}:style=${style}`], { timeout: 30000 })).stdout.split("\t");
  hints.push({ requestedFamily: family, requestedStyle: style, actualFamily: fields[0], postscriptName: fields[1], file: fields[2], index: Number(fields[3]), requestedFamilyMatched: fields[0] === family });
}
evidence.fonts = { candidates, fontconfigHints: hints, scope: "Application PDF actual faces are checked independently. System fc-match is not treated as Office's face resolver." };
const projections = [];
for (const stage of ["edit", "bold", "italic", "theme"]) for (const suffix of ["", "-round1", "-round2", "-round3"]) {
  const name = `fresh-${stage}-multi${suffix}`, bytes = await readFile(path.join(directory, name + ".pptx"));
  const session = await importPptx(bytes, { projection: "content", textFontProfile: "script-families-v3" });
  try {
    const document = normalizePresentation(session.document), before = structuredClone(document), controls = [];
    assert.equal(document.slides.length, 16);
    for (const [index, control] of expected.controls.entries()) {
      const slide = document.slides[index], label = slide.elements.find(e => e.name === "Label"), probe = slide.elements.find(e => e.name === "Probe"), c = evidence.cases[control.id];
      assert.equal(label.content.flatMap(p => p.runs).map(r => r.text).join(""), control.id);
      const runs = probe.content.flatMap(p => p.runs), wanted = c ? c.runs.map(([text, language]) => ({ text, language })) : [{ text: control.text, language: control.language }];
      assert.equal(runs.length, wanted.length);
      for (const [i, run] of runs.entries()) {
        assert.equal(run.text, wanted[i].text); assert.equal(run.style.language, wanted[i].language);
        assert.deepEqual([run.style.fontFamily, run.style.eastAsianFontFamily, run.style.complexScriptFontFamily], control.expectedSlots);
        assert.equal(run.style.fontSize, 48); assert.equal(run.style.bold, !!c && stage === "bold"); assert.equal(run.style.italic, !!c && stage === "italic");
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
evidence.negativeControls = negativeControls;
evidence.parser = { profile: "script-families-v3", artifacts: projections };
evidence.recordedAt = new Date().toISOString();
await writeFile("docs/testing/real-font-multi-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ stages: 4, editedPages: 5, roundsPerStage: 3, nativeArtifacts: projections.length, parserNoOpExact: projections.length, manifest: "docs/testing/real-font-multi-evidence.json" }));
