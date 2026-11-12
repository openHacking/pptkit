import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createHash } from "node:crypto";
import { normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";

if (process.argv.length !== 4) throw Error("Usage: node scripts/inspect-native-table-fonts.mjs <reference-and-empty-collection-directory> <filled-collection-directory>");
const exec = promisify(execFile), hash = bytes => createHash("sha256").update(bytes).digest("hex"), groups = [];
for (const input of process.argv.slice(2)) {
  const directory = path.resolve(input), expected = JSON.parse(await readFile(path.join(directory, "expected.json"), "utf8"));
  let negativeControls;
  for (const file of ["native-table-fonts.py", "native-table-fonts-checks.py"]) {
    const output = (await exec("python3", [path.resolve("scripts/research/oracles", file), directory], { timeout: 60000, maxBuffer: 4 * 1024 * 1024 })).stdout;
    console.log(output.trim()); if (file.endsWith("-checks.py")) negativeControls = JSON.parse(output);
  }
  const independent = JSON.parse(await readFile(path.join(directory, "table-fonts-inspection.json"), "utf8")), projections = [];
  for (let round = 0; round <= 3; round++) {
    const stem = round ? `table-fonts-round${round}` : "table-fonts", bytes = await readFile(path.join(directory, stem + ".pptx"));
    const session = await importPptx(bytes, expected.options);
    try {
      const document = normalizePresentation(session.document), before = structuredClone(document), controls = [];
      assert.equal(document.slides.length, expected.controls.length);
      for (const [index, control] of expected.controls.entries()) {
        const slide = document.slides[index], label = slide.elements.find(e => e.name === "Label"), table = slide.elements.find(e => e.name === "ProbeTable");
        assert.equal(label.content.flatMap(p => p.runs).map(r => r.text).join(""), control.id);
        assert.equal(table.rows.length, 1); assert.equal(table.rows[0].cells.length, 1);
        const run = table.rows[0].cells[0].content[0].runs[0];
        assert.equal(run.text, control.text); assert.equal(run.style.language, control.language); assert.equal(run.style.fontSize, 48);
        assert.deepEqual(run.style.scriptFontFamilies, expected.maps[control.kind]);
        const selected = resolveTextFontSegments(run.text, run.style, document.theme);
        assert.deepEqual([...new Set(selected.map(s => s.family))], [control.targetFamily]);
        const rules = session.report.tableStyleInheritance.properties.filter(p => p.target.part === `/ppt/slides/slide${index + 1}.xml` && p.property.startsWith("run.scriptFontFamilies."));
        assert.equal(rules.length, 8);
        for (const rule of rules) { assert.equal(rule.status, "mapped"); assert.equal(rule.value, expected.maps[control.kind][rule.property.split(".").at(-1)]); }
        controls.push({ id: control.id, style: run.style, declaredSelection: selected, mappedRules: rules.length });
      }
      const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
      assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), Buffer.from(bytes)); assert.deepEqual(normalizePresentation(session.document), before);
      projections.push({ file: stem + ".pptx", sha256: hash(bytes), noOpExact: true, controls });
    } finally { session.dispose(); }
  }
  groups.push({ expectedControls: expected.controls, sourceMaps: expected.maps, sourceFonts: expected.fonts, independent, negativeControls, parser: { profile: "script-families-v3", artifacts: projections } });
}
const office = "/Applications/Microsoft PowerPoint.app", candidateSpecs = [["MalgunGothic", "malgun.ttf", "한글"], ["Batang", "batang.ttc", "한글"], ["SimSun", "Simsun.ttc", "骨直令返"]], candidates = [];
for (const [postscriptName, fileName, sample] of candidateSpecs) {
  const file = path.join(office, "Contents/Resources/DFonts", fileName);
  const rows = (await exec("fc-query", ["-f", "%{family[0]}\t%{postscriptname}\t%{index}\t%{fontversion}\n", file])).stdout.trim().split("\n").map(line => line.split("\t")).filter(row => row[1] === postscriptName);
  assert.equal(rows.length, 1); const [family, , index, version] = rows[0];
  const shaped = JSON.parse((await exec("hb-shape", [file, `--face-index=${index}`, "--output-format=json", "--no-glyph-names", sample])).stdout);
  assert.ok(shaped.length && shaped.every(glyph => glyph.g !== 0));
  const hint = (await exec("fc-match", ["-f", "%{family[0]}\t%{postscriptname}", family])).stdout.split("\t");
  candidates.push({ family, postscriptName, file, index: Number(index), version: Number(version), sha256: hash(await readFile(file)), sample, harfBuzz: shaped, systemFontconfigHint: { family: hint[0], postscriptName: hint[1] }, scope: "same-PostScript-name Office-bundled candidate and coverage; not PDF-subset byte identity" });
}
const appInfo = JSON.parse((await exec("python3", ["-c", "import json,plistlib,sys; p=plistlib.load(open(sys.argv[1],'rb')); print(json.dumps({k:p[k] for k in ['CFBundleShortVersionString','CFBundleVersion']}))", path.join(office, "Contents/Info.plist")])).stdout);
const matches = groups.reduce((n, group) => n + group.independent.pdf.matches, 0), differences = groups.reduce((n, group) => n + group.independent.pdf.differences, 0);
assert.equal(matches, 20); assert.equal(differences, 16);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), application: { name: "Microsoft PowerPoint for Mac", ...appInfo }, groups, candidates, total: { controls: 36, matchedDeclaredSelections: matches, nativeDifferences: differences, nativeRoundsPerGroup: 3, parserExactNoOps: 8, rejectedControls: groups.reduce((n, group) => n + group.negativeControls.rejectedCount, 0) }, nativeEditing: "not tested", independentNotesTheme: "not tested", sourceFontWrites: "blocked", scope: "Named table source declarations, v3 rule transport and native actual font differences; no native-parity or full shaping claim." };
await writeFile("docs/testing/native-table-font-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify(evidence.total));
