import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, writeFile, copyFile } from "node:fs/promises";
import path from "node:path";
import { createHash } from "node:crypto";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { distinctNotesFontOverrideParts } from "../tests/fixtures/notes-master-isolation.mjs";

if (process.argv.length !== 4) throw Error("Usage: node scripts/prepare-native-notes-font-themes.mjs <verified-notes-edit-run> <verified-isolation-run>");
const editRoot = path.resolve(process.argv[2]), isolationRoot = path.resolve(process.argv[3]);
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const styles = JSON.parse(await readFile("docs/testing/native-notes-styles-evidence.json", "utf8"));
const isolation = JSON.parse(await readFile("docs/testing/notes-master-isolation-evidence.json", "utf8"));
assert.equal(path.resolve(styles.root), editRoot);
assert.equal(path.resolve(isolation.artifactsDirectory), isolationRoot);
const root = path.resolve(".tmp/native-notes-font-themes");
const inputs = [];
for (const variant of ["A", "B"]) {
  const from = path.join(editRoot, variant);
  const bytes = await readFile(path.join(from, "regular-round3.pptx")), pdf = await readFile(path.join(from, "regular-round3-notes.pdf"));
  const evidence = styles.groups.find(group => group.variant === variant);
  assert.equal(hash(bytes), evidence.inputs.find(input => input.stage === "regular" && input.round === 3).sha256);
  assert.equal(hash(pdf), evidence.independent.stages.find(stage => stage.stage === "regular").rounds[2].pdf.sha256);
  const expected = JSON.parse(await readFile(path.join(from, "expected.json"), "utf8"));
  Object.assign(expected, { group: variant, sourceAssembly: "Exact verified native regular Ω-style third-round input copy; original source assembler limitations retained.", themeBaseline: { pptx: hash(bytes), pdf: hash(pdf), directory: from }, themeOperation: "Office slide theme; actual notes-theme response must be observed" });
  inputs.push({ variant, from, expected });
}
const override = isolation.groups.find(group => group.classification === "single-master-notes-font-override-source");
assert.ok(override, "Expected verified single-master source classification");
const source = path.join(isolationRoot, "override"), bytes = await readFile(path.join(source, "round-0.pptx"));
assert.equal(hash(bytes), override.hashes.find(item => item.file === "round-0.pptx").sha256);
const controls = await readFile(path.join(source, "controls.json"));
await mkdir(root, { recursive: true });
const directory = await mkdtemp(path.join(root, "run-"));
for (const { variant, from, expected } of inputs) {
  const folder = path.join(directory, variant);
  await mkdir(folder);
  await copyFile(path.join(from, "regular-round3.pptx"), path.join(folder, "baseline.pptx"));
  await copyFile(path.join(from, "regular-round3-notes.pdf"), path.join(folder, "baseline-notes.pdf"));
  await writeFile(path.join(folder, "expected.json"), JSON.stringify(expected, null, 2) + "\n");
}
const folder = path.join(directory, "override");
await mkdir(folder);
await copyFile(path.join(source, "round-0.pptx"), path.join(folder, "source.pptx"));
const distinct = zip([...distinctNotesFontOverrideParts(readGeneratedZip(bytes))].map(([name, body]) => ({ name, body }))).bytes;
await writeFile(path.join(folder, "control-distinct-overrides.pptx"), distinct);
await writeFile(path.join(folder, "controls.json"), controls);
await writeFile(path.join(folder, "expected.json"), JSON.stringify({ group: "override", sourceSha256: hash(bytes), distinctSourceSha256: hash(distinct), sourceDirectory: source, sourceAssembly: override.sourceAssembly, sourceConformance: "Single actual/listed notes master; full schema not certified. Shared override is a diagnostic input; native positive path uses distinct per-notes parts.", requiresNativeBaseline: true, themeOperation: "Office slide theme; inspect actual notes/master/theme/override response", fontEnvironment: inputs[0].expected.fonts }, null, 2) + "\n");
await writeFile(path.join(directory, "progress.json"), JSON.stringify({ recordedAt: new Date().toISOString(), state: "prepared", groups: { A: { nativeThemeRounds: 0 }, B: { nativeThemeRounds: 0 }, override: { nativeBaselineRounds: 0, nativeThemeRounds: 0 } }, resume: "Use this explicit run; retain verified baselines. Capture override native baseline before Office theme operation. No unlisted-master positive native admission." }, null, 2) + "\n");
await writeFile(path.join(root, "latest.txt"), directory + "\n");
console.log(directory);
