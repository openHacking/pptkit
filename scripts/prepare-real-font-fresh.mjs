import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { nativeFontFamilies } from "./research/font-native-cases.mjs";
import { createRealFontFreshFixture } from "../tests/fixtures/real-font-fresh.mjs";

if (process.argv.length !== 2) throw Error("Usage: node scripts/prepare-real-font-fresh.mjs");
const hash = data => createHash("sha256").update(data).digest("hex");
const inventory = JSON.parse(await readFile("docs/testing/real-font-native-evidence.json", "utf8")).fonts;
for (const font of inventory) assert.equal(hash(await readFile(font.file)), font.sha256, `Recorded installed font changed: ${font.family}`);
const fixture = createRealFontFreshFixture(createPresentation, nativeFontFamilies), before = normalizePresentation(fixture.document);
const generated = await generatePptx(fixture.document);
assert.deepEqual(normalizePresentation(fixture.document), before);
const warnings = generated.warnings.filter(warning => warning.code === "script-font-map-materialized");
assert.equal(warnings.length, 8);
for (const control of fixture.controls) assert.equal(warnings.some(warning => warning.slideId === control.slideId && warning.elementId === control.elementId), control.mode === "map");
await mkdir(".tmp/real-font-fresh", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/real-font-fresh/run-"));
const file = path.join(directory, "fresh-fonts.pptx"); await writeFile(file, generated.bytes);
const variants = [];
for (const style of ["Regular", "Bold", "Italic"]) {
  const fields = (await promisify(execFile)("fc-match", ["-f", "%{family[0]}\t%{postscriptname}\t%{file}\t%{index}\t%{weight}\t%{slant}", `${nativeFontFamilies.sc}:style=${style}`], { timeout: 30000 })).stdout.split("\t");
  const [family, postscriptName, fontFile, index, weight, slant] = fields;
  variants.push({ requestedStyle: style, family, postscriptName, file: fontFile, index: Number(index), weight: Number(weight), slant: Number(slant), sha256: hash(await readFile(fontFile)) });
}
const expected = { recordedAt: new Date().toISOString(), sourceSha256: hash(generated.bytes), size: before.size, controls: fixture.controls, sourceMap: fixture.sourceMap, fonts: inventory, variants, warnings: generated.warnings, document: before };
await writeFile(path.join(directory, "expected.json"), JSON.stringify(expected, null, 2) + "\n");
await writeFile(".tmp/real-font-fresh/latest.txt", directory);
console.log(JSON.stringify({ directory, file, pages: fixture.controls.length, mapWarnings: warnings.length, sourceSha256: expected.sourceSha256, variants }));
