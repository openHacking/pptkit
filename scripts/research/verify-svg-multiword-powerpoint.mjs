import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";

const sha = data => createHash("sha256").update(data).digest("hex");
const evidence = JSON.parse(await readFile(process.argv[2] ?? "docs/testing/tab-svg-multiword-browser-powerpoint-evidence.json", "utf8"));
execFileSync("python3", ["scripts/research/verify-native-tab-phrases.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-phrases-powerpoint-evidence.json");
const sourceBytes = await readFile(evidence.source), source = JSON.parse(sourceBytes);
assert.equal(sha(sourceBytes), evidence.sourceSha256);
assert.equal(sha(await readFile("scripts/inspect-svg-multiword-powerpoint.mjs")), evidence.runnerSha256);
assert.deepEqual(evidence.bundles.map(item => [item.name, item.file]), [["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/svg-renderer", "packages/svg-renderer/dist/browser/index.js"]]);
for (const bundle of evidence.bundles) assert.equal(sha(await readFile(bundle.file)), bundle.sha256);
assert.deepEqual(evidence.fonts.map(item => [item.weight, item.file]), [["Regular", "/System/Library/Fonts/Supplemental/Arial.ttf"], ["Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]]);
for (const font of evidence.fonts) assert.equal(sha(await readFile(font.file)), font.sha256);
assert.equal(evidence.selectedFont.length, 1); assert.equal(evidence.selectedFont[0].postScriptName, "ArialMT"); assert.ok(evidence.selectedFont[0].isCustomFont && evidence.selectedFont[0].glyphCount > 0);
assert.deepEqual(evidence.parts, ["P\t", "I WIDE WORD"]); assert.ok(evidence.warnings.includes("text-tab-layout-unsupported"));
const page = source.observations.find(item => item.width === 220 && item.align === "center"), row = page.rows[0];
assert.equal(row.length, 1); const original = row[0].chars, origin = original[0].bbox[0], scale = (page.guide - origin) / 120;
assert.equal(original.map(char => char.char).join(""), "PIWIDEWORD"); assert.equal(evidence.chars.length, original.length);
for (const [index, char] of evidence.chars.entries()) {
  assert.equal(char.char, original[index].char);
  assert.equal(char.pdfX, (original[index].bbox[0] - origin) / scale); assert.equal(char.pdfRight, (original[index].bbox[2] - origin) / scale);
  assert.ok(Number.isFinite(char.svgX) && Number.isFinite(char.svgRight));
  assert.equal(char.xDifference, char.svgX - char.pdfX); assert.equal(char.rightDifference, char.svgRight - char.pdfRight);
  assert.ok(Math.abs(char.xDifference) < 0.5 && Math.abs(char.rightDifference) < 0.5);
}
console.log(JSON.stringify({ verified: true, characters: evidence.chars.length, maxDifference: Math.max(...evidence.chars.flatMap(char => [Math.abs(char.xDifference), Math.abs(char.rightDifference)])) }));
