import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";

const evidence = JSON.parse(await readFile(process.argv[2] ?? "docs/testing/tab-svg-shape-browser-powerpoint-evidence.json", "utf8"));
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
execFileSync("python3", ["scripts/research/verify-native-shape-tabs.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-shape-powerpoint-evidence.json");
const bytes = await readFile(evidence.source), source = JSON.parse(bytes);
assert.equal(sha(bytes), evidence.sourceSha256); assert.equal(sha(await readFile("scripts/inspect-svg-shape-tab-powerpoint.mjs")), evidence.runnerSha256);
assert.deepEqual(evidence.bundles.map(bundle => [bundle.name, bundle.file]), [["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/svg-renderer", "packages/svg-renderer/dist/browser/index.js"]]);
for (const bundle of evidence.bundles) assert.equal(sha(await readFile(bundle.file)), bundle.sha256);
assert.deepEqual(evidence.fonts.map(font => [font.weight, font.file]), [["Regular", "/System/Library/Fonts/Supplemental/Arial.ttf"], ["Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]]);
for (const font of evidence.fonts) assert.equal(sha(await readFile(font.file)), font.sha256);
assert.deepEqual(evidence.selectedFonts.map(row => row.map(font => font.postScriptName)), [["ArialMT"], ["Arial-BoldMT"]]);
assert.ok(evidence.selectedFonts.every(row => row.length === 1 && row[0].isCustomFont && row[0].glyphCount > 0));
const expected = source.slides.flatMap(page => page.rows.map((row, index) => { const origin = row.prefix[0], scale = (page.guide - origin) / 280; return { align: page.slide, text: row.text, bold: index === 2, pdfPrefixRight: (row.prefix[2] - origin) / scale, pdfLeft: (row.word[0] - origin) / scale, pdfRight: (row.word[2] - origin) / scale }; }));
assert.equal(expected.length, 12); assert.equal(evidence.items.length, expected.length);
for (const [index, item] of evidence.items.entries()) {
  for (const [key, value] of Object.entries(expected[index])) assert.equal(item[key], value, `${index}/${key}`);
  assert.equal(item.svgFill, "#F2F5FA");
  assert.equal(item.svgText, `P\t${item.text}`);
  assert.equal(item.svgPrefixWeight, "400"); assert.ok(item.svgSuffixWeights.length > 0 && item.svgSuffixWeights.every(weight => weight === (item.bold ? "700" : "400")));
  for (const edge of ["Left", "Right"]) { assert.ok(Number.isFinite(item[`svg${edge}`])); assert.equal(item[`${edge.toLowerCase()}Difference`], item[`svg${edge}`] - item[`pdf${edge}`]); assert.ok(Math.abs(item[`${edge.toLowerCase()}Difference`]) < 0.5); }
  assert.ok(Number.isFinite(item.svgPrefixRight)); assert.equal(item.prefixDifference, item.svgPrefixRight - item.pdfPrefixRight); assert.ok(Math.abs(item.prefixDifference) < 0.5);
}
console.log(JSON.stringify({ verified: true, comparisons: expected.length, maxDifference: Math.max(...evidence.items.flatMap(item => [Math.abs(item.prefixDifference), Math.abs(item.leftDifference), Math.abs(item.rightDifference)])) }));
