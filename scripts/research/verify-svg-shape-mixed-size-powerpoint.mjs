import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
const evidence = JSON.parse(await readFile(process.argv[2] ?? "docs/testing/tab-svg-shape-mixed-size-powerpoint-evidence.json", "utf8"));
execFileSync("python3", ["scripts/research/verify-shape-mixed-size-powerpoint.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-shape-mixed-size-powerpoint-evidence.json");
const bytes = await readFile(evidence.source), source = JSON.parse(bytes);
assert.equal(sha(bytes), evidence.sourceSha256); assert.equal(sha(await readFile("scripts/inspect-svg-shape-mixed-size-powerpoint.mjs")), evidence.runnerSha256);
assert.deepEqual(evidence.bundles.map(({ name, file }) => [name, file]), ["core", "layout", "svg-renderer"].map(name => [`@pptkit/${name}`, `packages/${name}/dist/browser/index.js`]));
for (const item of evidence.bundles) assert.equal(sha(await readFile(item.file)), item.sha256);
assert.deepEqual(evidence.fonts.map(({ weight, file }) => [weight, file]), [["Regular", "/System/Library/Fonts/Supplemental/Arial.ttf"], ["Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]]);
for (const item of evidence.fonts) assert.equal(sha(await readFile(item.file)), item.sha256);
assert.deepEqual(evidence.selectedFonts.map(row => row.map(f => f.postScriptName)), [["ArialMT"], ["Arial-BoldMT"]]);
assert.ok(evidence.selectedFonts.every(row => row.length === 1 && row[0].isCustomFont && row[0].glyphCount > 0));
const expected = source.pages.flatMap(page => page.rows.map((row, index) => ({ align: page.kind, word: row.word, suffixSize: row.suffixSize, bold: index === 2, characters: row.characters })));
assert.equal(expected.length, 12); assert.equal(evidence.items.length, expected.length);
for (const [index, item] of evidence.items.entries()) {
  const target = expected[index]; for (const key of ["align", "word", "suffixSize", "bold"]) assert.equal(item[key], target[key]);
  assert.equal(item.chars.length, target.characters.length);
  for (const [i, char] of item.chars.entries()) {
    assert.deepEqual(char.native, target.characters[i]); assert.equal(char.char, char.native.char); assert.equal(char.svg.char, char.char);
    for (const key of ["left", "right", "baseline"]) { assert.ok(Number.isFinite(char.svg[key])); assert.equal(char[`${key}Difference`], char.svg[key] - char.native[key]); }
  }
}
const chars = evidence.items.flatMap(item => item.chars);
assert.equal(chars.length, 44); const horizontal = Math.max(...chars.flatMap(c => [Math.abs(c.leftDifference), Math.abs(c.rightDifference)])), baseline = Math.max(...chars.map(c => Math.abs(c.baselineDifference)));
assert.equal(evidence.maxHorizontal, horizontal); assert.equal(evidence.maxBaseline, baseline);
assert.ok(horizontal < .60 && baseline < .85);
console.log(JSON.stringify({ verified: true, samples: evidence.items.length, characters: chars.length, maxHorizontal: horizontal, maxBaseline: baseline }));
