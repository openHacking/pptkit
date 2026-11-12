import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";

const evidence = JSON.parse(await readFile(process.argv[2] ?? "docs/testing/tab-stop-wrap-matrix-browser-powerpoint-evidence.json", "utf8"));
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
execFileSync("python3", ["scripts/research/verify-native-tab-wrap-matrix.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-stop-wrap-matrix-powerpoint-evidence.json");
const bytes = await readFile(evidence.source), source = JSON.parse(bytes);
assert.equal(sha(bytes), evidence.sourceSha256); assert.equal(sha(await readFile("scripts/inspect-native-tab-wrap-matrix-cross-app.mjs")), evidence.runnerSha256);
assert.deepEqual(evidence.fonts.map(font => [font.weight, font.file]), [["Regular", "/System/Library/Fonts/Supplemental/Arial.ttf"], ["Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]]);
for (const font of evidence.fonts) assert.equal(sha(await readFile(font.file)), font.sha256);
assert.deepEqual(evidence.selectedFonts.map(row => row.map(font => font.postScriptName)), [["ArialMT"], ["Arial-BoldMT"]]);
assert.ok(evidence.selectedFonts.every(row => row.length === 1 && row[0].isCustomFont && row[0].glyphCount > 0));
const labels = ["I WORD", "WIDE WORD", "MIXED WORD"], decimals = ["1.2 WORD", "123.45 WORD", "0.678 WORD"];
const expected = source.observations.filter(page => page.align !== "right").flatMap(page => page.rows.map((rows, index) => {
  const origin = rows[0].chars[0].bbox[0], scale = (page.guide - origin) / 120;
  const chars = rows.flatMap(row => row.chars.map(char => ({ char: char.char, pdfX: (char.bbox[0] - origin) / scale, pdfLine: Math.round((row.y - rows[0].y) / (33.6 * scale)) })));
  return { width: page.width, align: page.align, text: (page.align === "decimal" ? decimals : labels)[index], bold: index === 2, pdfRight: (Math.max(...rows.flatMap(row => row.chars.map(char => char.bbox[2]))) - origin) / scale, pdfLines: Math.round((rows.at(-1).y - rows[0].y) / (33.6 * scale)) + 1, chars };
}));
assert.equal(expected.length, 36); assert.equal(evidence.items.length, expected.length);
for (const [index, item] of evidence.items.entries()) {
  for (const [key, value] of Object.entries(expected[index])) if (key !== "chars") assert.equal(item[key], value, `${index}/${key}`);
  assert.equal(item.chars.length, expected[index].chars.length);
  for (const [charIndex, char] of item.chars.entries()) {
    for (const [key, value] of Object.entries(expected[index].chars[charIndex])) assert.equal(char[key], value, `${index}/${charIndex}/${key}`);
    assert.equal(char.browserLine, char.pdfLine); assert.ok(Number.isFinite(char.browserX)); assert.equal(char.xDifference, char.browserX - char.pdfX); assert.ok(Math.abs(char.xDifference) < 1);
  }
  assert.equal(item.browserLines, item.pdfLines); assert.ok(Number.isFinite(item.browserRight)); assert.equal(item.rightDifference, item.browserRight - item.pdfRight); assert.ok(Math.abs(item.rightDifference) < 1);
}
console.log(JSON.stringify({ verified: true, comparisons: expected.length, chars: evidence.items.reduce((count, item) => count + item.chars.length, 0), maxRightDifference: Math.max(...evidence.items.map(item => Math.abs(item.rightDifference))), maxCharDifference: Math.max(...evidence.items.flatMap(item => item.chars.map(char => Math.abs(char.xDifference)))) }));
