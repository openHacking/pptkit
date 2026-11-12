import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";

const evidence = JSON.parse(await readFile(process.argv[2] ?? "docs/testing/tab-stop-wrap-browser-libreoffice-evidence.json", "utf8"));
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
execFileSync("python3", ["scripts/research/verify-tab-stop-wrap.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-stop-wrap-libreoffice-evidence.json");
const sourceBytes = await readFile(evidence.source), source = JSON.parse(sourceBytes);
assert.equal(sha(sourceBytes), evidence.sourceSha256);
assert.equal(sha(await readFile("scripts/inspect-tab-wrap-cross-app.mjs")), evidence.runnerSha256);
const wrapper = execFileSync("which", ["soffice"], { encoding: "utf8" }).trim();
const fontDir = path.resolve(path.dirname(wrapper), "../../native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype");
assert.deepEqual(evidence.fonts.map(font => font.weight), ["Regular", "Bold"]);
for (const font of evidence.fonts) { assert.equal(font.file, path.join(fontDir, `LiberationSans-${font.weight}.ttf`)); assert.equal(sha(await readFile(font.file)), font.sha256); }
assert.deepEqual(evidence.selectedFonts.map(row => row.map(font => font.postScriptName)), [["LiberationSans"], ["LiberationSans-Bold"]]);
assert.ok(evidence.selectedFonts.every(row => row.length === 1 && row[0].isCustomFont && row[0].glyphCount > 0));
const expected = source.cases.filter(item => !item.plain).flatMap(c => c.pages.flatMap((page, pageIndex) => [0, 1, 2].map(index => {
  const prefix = page[2 * index], sample = page[2 * index + 1];
  return { case: c.name, width: c.width, wrap: c.wrap === "wrap", align: ["left", "center", "right", "decimal"][pageIndex], text: sample.text, bold: index === 2, pdfRight: Math.max(prefix.bbox[2], sample.bbox[2]) - 60, pdfWrapped: sample.bbox[1] > prefix.bbox[1] + 10 };
})));
assert.equal(expected.length, 36); assert.equal(evidence.items.length, expected.length);
for (const [index, item] of evidence.items.entries()) {
  for (const [key, value] of Object.entries(expected[index])) assert.equal(item[key], value, `${index}/${key}`);
  assert.equal(item.browserWrapped, item.pdfWrapped); assert.ok(Number.isFinite(item.browserRight));
  assert.equal(item.rightDifference, item.browserRight - item.pdfRight); assert.ok(Math.abs(item.rightDifference) < 0.5);
}
console.log(JSON.stringify({ verified: true, comparisons: expected.length, maxRightDifference: Math.max(...evidence.items.map(item => Math.abs(item.rightDifference))) }));
