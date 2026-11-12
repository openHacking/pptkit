import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";

const evidencePath = process.argv[2] ?? "docs/testing/tab-list-browser-libreoffice-evidence.json";
const evidence = JSON.parse(await readFile(evidencePath, "utf8")), sha = bytes => createHash("sha256").update(bytes).digest("hex");
execFileSync("python3", ["scripts/research/verify-tab-lists.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-list-libreoffice-evidence.json");
const sourceBytes = await readFile(evidence.source), source = JSON.parse(sourceBytes);
assert.equal(sha(sourceBytes), evidence.sourceSha256); assert.equal(sha(await readFile("scripts/inspect-tab-list-cross-app.mjs")), evidence.runnerSha256);
const wrapper = execFileSync("which", ["soffice"], { encoding: "utf8" }).trim();
const fontFile = path.resolve(path.dirname(wrapper), "../../native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype/LiberationSans-Regular.ttf");
assert.equal(evidence.font.file, fontFile); assert.equal(sha(await readFile(fontFile)), evidence.font.sha256);
assert.equal(evidence.selectedFonts.length, 1); assert.equal(evidence.selectedFonts[0].postScriptName, "LiberationSans"); assert.ok(evidence.selectedFonts[0].isCustomFont && evidence.selectedFonts[0].glyphCount > 0);
const expected = source.observations.map(item => ({ index: item.index, bullet: item.bullet, indent: item.indent, hanging: item.hanging, align: item.align, text: item.text,
  pdfWidth: item.sample.bbox[2] - item.sample.bbox[0], pdfRight: item.sample.bbox[2] - 60 }));
assert.equal(expected.length, 36); assert.equal(evidence.items.length, expected.length); const widths = new Map();
for (const [index, values] of expected.entries()) {
  const item = evidence.items[index]; for (const [key, value] of Object.entries(values)) assert.equal(item[key], value, `${index}/${key}`);
  if (!widths.has(item.text)) {
    const glyphs = JSON.parse(execFileSync("hb-shape", [fontFile, item.text, "--output-format=json", "--no-glyph-names", "--font-size=28000", "--direction=ltr", "--language=en"], { encoding: "utf8" }));
    assert.ok(glyphs.every(glyph => glyph.g !== 0)); widths.set(item.text, glyphs.reduce((sum, glyph) => sum + glyph.ax, 0) / 1000);
  }
  assert.equal(item.harfBuzzWidth, widths.get(item.text)); assert.ok(Math.abs(item.harfBuzzWidth - item.pdfWidth) < 0.5);
  assert.ok(Number.isFinite(item.browserRight)); assert.equal(item.rightDifference, item.browserRight - item.pdfRight); assert.ok(Math.abs(item.rightDifference) < 0.5);
}
console.log(JSON.stringify({ verified: true, comparisons: expected.length, maxRightDifference: Math.max(...evidence.items.map(item => Math.abs(item.rightDifference))) }));
