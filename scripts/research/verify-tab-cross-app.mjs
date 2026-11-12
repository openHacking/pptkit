// Recheck retained source PDFs, font identities and recorded comparison arithmetic.
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";

const evidencePath = process.argv[2] ?? "docs/testing/tab-stop-browser-libreoffice-evidence.json";
const evidence = JSON.parse(await readFile(evidencePath, "utf8")), sha = bytes => createHash("sha256").update(bytes).digest("hex");
execFileSync("python3", ["scripts/research/verify-tab-stop-paragraph.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-stop-paragraph-libreoffice-evidence.json");
const sourceBytes = await readFile(evidence.source), source = JSON.parse(sourceBytes);
assert.equal(sha(sourceBytes), evidence.sourceSha256);
assert.equal(sha(await readFile("scripts/inspect-tab-stops-cross-app.mjs")), evidence.runnerSha256);
const wrapper = execFileSync("which", ["soffice"], { encoding: "utf8" }).trim();
const fontDir = path.resolve(path.dirname(wrapper), "../../native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype");
assert.deepEqual(evidence.fonts.map(font => font.weight), ["Regular", "Bold"]);
for (const font of evidence.fonts) {
  assert.equal(font.file, path.join(fontDir, `LiberationSans-${font.weight}.ttf`));
  assert.equal(sha(await readFile(font.file)), font.sha256);
}
assert.deepEqual(evidence.selectedFonts.map(row => row.map(font => font.postScriptName)), [["LiberationSans"], ["LiberationSans-Bold"]]);
assert.ok(evidence.selectedFonts.every(row => row.length === 1 && row[0].isCustomFont && row[0].glyphCount > 0));
const sourceItems = source.cases.flatMap(c => c.pages.flatMap((p, page) => p.samples.map((sample, index) => ({
  case: c.name, indent: c.margin, hanging: c.indent === 0 ? 0 : -c.indent, align: ["left", "center", "right", "decimal"][page], text: sample.text,
  bold: index === 2, pdfWidth: sample.bbox[2] - sample.bbox[0], pdfRight: sample.bbox[2] - 60,
}))));
assert.equal(sourceItems.length, 48); assert.equal(evidence.items.length, sourceItems.length);
const widths = new Map();
for (const [index, expected] of sourceItems.entries()) {
  const item = evidence.items[index];
  for (const [key, value] of Object.entries(expected)) assert.equal(item[key], value, `${index}/${key}`);
  const font = evidence.fonts[item.bold ? 1 : 0], cacheKey = `${font.weight}:${item.text}`;
  if (!widths.has(cacheKey)) {
    const glyphs = JSON.parse(execFileSync("hb-shape", [font.file, item.text, "--output-format=json", "--no-glyph-names", "--font-size=28000", "--direction=ltr", "--language=en"], { encoding: "utf8" }));
    assert.ok(glyphs.every(glyph => glyph.g !== 0)); widths.set(cacheKey, glyphs.reduce((sum, glyph) => sum + glyph.ax, 0) / 1000);
  }
  assert.equal(item.harfBuzzWidth, widths.get(cacheKey));
  assert.ok(Math.abs(item.harfBuzzWidth - item.pdfWidth) < 0.5);
  assert.ok(Number.isFinite(item.browserRight)); assert.equal(item.rightDifference, item.browserRight - item.pdfRight);
  assert.ok(Math.abs(item.rightDifference) < 0.5);
}
console.log(JSON.stringify({ verified: true, comparisons: sourceItems.length, maxRightDifference: Math.max(...evidence.items.map(item => Math.abs(item.rightDifference))) }));
