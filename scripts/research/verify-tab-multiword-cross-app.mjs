import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";

const evidence = JSON.parse(await readFile(process.argv[2] ?? "docs/testing/tab-stop-multiword-browser-libreoffice-evidence.json", "utf8"));
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
execFileSync("python3", ["scripts/research/verify-tab-stop-multiword.py"], { stdio: "pipe" });
assert.equal(evidence.source, "docs/testing/tab-stop-multiword-libreoffice-evidence.json");
const bytes = await readFile(evidence.source), source = JSON.parse(bytes);
assert.equal(sha(bytes), evidence.sourceSha256); assert.equal(sha(await readFile("scripts/inspect-tab-multiword-cross-app.mjs")), evidence.runnerSha256);
const wrapper = execFileSync("which", ["soffice"], { encoding: "utf8" }).trim();
const fontDir = path.resolve(path.dirname(wrapper), "../../native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype");
assert.deepEqual(evidence.fonts.map(font => font.weight), ["Regular", "Bold"]);
for (const font of evidence.fonts) { assert.equal(font.file, path.join(fontDir, `LiberationSans-${font.weight}.ttf`)); assert.equal(sha(await readFile(font.file)), font.sha256); }
assert.deepEqual(evidence.selectedFonts.map(row => row.map(font => font.postScriptName)), [["LiberationSans"], ["LiberationSans-Bold"]]);
assert.ok(evidence.selectedFonts.every(row => row.length === 1 && row[0].isCustomFont && row[0].glyphCount > 0));
const expected = source.cases.flatMap(c => c.pages.flatMap((page, pageIndex) => page.map((words, index) => ({ case: c.name, width: c.name === "fit300" ? 300 : 150, wrap: c.name !== "nowrap150", align: ["left", "center", "right", "decimal"][pageIndex], text: `${words[1].text} ${words[2].text}`, bold: index === 2, pdfRight: Math.max(...words.map(word => word.bbox[2])) - 60, pdfLines: new Set(words.map(word => Math.round(word.bbox[1]))).size }))));
assert.equal(expected.length, 27); assert.equal(evidence.items.length, expected.length);
for (const [index, item] of evidence.items.entries()) { for (const [key, value] of Object.entries(expected[index])) assert.equal(item[key], value, `${index}/${key}`); assert.equal(item.browserLines, item.pdfLines); assert.ok(Number.isFinite(item.browserRight)); assert.equal(item.rightDifference, item.browserRight - item.pdfRight); assert.ok(Math.abs(item.rightDifference) < 0.5); }
console.log(JSON.stringify({ verified: true, comparisons: expected.length, maxRightDifference: Math.max(...evidence.items.map(item => Math.abs(item.rightDifference))) }));
