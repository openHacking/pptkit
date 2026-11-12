import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { promisify } from "node:util";
import { chromium } from "@playwright/test";

const run = promisify(execFile), sha = bytes => createHash("sha256").update(bytes).digest("hex");
await run("python3", ["scripts/research/verify-native-tab-position.py"]);
const sourcePath = "docs/testing/tab-position-powerpoint-evidence.json", bytes = await readFile(sourcePath), source = JSON.parse(bytes);
const fonts = await Promise.all([["Regular", "/System/Library/Fonts/Supplemental/Arial.ttf"], ["Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]].map(async ([weight, file]) => { const data = await readFile(file); return { weight, file, sha256: sha(data), bytes: data.toString("base64") }; }));
const labels = ["I WORD", "WIDE WORD", "MIXED WORD"], decimals = ["1.2 WORD", "123.45 WORD", "0.678 WORD"];
const items = source.pages.flatMap(page => page.rows.map((rows, index) => {
  const origin = rows[0].chars[0].bbox[0], scale = (page.guide - origin) / page.position;
  const chars = rows.flatMap(row => row.chars.map(char => ({ char: char.char, pdfX: (char.bbox[0] - origin) / scale, pdfLine: Math.round((row.y - rows[0].y) / (33.6 * scale)) })));
  return { width: 150, position: page.position, align: page.align, text: (page.align === "decimal" ? decimals : labels)[index], bold: index === 2, pdfRight: (Math.max(...rows.flatMap(row => row.chars.map(char => char.bbox[2]))) - origin) / scale, pdfLines: Math.round((rows.at(-1).y - rows[0].y) / (33.6 * scale)) + 1, chars };
}));
const imports = Object.fromEntries(await Promise.all([["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/layout/dom", "packages/layout/dist/browser/dom.js"]].map(async ([name, file]) => [name, `data:text/javascript;base64,${(await readFile(file)).toString("base64")}`])));
const browser = await chromium.launch(); let observed, selectedFonts, version;
try {
  version = browser.version(); const page = await browser.newPage(); await page.route("**/*", route => route.abort()); await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  observed = await page.evaluate(async ({ items, fonts }) => {
    const { createPresentation } = await import("@pptkit/core"), { resolveLayout } = await import("@pptkit/layout"), { createBrowserTextMeasurement } = await import("@pptkit/layout/dom");
    const faces = await Promise.all(fonts.map(font => new FontFace("Native Tab Probe", `url(data:font/ttf;base64,${font.bytes})`, { weight: font.weight === "Bold" ? "700" : "400" }).load()));
    for (const face of faces) document.fonts.add(face);
    const adapter = createBrowserTextMeasurement({ document, fonts: [{ family: "Native Tab Probe", faces, covers: ({ text, italic }) => !italic && /^[\x00-\x7f]*$/u.test(text) }], tabWrap: "powerpoint" });
    for (const [id, bold] of [["regular-face", false], ["bold-face", true]]) { const span = document.createElement("span"); span.id = id; span.style.cssText = `font-family:"Native Tab Probe";font-size:28px;font-weight:${bold ? 700 : 400};font-synthesis:none`; span.textContent = "MIXED"; document.body.appendChild(span); }
    try { return items.map(item => {
      const deck = createPresentation({ theme: { fonts: { body: "Native Tab Probe", heading: "Native Tab Probe" } } });
      deck.addSlide({ id: "s" }).addElement({ id: "t", type: "text", box: { x: 60, y: 100, width: item.width, height: 200 }, frame: { margin: 0, wrap: true }, content: [{ style: { lineSpacing: 1.2, tabStops: [{ position: item.position, align: item.align }] }, runs: [{ text: "P\t", style: { fontFamily: "Native Tab Probe", fontSize: 28 } }, { text: item.text, style: { fontFamily: "Native Tab Probe", fontSize: 28, bold: item.bold } }] }] });
      const selected = [], original = Range.prototype.selectNodeContents;
      Range.prototype.selectNodeContents = function(node) {
        if (node instanceof HTMLElement && (node.textContent === "P" || node.textContent === item.text) && node.firstChild?.nodeType === Node.TEXT_NODE) {
          const origin = node.parentElement.parentElement.parentElement.getBoundingClientRect();
          for (let index = 0; index < node.textContent.length; index++) {
            const char = node.textContent[index]; if (/\s/u.test(char)) continue;
            const range = document.createRange(); range.setStart(node.firstChild, index); range.setEnd(node.firstChild, index + 1); const rect = range.getBoundingClientRect();
            selected.push({ char, x: rect.left - origin.left, y: rect.top - origin.top });
          }
        }
        return original.call(this, node);
      };
      try { const analysis = resolveLayout(deck, { textMeasurement: adapter }).analysis; return { diagnostics: analysis.diagnostics, bounds: analysis.measurements[0]?.bounds, selected }; } finally { Range.prototype.selectNodeContents = original; }
    }); } finally { adapter.dispose(); }
  }, { items, fonts });
  const cdp = await page.context().newCDPSession(page); await cdp.send("DOM.enable"); await cdp.send("CSS.enable"); const { root } = await cdp.send("DOM.getDocument"); selectedFonts = [];
  for (const id of ["regular-face", "bold-face"]) { const { nodeId } = await cdp.send("DOM.querySelector", { nodeId: root.nodeId, selector: `#${id}` }); selectedFonts.push((await cdp.send("CSS.getPlatformFontsForNode", { nodeId })).fonts); }
} finally { await browser.close(); }
assert.deepEqual(selectedFonts.map(row => row.map(font => font.postScriptName)), [["ArialMT"], ["Arial-BoldMT"]]);
for (const [index, item] of items.entries()) {
  const result = observed[index]; assert.deepEqual(result.diagnostics, []); item.browserRight = result.bounds.width + result.bounds.x; item.browserLines = Math.round(result.bounds.height / 33.6); item.rightDifference = item.browserRight - item.pdfRight;
  assert.equal(item.browserLines, item.pdfLines, JSON.stringify(item)); assert.ok(Math.abs(item.rightDifference) < 1, JSON.stringify(item));
  assert.equal(result.selected.length, item.chars.length, JSON.stringify(item));
  const firstY = result.selected[0].y;
  for (const [charIndex, char] of item.chars.entries()) {
    const candidate = result.selected[charIndex]; assert.equal(candidate.char, char.char);
    char.browserX = candidate.x; char.browserLine = Math.round((candidate.y - firstY) / 33.6); char.xDifference = char.browserX - char.pdfX;
    assert.equal(char.browserLine, char.pdfLine, JSON.stringify(item)); assert.ok(Math.abs(char.xDifference) < 1, JSON.stringify(item));
  }
}
const result = { scope: "36 center, right and decimal position samples, PowerPoint PDF versus live Chromium using named Arial font candidates; CSS line spacing 1.2 is calibrated to the observed PDF, not a general native spacing claim.", source: sourcePath, sourceSha256: sha(bytes), runnerSha256: sha(await readFile("scripts/inspect-native-tab-position-cross-app.mjs")), browser: version, fonts: fonts.map(({ bytes: _bytes, ...font }) => font), selectedFonts, items };
await writeFile("docs/testing/tab-position-browser-powerpoint-evidence.json", JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ comparisons: items.length, chars: items.reduce((count, item) => count + item.chars.length, 0), maxRightDifference: Math.max(...items.map(item => Math.abs(item.rightDifference))), maxCharDifference: Math.max(...items.flatMap(item => item.chars.map(char => Math.abs(char.xDifference)))) }));
