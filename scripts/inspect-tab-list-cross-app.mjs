// Optional local comparison; PDF marker and browser marker may use different fonts.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import { chromium } from "@playwright/test";

const run = promisify(execFile), sha = data => createHash("sha256").update(data).digest("hex");
await run("python3", ["scripts/research/verify-tab-lists.py"]);
const sourcePath = "docs/testing/tab-list-libreoffice-evidence.json", sourceBytes = await readFile(sourcePath), source = JSON.parse(sourceBytes);
const wrapper = (await run("which", ["soffice"])).stdout.trim();
const fontFile = path.resolve(path.dirname(wrapper), "../../native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype/LiberationSans-Regular.ttf");
const fontBytes = await readFile(fontFile);
const items = source.observations.map(item => ({ index: item.index, bullet: item.bullet, indent: item.indent, hanging: item.hanging, align: item.align, text: item.text,
  pdfWidth: item.sample.bbox[2] - item.sample.bbox[0], pdfRight: item.sample.bbox[2] - 60 }));
for (const item of items) {
  const glyphs = JSON.parse((await run("hb-shape", [fontFile, item.text, "--output-format=json", "--no-glyph-names", "--font-size=28000", "--direction=ltr", "--language=en"])).stdout);
  assert.ok(glyphs.every(g => g.g !== 0)); item.harfBuzzWidth = glyphs.reduce((sum, glyph) => sum + glyph.ax, 0) / 1000;
  assert.ok(Math.abs(item.harfBuzzWidth - item.pdfWidth) < 0.5);
}
const imports = Object.fromEntries(await Promise.all([
  ["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/layout/dom", "packages/layout/dist/browser/dom.js"],
].map(async ([name, file]) => [name, `data:text/javascript;base64,${(await readFile(file)).toString("base64")}`])));
const browser = await chromium.launch(); let observed, selectedFonts, browserVersion;
try {
  browserVersion = browser.version(); const page = await browser.newPage(); await page.route("**/*", route => route.abort());
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  observed = await page.evaluate(async ({ bytes, items }) => {
    const { createPresentation } = await import("@pptkit/core"), { resolveLayout } = await import("@pptkit/layout"), { createBrowserTextMeasurement } = await import("@pptkit/layout/dom");
    const face = await new FontFace("List Tab Probe", `url(data:font/ttf;base64,${bytes})`).load(); document.fonts.add(face);
    const adapter = createBrowserTextMeasurement({ document, fonts: [{ family: "List Tab Probe", faces: [face], covers: ({ text, bold, italic }) => !bold && !italic && [...text].every(char => /[\x00-\x7f]/u.test(char) || char === "•") }] });
    const sample = document.createElement("span"); sample.id = "font-control"; sample.style.cssText = 'font-family:"List Tab Probe";font-size:28px;font-synthesis:none'; sample.textContent = "• WIDE"; document.body.appendChild(sample);
    try {
      return items.map(item => {
        const bullet = item.bullet === "symbol" ? { type: "bullet", character: "•" } : { type: "number", startAt: Number(item.bullet.slice(6)) };
        const deck = createPresentation({ theme: { fonts: { body: "List Tab Probe", heading: "List Tab Probe" } } });
        deck.addSlide({ id: "s" }).addElement({ id: "t", type: "text", box: { x: 60, y: 100, width: 800, height: 80 }, frame: { margin: 0, wrap: false }, content: [{
          style: { bullet, indent: item.indent, hanging: item.hanging, tabStops: [{ position: 280, align: item.align }] },
          runs: [{ text: "P\t", style: { fontFamily: "List Tab Probe", fontSize: 28 } }, { text: item.text, style: { fontFamily: "List Tab Probe", fontSize: 28 } }],
        }] });
        const analysis = resolveLayout(deck, { textMeasurement: adapter }).analysis;
        return { diagnostics: analysis.diagnostics, bounds: analysis.measurements[0]?.bounds };
      });
    } finally { adapter.dispose(); }
  }, { bytes: fontBytes.toString("base64"), items });
  const cdp = await page.context().newCDPSession(page); await cdp.send("DOM.enable"); await cdp.send("CSS.enable"); const { root } = await cdp.send("DOM.getDocument");
  const { nodeId } = await cdp.send("DOM.querySelector", { nodeId: root.nodeId, selector: "#font-control" });
  selectedFonts = (await cdp.send("CSS.getPlatformFontsForNode", { nodeId })).fonts;
} finally { await browser.close(); }
assert.equal(selectedFonts.length, 1); assert.equal(selectedFonts[0].postScriptName, "LiberationSans"); assert.ok(selectedFonts[0].isCustomFont && selectedFonts[0].glyphCount > 0);
for (const [index, item] of items.entries()) {
  assert.deepEqual(observed[index].diagnostics, []); item.browserRight = observed[index].bounds.x + observed[index].bounds.width;
  item.rightDifference = item.browserRight - item.pdfRight; assert.ok(Math.abs(item.rightDifference) < 0.5, JSON.stringify(item));
}
const result = { scope: "36 generated PPTX list/tab controls, LibreOffice PDF versus Chromium Layout; marker positions are not directly compared, and no PowerPoint/edit claim is made.",
  source: sourcePath, sourceSha256: sha(sourceBytes), runnerSha256: sha(await readFile("scripts/inspect-tab-list-cross-app.mjs")), browser: browserVersion,
  font: { file: fontFile, sha256: sha(fontBytes) }, selectedFonts, items };
await writeFile("docs/testing/tab-list-browser-libreoffice-evidence.json", JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ comparisons: items.length, maxRightDifference: Math.max(...items.map(item => Math.abs(item.rightDifference))) }));
