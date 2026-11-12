// Optional local experiment: exact Liberation faces used by the retained LibreOffice PDFs.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import { chromium } from "@playwright/test";

const run = promisify(execFile), sha = data => createHash("sha256").update(data).digest("hex");
await run("python3", ["scripts/research/verify-tab-stop-paragraph.py"]);
const sourcePath = "docs/testing/tab-stop-paragraph-libreoffice-evidence.json", sourceBytes = await readFile(sourcePath), source = JSON.parse(sourceBytes);
const wrapper = (await run("which", ["soffice"])).stdout.trim();
const fontDir = path.resolve(path.dirname(wrapper), "../../native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype");
const fonts = await Promise.all(["Regular", "Bold"].map(async weight => {
  const file = path.join(fontDir, `LiberationSans-${weight}.ttf`), bytes = await readFile(file);
  return { weight, file, sha256: sha(bytes), bytes: bytes.toString("base64") };
}));
const items = source.cases.flatMap(caseInfo => caseInfo.pages.flatMap((page, pageIndex) => page.samples.map((sample, sampleIndex) => ({
  case: caseInfo.name, indent: caseInfo.margin, hanging: -caseInfo.indent, align: ["left", "center", "right", "decimal"][pageIndex], text: sample.text,
  bold: sampleIndex === 2, pdfWidth: sample.bbox[2] - sample.bbox[0], pdfRight: sample.bbox[2] - 60,
}))));
for (const item of items) {
  const font = fonts[item.bold ? 1 : 0];
  const glyphs = JSON.parse((await run("hb-shape", [font.file, item.text, "--output-format=json", "--no-glyph-names", "--font-size=28000", "--direction=ltr", "--language=en"], { maxBuffer: 1024 * 1024 })).stdout);
  assert.ok(glyphs.every(g => g.g !== 0)); item.harfBuzzWidth = glyphs.reduce((sum, glyph) => sum + glyph.ax, 0) / 1000;
  assert.ok(Math.abs(item.harfBuzzWidth - item.pdfWidth) < 0.5, `PDF/font advance mismatch ${JSON.stringify(item)}`);
}
const moduleUrl = async file => `data:text/javascript;base64,${(await readFile(file)).toString("base64")}`;
const imports = Object.fromEntries(await Promise.all([
  ["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/layout/dom", "packages/layout/dist/browser/dom.js"],
].map(async ([name, file]) => [name, await moduleUrl(file)])));
const browser = await chromium.launch(); let observed, selectedFonts, browserVersion;
try {
  browserVersion = browser.version(); const page = await browser.newPage(); await page.route("**/*", route => route.abort());
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  observed = await page.evaluate(async ({ fonts, items }) => {
    const { createPresentation } = await import("@pptkit/core"), { resolveLayout } = await import("@pptkit/layout"), { createBrowserTextMeasurement } = await import("@pptkit/layout/dom");
    const faces = await Promise.all(fonts.map(async font => new FontFace("Tab Probe", `url(data:font/ttf;base64,${font.bytes})`, { weight: font.weight === "Bold" ? "700" : "400" }).load()));
    for (const face of faces) document.fonts.add(face);
    const adapter = createBrowserTextMeasurement({ document, fonts: [{ family: "Tab Probe", faces, covers: ({ text, italic }) => !italic && /^[\x00-\x7f]*$/u.test(text) }] });
    for (const [id, bold] of [["regular-face", false], ["bold-face", true]]) {
      const span = document.createElement("span"); span.id = id; span.style.cssText = `font-family:"Tab Probe";font-size:28px;font-weight:${bold ? 700 : 400};font-synthesis:none`; span.textContent = "MIXED"; document.body.appendChild(span);
    }
    try {
      return items.map(item => {
        const deck = createPresentation({ theme: { fonts: { body: "Tab Probe", heading: "Tab Probe" } } });
        deck.addSlide({ id: "s" }).addElement({ id: "t", type: "text", box: { x: 60, y: 120, width: 700, height: 65 }, frame: { margin: 0, wrap: false }, content: [{
          style: { indent: item.indent, hanging: item.hanging, tabStops: [{ position: 280, align: item.align }] },
          runs: [{ text: "P\t", style: { fontFamily: "Tab Probe", fontSize: 28 } }, { text: item.text, style: { fontFamily: "Tab Probe", fontSize: 28, bold: item.bold } }],
        }] });
        const analysis = resolveLayout(deck, { textMeasurement: adapter }).analysis;
        return { diagnostics: analysis.diagnostics, bounds: analysis.measurements[0]?.bounds };
      });
    } finally { adapter.dispose(); }
  }, { fonts, items });
  const cdp = await page.context().newCDPSession(page); await cdp.send("DOM.enable"); await cdp.send("CSS.enable"); const { root } = await cdp.send("DOM.getDocument");
  selectedFonts = [];
  for (const id of ["regular-face", "bold-face"]) {
    const { nodeId } = await cdp.send("DOM.querySelector", { nodeId: root.nodeId, selector: `#${id}` });
    selectedFonts.push((await cdp.send("CSS.getPlatformFontsForNode", { nodeId })).fonts);
  }
} finally { await browser.close(); }
assert.deepEqual(selectedFonts.map(row => row.map(font => font.postScriptName)), [["LiberationSans"], ["LiberationSans-Bold"]]);
assert.ok(selectedFonts.every(row => row.length === 1 && row[0].isCustomFont && row[0].glyphCount > 0));
for (const [index, item] of items.entries()) {
  assert.deepEqual(observed[index].diagnostics, []); item.browserRight = observed[index].bounds.width + observed[index].bounds.x;
  item.rightDifference = item.browserRight - item.pdfRight;
  assert.ok(Math.abs(item.rightDifference) < 0.5, `PDF/browser position mismatch ${JSON.stringify(item)}`);
}
const result = { scope: "Four synthetic ODF paragraph controls, three samples, four alignments; LibreOffice PDF versus Chromium Layout using named matching font candidates. No PowerPoint or native edit claim.",
  source: sourcePath, sourceSha256: sha(sourceBytes), runnerSha256: sha(await readFile("scripts/inspect-tab-stops-cross-app.mjs")), browser: browserVersion,
  fonts: fonts.map(({ bytes: _bytes, ...font }) => font), selectedFonts, items };
await writeFile("docs/testing/tab-stop-browser-libreoffice-evidence.json", JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ comparisons: items.length, maxRightDifference: Math.max(...items.map(item => Math.abs(item.rightDifference))), browser: browserVersion }));
