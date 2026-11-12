// Optional live comparison against the retained LibreOffice PDF observations.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import { chromium } from "@playwright/test";

const run = promisify(execFile), sha = data => createHash("sha256").update(data).digest("hex");
await run("python3", ["scripts/research/verify-tab-stop-wrap.py"]);
const sourcePath = "docs/testing/tab-stop-wrap-libreoffice-evidence.json", sourceBytes = await readFile(sourcePath), source = JSON.parse(sourceBytes);
const wrapper = (await run("which", ["soffice"])).stdout.trim();
const fontDir = path.resolve(path.dirname(wrapper), "../../native/libreoffice-headless/libreoffice/LibreOfficeDev.app/Contents/Resources/fonts/truetype");
const fonts = await Promise.all(["Regular", "Bold"].map(async weight => {
  const file = path.join(fontDir, `LiberationSans-${weight}.ttf`), bytes = await readFile(file);
  return { weight, file, sha256: sha(bytes), bytes: bytes.toString("base64") };
}));
const items = source.cases.filter(item => !item.plain).flatMap(caseInfo => caseInfo.pages.flatMap((page, pageIndex) => [0, 1, 2].map(sampleIndex => {
  const prefix = page[2 * sampleIndex], sample = page[2 * sampleIndex + 1];
  return { case: caseInfo.name, width: caseInfo.width, wrap: caseInfo.wrap === "wrap", align: ["left", "center", "right", "decimal"][pageIndex], text: sample.text, bold: sampleIndex === 2,
    pdfRight: Math.max(prefix.bbox[2], sample.bbox[2]) - 60, pdfWrapped: sample.bbox[1] > prefix.bbox[1] + 10 };
})));
const imports = Object.fromEntries(await Promise.all([["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/layout/dom", "packages/layout/dist/browser/dom.js"]].map(async ([name, file]) => [name, `data:text/javascript;base64,${(await readFile(file)).toString("base64")}`])));
const browser = await chromium.launch(); let observations, selectedFonts, browserVersion;
try {
  browserVersion = browser.version(); const page = await browser.newPage(); await page.route("**/*", route => route.abort());
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  observations = await page.evaluate(async ({ fonts, items }) => {
    const { createPresentation } = await import("@pptkit/core"), { resolveLayout } = await import("@pptkit/layout"), { createBrowserTextMeasurement } = await import("@pptkit/layout/dom");
    const faces = await Promise.all(fonts.map(async font => new FontFace("Tab Probe", `url(data:font/ttf;base64,${font.bytes})`, { weight: font.weight === "Bold" ? "700" : "400" }).load()));
    for (const face of faces) document.fonts.add(face);
    const adapter = createBrowserTextMeasurement({ document, fonts: [{ family: "Tab Probe", faces, covers: ({ text, italic }) => !italic && /^[\x00-\x7f]*$/u.test(text) }] });
    for (const [id, bold] of [["regular-face", false], ["bold-face", true]]) { const span = document.createElement("span"); span.id = id; span.style.cssText = `font-family:"Tab Probe";font-size:28px;font-weight:${bold ? 700 : 400};font-synthesis:none`; span.textContent = "MIXED"; document.body.appendChild(span); }
    try { return items.map(item => {
      const deck = createPresentation({ theme: { fonts: { body: "Tab Probe", heading: "Tab Probe" } } });
      deck.addSlide({ id: "s" }).addElement({ id: "t", type: "text", box: { x: 60, y: 120, width: item.width, height: 200 }, frame: { margin: 0, wrap: item.wrap }, content: [{ style: { tabStops: [{ position: 120, align: item.align }] }, runs: [{ text: "P\t", style: { fontFamily: "Tab Probe", fontSize: 28 } }, { text: item.text, style: { fontFamily: "Tab Probe", fontSize: 28, bold: item.bold } }] }] });
      const analysis = resolveLayout(deck, { textMeasurement: adapter }).analysis; return { diagnostics: analysis.diagnostics, bounds: analysis.measurements[0]?.bounds };
    }); } finally { adapter.dispose(); }
  }, { fonts, items });
  const cdp = await page.context().newCDPSession(page); await cdp.send("DOM.enable"); await cdp.send("CSS.enable"); const { root } = await cdp.send("DOM.getDocument"); selectedFonts = [];
  for (const id of ["regular-face", "bold-face"]) { const { nodeId } = await cdp.send("DOM.querySelector", { nodeId: root.nodeId, selector: `#${id}` }); selectedFonts.push((await cdp.send("CSS.getPlatformFontsForNode", { nodeId })).fonts); }
} finally { await browser.close(); }
assert.deepEqual(selectedFonts.map(row => row.map(font => font.postScriptName)), [["LiberationSans"], ["LiberationSans-Bold"]]);
for (const [index, item] of items.entries()) {
  const observation = observations[index]; assert.deepEqual(observation.diagnostics, []);
  item.browserRight = observation.bounds.width + observation.bounds.x; item.browserWrapped = observation.bounds.height > 40;
  item.rightDifference = item.browserRight - item.pdfRight;
  assert.equal(item.browserWrapped, item.pdfWrapped, JSON.stringify(item)); assert.ok(Math.abs(item.rightDifference) < 0.5, JSON.stringify(item));
}
const result = { scope: "Synthetic ODF controls, LibreOffice PDF versus live Chromium Layout using named font candidates; no PowerPoint or native-edit claim.", source: sourcePath, sourceSha256: sha(sourceBytes), runnerSha256: sha(await readFile("scripts/inspect-tab-wrap-cross-app.mjs")), browser: browserVersion, fonts: fonts.map(({ bytes: _bytes, ...font }) => font), selectedFonts, items };
await writeFile("docs/testing/tab-stop-wrap-browser-libreoffice-evidence.json", JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ comparisons: items.length, maxRightDifference: Math.max(...items.map(item => Math.abs(item.rightDifference))) }));
