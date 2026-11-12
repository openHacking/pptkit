import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { promisify } from "node:util";
import { chromium } from "@playwright/test";

const run = promisify(execFile), sha = bytes => createHash("sha256").update(bytes).digest("hex");
await run("python3", ["scripts/research/verify-native-tab-stops.py"]);
const sourcePath = "docs/testing/tab-stop-powerpoint-evidence.json", sourceBytes = await readFile(sourcePath), source = JSON.parse(sourceBytes);
const fonts = await Promise.all([["Regular", "/System/Library/Fonts/Supplemental/Arial.ttf"], ["Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]].map(async ([weight, file]) => { const data = await readFile(file); return { weight, file, sha256: sha(data), bytes: data.toString("base64") }; }));
const items = source.slides.flatMap(page => page.rows.map((row, index) => {
  const origin = row.prefix[0], scale = (page.guide - origin) / 280;
  return { align: page.slide, text: row.text, bold: index === 2, pdfLeft: (row.word[0] - origin) / scale, pdfRight: (row.word[2] - origin) / scale };
}));
const bundleFiles = [["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/svg-renderer", "packages/svg-renderer/dist/browser/index.js"]];
const bundles = await Promise.all(bundleFiles.map(async ([name, file]) => { const data = await readFile(file); return { name, file, sha256: sha(data), url: `data:text/javascript;base64,${data.toString("base64")}` }; }));
const imports = Object.fromEntries(bundles.map(bundle => [bundle.name, bundle.url]));
const browser = await chromium.launch(); let observations, selectedFonts, version;
try {
  version = browser.version(); const page = await browser.newPage(); await page.route("**/*", route => route.abort()); await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  observations = await page.evaluate(async ({ items, fonts }) => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    for (const font of fonts) { const face = await new FontFace("SVG Tab Probe", `url(data:font/ttf;base64,${font.bytes})`, { weight: font.weight === "Bold" ? "700" : "400" }).load(); document.fonts.add(face); }
    for (const [id, bold] of [["regular-face", false], ["bold-face", true]]) { const span = document.createElement("span"); span.id = id; span.style.cssText = `font-family:"SVG Tab Probe";font-size:28px;font-weight:${bold ? 700 : 400};font-synthesis:none`; span.textContent = "MIXED"; document.body.appendChild(span); }
    return Promise.all(items.map(async item => {
      const deck = createPresentation({ size: { width: 960, height: 540 }, theme: { fonts: { body: "SVG Tab Probe", heading: "SVG Tab Probe" } } });
      deck.addSlide({ id: "s" }).addElement({ id: "t", type: "text", box: { x: 60, y: 120, width: 700, height: 65 }, frame: { margin: 0, wrap: false }, content: [{ style: { tabStops: [{ position: 280, align: item.align }] }, runs: [{ text: `P\t${item.text}`, style: { fontFamily: "SVG Tab Probe", fontSize: 28, bold: item.bold } }] }] });
      const result = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = result.slides[0].svg; document.body.appendChild(holder);
      try {
        const parts = [...holder.querySelectorAll("text[data-pptkit-tab-part]")], suffix = parts.filter(node => node.dataset.pptkitTabPart !== "prefix");
        return { warnings: result.warnings.map(w => w.code), parts: parts.map(node => [node.dataset.pptkitTabPart, node.textContent]), left: suffix[0]?.getStartPositionOfChar(0).x - 60, right: suffix.at(-1)?.getEndPositionOfChar(suffix.at(-1).getNumberOfChars() - 1).x - 60 };
      } finally { holder.remove(); }
    }));
  }, { items, fonts });
  const cdp = await page.context().newCDPSession(page); await cdp.send("DOM.enable"); await cdp.send("CSS.enable"); const { root } = await cdp.send("DOM.getDocument"); selectedFonts = [];
  for (const id of ["regular-face", "bold-face"]) { const { nodeId } = await cdp.send("DOM.querySelector", { nodeId: root.nodeId, selector: `#${id}` }); selectedFonts.push((await cdp.send("CSS.getPlatformFontsForNode", { nodeId })).fonts); }
} finally { await browser.close(); }
assert.deepEqual(selectedFonts.map(row => row.map(font => font.postScriptName)), [["ArialMT"], ["Arial-BoldMT"]]);
for (const [index, item] of items.entries()) {
  const result = observations[index]; assert.ok(result.warnings.includes("text-tab-layout-unsupported"));
  assert.equal(result.parts.map(part => part[1]).join(""), `P\t${item.text}`);
  item.svgLeft = result.left; item.svgRight = result.right; item.leftDifference = item.svgLeft - item.pdfLeft; item.rightDifference = item.svgRight - item.pdfRight;
  assert.ok(Math.abs(item.leftDifference) < 0.5 && Math.abs(item.rightDifference) < 0.5, JSON.stringify(item));
}
const result = { scope: "12 no-wrap suffixes, PowerPoint PDF versus SVG native text anchors in Chromium with named Arial candidates. SVG source uses one run; the PowerPoint source used two and its prefix is regular. No edit/reopen or wrapping claim.", source: sourcePath, sourceSha256: sha(sourceBytes), runnerSha256: sha(await readFile("scripts/inspect-svg-tab-powerpoint.mjs")), browser: version, bundles: bundles.map(({ url: _url, ...bundle }) => bundle), fonts: fonts.map(({ bytes: _bytes, ...font }) => font), selectedFonts, items };
await writeFile("docs/testing/tab-svg-browser-powerpoint-evidence.json", JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ comparisons: items.length, maxDifference: Math.max(...items.flatMap(item => [Math.abs(item.leftDifference), Math.abs(item.rightDifference)])) }));
