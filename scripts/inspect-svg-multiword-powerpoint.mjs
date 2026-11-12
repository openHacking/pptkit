import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { promisify } from "node:util";
import { chromium } from "@playwright/test";

const sha = data => createHash("sha256").update(data).digest("hex");
await promisify(execFile)("python3", ["scripts/research/verify-native-tab-phrases.py"]);
const sourcePath = "docs/testing/tab-phrases-powerpoint-evidence.json", sourceBytes = await readFile(sourcePath), source = JSON.parse(sourceBytes);
const selected = source.observations.find(page => page.width === 220 && page.align === "center"), row = selected.rows[0];
assert.equal(row.length, 1);
const originalChars = row[0].chars, origin = originalChars[0].bbox[0], scale = (selected.guide - origin) / 120;
assert.equal(originalChars.map(char => char.char).join(""), "PIWIDEWORD");
const fonts = await Promise.all([["Regular", "/System/Library/Fonts/Supplemental/Arial.ttf"], ["Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]].map(async ([weight, file]) => { const bytes = await readFile(file); return { weight, file, sha256: sha(bytes), data: bytes.toString("base64") }; }));
const bundles = await Promise.all([["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/svg-renderer", "packages/svg-renderer/dist/browser/index.js"]].map(async ([name, file]) => { const bytes = await readFile(file); return { name, file, sha256: sha(bytes), url: `data:text/javascript;base64,${bytes.toString("base64")}` }; }));
const browser = await chromium.launch(); let measured, selectedFont, version;
try {
  version = browser.version(); const page = await browser.newPage(); await page.route("**/*", route => route.abort());
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: Object.fromEntries(bundles.map(item => [item.name, item.url])) })}</script>`);
  measured = await page.evaluate(async ({ fonts }) => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    for (const font of fonts) document.fonts.add(await new FontFace("Phrase Probe", `url(data:font/ttf;base64,${font.data})`, { weight: font.weight === "Bold" ? "700" : "400" }).load());
    const probe = document.createElement("span"); probe.id = "face-probe"; probe.style.cssText = 'font-family:"Phrase Probe";font-size:28px;font-weight:400;font-synthesis:none'; probe.textContent = "I WIDE WORD"; document.body.append(probe);
    const deck = createPresentation({ size: { width: 960, height: 540 }, theme: { fonts: { body: "Phrase Probe", heading: "Phrase Probe" } } });
    deck.addSlide().addElement({ type: "text", box: { x: 60, y: 100, width: 220, height: 180 }, frame: { margin: 0, wrap: false }, content: [{ style: { lineSpacing: 1.2, tabStops: [{ position: 120, align: "center" }] }, runs: [{ text: "P\t", style: { fontSize: 28, fontFamily: "Phrase Probe" } }, { text: "I WIDE WORD", style: { fontSize: 28, fontFamily: "Phrase Probe" } }] }] });
    const rendered = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
    try {
      const nodes = [...holder.querySelectorAll("text[data-pptkit-tab-part]")];
      return { warnings: rendered.warnings.map(w => w.code), parts: nodes.map(node => node.textContent), chars: nodes.flatMap(node => [...node.textContent].flatMap((char, index) => /\s/u.test(char) ? [] : [{ char, x: node.getStartPositionOfChar(index).x - 60, right: node.getEndPositionOfChar(index).x - 60 }])) };
    } finally { holder.remove(); }
  }, { fonts });
  const cdp = await page.context().newCDPSession(page); await cdp.send("DOM.enable"); await cdp.send("CSS.enable"); const { root } = await cdp.send("DOM.getDocument"), { nodeId } = await cdp.send("DOM.querySelector", { nodeId: root.nodeId, selector: "#face-probe" }); selectedFont = (await cdp.send("CSS.getPlatformFontsForNode", { nodeId })).fonts;
} finally { await browser.close(); }
assert.deepEqual(selectedFont.map(font => font.postScriptName), ["ArialMT"]);
assert.deepEqual(measured.parts, ["P\t", "I WIDE WORD"]); assert.ok(measured.warnings.includes("text-tab-layout-unsupported"));
const chars = originalChars.map((char, index) => {
  const svg = measured.chars[index]; assert.equal(svg.char, char.char);
  const pdfX = (char.bbox[0] - origin) / scale, pdfRight = (char.bbox[2] - origin) / scale;
  return { char: char.char, pdfX, pdfRight, svgX: svg.x, svgRight: svg.right, xDifference: svg.x - pdfX, rightDifference: svg.right - pdfRight };
});
assert.equal(chars.length, measured.chars.length);
assert.ok(chars.every(char => Math.abs(char.xDifference) < 0.5 && Math.abs(char.rightDifference) < 0.5), JSON.stringify(chars));
const result = { scope: "One non-overlapping centered multiword sample: fitting PowerPoint wrapped frame versus no-wrap SVG text in Chromium. No native no-wrap, shape, edit/reopen, overlap or general typography claim.", source: sourcePath, sourceSha256: sha(sourceBytes), runnerSha256: sha(await readFile("scripts/inspect-svg-multiword-powerpoint.mjs")), browser: version, bundles: bundles.map(({ url: _url, ...item }) => item), fonts: fonts.map(({ data: _data, ...item }) => item), selectedFont, parts: measured.parts, warnings: measured.warnings, chars };
await writeFile("docs/testing/tab-svg-multiword-browser-powerpoint-evidence.json", JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ characters: chars.length, maxDifference: Math.max(...chars.flatMap(char => [Math.abs(char.xDifference), Math.abs(char.rightDifference)])) }));
