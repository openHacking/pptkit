import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { promisify } from "node:util";
import { chromium } from "@playwright/test";

const sha = bytes => createHash("sha256").update(bytes).digest("hex");
await promisify(execFile)("python3", ["scripts/research/verify-shape-mixed-size-powerpoint.py"]);
const sourcePath = "docs/testing/tab-shape-mixed-size-powerpoint-evidence.json", sourceBytes = await readFile(sourcePath), source = JSON.parse(sourceBytes);
const inputs = source.pages.flatMap(page => page.rows.map((row, i) => ({ align: page.kind, word: row.word, suffixSize: row.suffixSize, bold: i === 2, characters: row.characters })));
const fontFiles = [["Regular", "/System/Library/Fonts/Supplemental/Arial.ttf"], ["Bold", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]];
const fonts = await Promise.all(fontFiles.map(async ([weight, file]) => { const data = await readFile(file); return { weight, file, sha256: sha(data), data: data.toString("base64") }; }));
const bundleFiles = ["core", "layout", "svg-renderer"].map(name => [`@pptkit/${name}`, `packages/${name}/dist/browser/index.js`]);
const bundles = await Promise.all(bundleFiles.map(async ([name, file]) => { const data = await readFile(file); return { name, file, sha256: sha(data), url: `data:text/javascript;base64,${data.toString("base64")}` }; }));
const browser = await chromium.launch(); let observations, selectedFonts, version;
try {
  version = browser.version(); const page = await browser.newPage(); await page.route("**/*", route => route.abort());
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: Object.fromEntries(bundles.map(b => [b.name, b.url])) })}</script>`);
  observations = await page.evaluate(async ({ inputs, fonts }) => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    for (const font of fonts) { const face = new FontFace("SVG Mixed Probe", `url(data:font/ttf;base64,${font.data})`, { weight: font.weight === "Bold" ? "700" : "400" }); document.fonts.add(await face.load()); }
    for (const [id, bold] of [["regular-face", false], ["bold-face", true]]) { const e = document.createElement("span"); e.id = id; e.style.cssText = `font-family:"SVG Mixed Probe";font-size:28px;font-weight:${bold ? 700 : 400};font-synthesis:none`; e.textContent = "MIXED"; document.body.append(e); }
    return Promise.all(inputs.map(async item => {
      const deck = createPresentation({ size: { width: 960, height: 540 } });
      deck.addSlide().addElement({ id: "shape", type: "shape", shape: "rect", box: { x: 60, y: 120, width: 700, height: 65 }, style: { fill: { type: "solid", color: "F2F5FA" }, stroke: { paint: { type: "none" } } }, text: { frame: { margin: 0, wrap: false }, content: [{ style: { tabStops: [{ position: 280, align: item.align }] }, runs: [{ text: "P\t", style: { fontFamily: "SVG Mixed Probe", fontSize: 20 } }, { text: item.word, style: { fontFamily: "SVG Mixed Probe", fontSize: item.suffixSize, bold: item.bold } }] }] } });
      const rendered = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
      try {
        const parts = [...holder.querySelectorAll("text[data-pptkit-tab-part]")], suffix = parts.filter(n => n.dataset.pptkitTabPart !== "prefix"), chars = [];
        for (const element of suffix) for (let i=0; i<element.getNumberOfChars(); i++) { const start = element.getStartPositionOfChar(i), end = element.getEndPositionOfChar(i); chars.push({ char: element.textContent[i], left: start.x - 60, right: end.x - 60, baseline: start.y - 120 }); }
        return { chars, text: parts.map(n => n.textContent).join(""), fill: holder.querySelector('[data-pptkit-element-id="shape"] rect')?.getAttribute("fill"), warnings: rendered.warnings.map(w => w.code) };
      } finally { holder.remove(); }
    }));
  }, { inputs, fonts });
  const cdp = await page.context().newCDPSession(page); await cdp.send("DOM.enable"); await cdp.send("CSS.enable"); const { root } = await cdp.send("DOM.getDocument"); selectedFonts = [];
  for (const id of ["regular-face", "bold-face"]) { const { nodeId } = await cdp.send("DOM.querySelector", { nodeId: root.nodeId, selector: `#${id}` }); selectedFonts.push((await cdp.send("CSS.getPlatformFontsForNode", { nodeId })).fonts); }
} finally { await browser.close(); }
assert.deepEqual(selectedFonts.map(row => row.map(font => font.postScriptName)), [["ArialMT"], ["Arial-BoldMT"]]);
const items = inputs.map((input, index) => {
  const observed = observations[index]; assert.equal(observed.text, `P\t${input.word}`); assert.equal(observed.fill, "#F2F5FA"); assert.ok(observed.warnings.includes("text-tab-layout-unsupported"));
  assert.deepEqual(observed.chars.map(c => c.char), input.characters.map(c => c.char));
  return { align: input.align, word: input.word, suffixSize: input.suffixSize, bold: input.bold, chars: input.characters.map((native, i) => ({ char: native.char, native, svg: observed.chars[i], leftDifference: observed.chars[i].left - native.left, rightDifference: observed.chars[i].right - native.right, baselineDifference: observed.chars[i].baseline - native.baseline })) };
});
const chars = items.flatMap(item => item.chars), maxHorizontal = Math.max(...chars.flatMap(c => [Math.abs(c.leftDifference), Math.abs(c.rightDifference)])), maxBaseline = Math.max(...chars.map(c => Math.abs(c.baselineDifference)));
const result = { scope: "First-open PowerPoint PDF versus Chromium SVG for suffix characters in 12 mixed-size Arial shape samples; compares per-character left/right and baseline, not native editing, reopening, other fonts or transforms.", source: sourcePath, sourceSha256: sha(sourceBytes), runnerSha256: sha(await readFile("scripts/inspect-svg-shape-mixed-size-powerpoint.mjs")), browser: version, bundles: bundles.map(({ url, ...item }) => item), fonts: fonts.map(({ data, ...item }) => item), selectedFonts, items, maxHorizontal, maxBaseline };
await writeFile("docs/testing/tab-svg-shape-mixed-size-powerpoint-evidence.json", JSON.stringify(result, null, 2) + "\n");
console.log(JSON.stringify({ samples: items.length, characters: chars.length, maxHorizontal, maxBaseline }));
