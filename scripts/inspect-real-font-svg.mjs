// Optional exact-face SVG experiment. Requires the local shaping runner's artifacts.
import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { platform, release, arch } from "node:os";
import { chromium } from "@playwright/test";

if (process.argv.length > 3) throw Error("Usage: node scripts/inspect-real-font-svg.mjs [shaping-evidence.json]");
const evidenceFile = process.argv[2] ?? "docs/testing/real-font-shaping-evidence.json";
const source = JSON.parse(await readFile(evidenceFile, "utf8")), hash = data => createHash("sha256").update(data).digest("hex");
const fonts = [];
for (const font of source.fonts) {
  const bytes = await readFile(path.join(source.artifactsDirectory, `${font.id}.ttf`));
  assert.equal(hash(bytes), font.extractedSha256);
  assert.equal(hash(await readFile(font.file)), font.sourceSha256);
  fonts.push({ ...font, alias: `PPTKitSVG ${font.id} ${font.extractedSha256.slice(0, 12)}`, bytes: bytes.toString("base64") });
}
await mkdir(".tmp/real-font-svg", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/real-font-svg/run-"));
const moduleUrl = async file => `data:text/javascript;base64,${(await readFile(file)).toString("base64")}`;
const imports = Object.fromEntries(await Promise.all([
  ["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/svg-renderer", "packages/svg-renderer/dist/browser/index.js"],
].map(async ([name, file]) => [name, await moduleUrl(file)])));
const browser = await chromium.launch();
let result, browserVersion;
try {
  browserVersion = browser.version();
  const page = await browser.newPage({ viewport: { width: 1160, height: 1000 }, deviceScaleFactor: 1 });
  const requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script><style>html,body{margin:0;background:white;color:black;font:14px sans-serif}.row{display:grid;grid-template-columns:560px 560px;gap:12px;margin-bottom:8px}.label{height:20px;line-height:20px}svg,img{display:block;width:550px;height:120px}</style>`);
  result = await page.evaluate(async ({ fonts, specimens }) => {
    const { createPresentation } = await import("@pptkit/core");
    const { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    for (const font of fonts) { const face = await new FontFace(font.alias, `url(data:font/ttf;base64,${font.bytes})`).load(); document.fonts.add(face); }
    const latin = fonts.find(font => font.id === "latin");
    const cases = [];
    let unembeddedControl;
    for (const [sampleIndex, specimen] of specimens.entries()) for (const mode of ["native", "mixed"]) {
      const font = fonts.find(font => font.id === specimen.fontId);
      const deck = createPresentation({ size: { width: 550, height: 120 }, theme: { fonts: { body: font.alias, heading: font.alias } } });
      const style = { fontFamily: font.alias, fontSize: 48, language: specimen.language, color: "000000" };
      const runs = [...(mode === "mixed" ? [{ text: "A ", style: { fontFamily: latin.alias, fontSize: 14, color: "000000" } }] : []), { text: specimen.text, style }];
      deck.addSlide({ id: `s${sampleIndex}${mode}` }).addElement({ id: "t", type: "text", box: { x: 20, y: 20, width: 500, height: 90 }, frame: { margin: 0, wrap: false }, content: [{ style: { lineSpacing: 1.5 }, runs }] });
      const before = JSON.stringify(deck.slides[0].elements), rendered = await renderPresentationToSvg(deck);
      if (before !== JSON.stringify(deck.slides[0].elements)) throw Error("Rendering mutated author runs");
      const index = cases.length, wrapper = document.createElement("div"); wrapper.className = "row";
      const inline = document.createElement("div"); inline.id = `case-${index}`;
      const label = document.createElement("div"); label.className = "label"; label.textContent = `${specimen.family} / ${specimen.language} / ${mode} / inline`;
      inline.append(label); inline.insertAdjacentHTML("beforeend", rendered.slides[0].svg); wrapper.append(inline); document.body.append(wrapper);
      const selected = inline.querySelector(mode === "native" ? "tspan" : "foreignObject span[lang]:last-of-type");
      if (!selected || selected.textContent !== specimen.text) throw Error("Actual SVG text path changed source text");
      selected.id = `probe-${index}`;
      let advance;
      if (mode === "native") advance = selected.getComputedTextLength();
      else { const range = document.createRange(); range.selectNodeContents(selected); advance = range.getBoundingClientRect().width; }
      const embedded = [font, ...(mode === "mixed" && font.id !== latin.id ? [latin] : [])].map(font => `@font-face{font-family:"${font.alias}";src:url(data:font/ttf;base64,${font.bytes});font-weight:400;font-style:normal}`).join("");
      const svg = rendered.slides[0].svg.replace('</defs>', `<style>${embedded}</style></defs>`);
      const raster = document.createElement("div"), imageLabel = label.cloneNode(); imageLabel.textContent = "same SVG as image, embedded exact face"; raster.append(imageLabel);
      const image = document.createElement("img"); image.id = `image-${index}`; image.width = 550; image.height = 120; image.src = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`; await image.decode(); raster.append(image); wrapper.append(raster);
      cases.push({ index, sampleIndex, fontId: specimen.fontId, registeredFamily: font.alias, mode, text: specimen.text, expectedAdvance: specimen.advance, advance, warnings: rendered.warnings, svg: rendered.slides[0].svg });
      if (font.id === "latin" && mode === "native") unembeddedControl = { index, svg: rendered.slides[0].svg };
    }
    if (!unembeddedControl) throw Error("A Latin native specimen is required for the isolated-image font control");
    const label = document.createElement("div"); label.textContent = "Negative control: same alias SVG image without embedded font"; document.body.append(label);
    const unembedded = document.createElement("img"); unembedded.id = "unembedded-control"; unembedded.src = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(unembeddedControl.svg)))}`; await unembedded.decode(); document.body.append(unembedded);
    const control = createPresentation(); control.addSlide().addElement({ type: "text", id: "nowrap", box: { x: 0, y: 0, width: 35, height: 200 }, frame: { margin: 0, wrap: false }, content: "AAAA AAAA AAAA" });
    const controlSvg = (await renderPresentationToSvg(control)).slides[0].svg;
    return { cases, unembeddedControlCase: unembeddedControl.index, noWrap: { nativeLineCount: (controlSvg.match(/data-pptkit-text-line=/g) ?? []).length, svg: controlSvg } };
  }, { fonts, specimens: source.comparisons });
  const session = await page.context().newCDPSession(page); await session.send("DOM.enable"); await session.send("CSS.enable"); const { root } = await session.send("DOM.getDocument");
  for (const control of result.cases) {
    const { nodeId } = await session.send("DOM.querySelector", { nodeId: root.nodeId, selector: `#probe-${control.index}` });
    control.platformFonts = (await session.send("CSS.getPlatformFontsForNode", { nodeId })).fonts;
    const expectedFont = fonts.find(font => font.id === control.fontId);
    control.exactFace = control.platformFonts.length === 1 && control.platformFonts[0].postScriptName === expectedFont.postscriptName && control.platformFonts[0].isCustomFont && control.platformFonts[0].glyphCount > 0;
    assert(control.exactFace, `Actual SVG face differs: ${control.fontId}/${control.mode}`);
    assert(Math.abs(control.advance - control.expectedAdvance) < .05, `SVG shaping advance differs: ${control.fontId}/${control.mode}`);
    assert.deepEqual(control.warnings, []);
    await writeFile(path.join(directory, `case-${control.index}.svg`), control.svg);
    await page.locator(`#case-${control.index} svg`).screenshot({ path: path.join(directory, `inline-${control.index}.png`) });
    await page.locator(`#image-${control.index}`).screenshot({ path: path.join(directory, `image-${control.index}.png`) });
    control.svgSha256 = hash(control.svg); delete control.svg;
  }
  await page.locator("#unembedded-control").screenshot({ path: path.join(directory, "unembedded.png") });
  await page.screenshot({ path: path.join(directory, "contact.png"), fullPage: true });
  result.networkRequests = requests;
  assert.equal(result.noWrap.nativeLineCount, 1, "Native SVG ignored wrap:false");
  assert.deepEqual(requests, []);
} finally { await browser.close(); }
const raster = JSON.parse((await promisify(execFile)("python3", ["scripts/research/oracles/real-font-svg.py", directory, String(result.cases.length), String(result.unembeddedControlCase)], { timeout: 30000, maxBuffer: 1024 * 1024 })).stdout);
const evidence = { schemaVersion: 1, recordedAt: new Date().toISOString(), status: "observed", sourceEvidence: evidenceFile, sourceEvidenceSha256: hash(await readFile(evidenceFile)), artifactsDirectory: path.relative(process.cwd(), directory), chromium: browserVersion, platform: `${platform()} ${release()} ${arch()}`, fonts: fonts.map(({ bytes, ranges, ...font }) => font), ...result, raster, limitations: ["Named regular faces and isolated strings only", "No mixed bidi, emoji or native Office layout admission", "SVG image and foreignObject portability evaluated separately", "Font embedding is experiment-owned, not a Renderer API", "No font binaries redistributed"] };
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
await writeFile("docs/testing/real-font-svg-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, cases: result.cases.length, noWrapNativeLines: result.noWrap.nativeLineCount, raster }));
