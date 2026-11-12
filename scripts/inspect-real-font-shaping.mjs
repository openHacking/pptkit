// Optional installed-font experiment. Extracted proprietary faces stay under .tmp.
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile, writeFile, mkdir, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { chromium } from "@playwright/test";
import { nativeFontFamilies } from "./research/font-native-cases.mjs";

const execute = promisify(execFile), execution = { timeout: 45000, maxBuffer: 16 * 1024 * 1024 };
const hash = data => createHash("sha256").update(data).digest("hex");
await mkdir(".tmp/real-font-shaping", { recursive: true });
const directory = await mkdtemp(path.resolve(".tmp/real-font-shaping/run-"));
const specimens = [
  { id: "sc", text: "骨直令返", language: "zh-CN", script: "Hans", direction: "ltr" },
  { id: "tc", text: "骨直令返", language: "zh-TW", script: "Hant", direction: "ltr" },
  { id: "ja", text: "骨直令返", language: "ja", script: "Jpan", direction: "ltr" },
  { id: "ko", text: "한글", language: "ko", script: "Hang", direction: "ltr" },
  { id: "ar", text: "السَّلَامُ", language: "ar", script: "Arab", direction: "rtl" },
  { id: "ar", text: "لا", language: "ar", script: "Arab", direction: "rtl" },
  { id: "he", text: "שָׁלוֹם", language: "he", script: "Hebr", direction: "rtl" },
  { id: "latin", text: "a\u0301", language: "en", script: "Latn", direction: "ltr" },
];
const shape = async (file, index, specimen) => JSON.parse((await execute("hb-shape", [file, specimen.text, `--face-index=${index}`, "--output-format=json", "--no-glyph-names", "--font-size=48000", `--direction=${specimen.direction}`, `--script=${specimen.script}`, `--language=${specimen.language}`], execution)).stdout);
const fonts = [];
for (const id of [...new Set([...specimens.map(s => s.id), "generic"])]) {
  const family = nativeFontFamilies[id];
  const values = (await execute("fc-match", ["-f", "%{family[0]}\t%{postscriptname}\t%{file}\t%{index}", family], execution)).stdout.split("\t");
  assert.equal(values[0], family, `Installed font required: ${family}`);
  const [, postscriptName, file, faceIndex] = values, index = Number(faceIndex);
  const query = (await execute("fc-query", ["--index", String(index), "-f", "%{fontversion}\t%{charset}", file], execution)).stdout.split("\t");
  const ranges = query[1].trim().split(/\s+/).map(range => range.split("-").map(n => parseInt(n, 16))).map(([start, end = start]) => [start, end]);
  assert(ranges.length > 0 && ranges.every(([start, end]) => Number.isInteger(start) && start <= end));
  const extracted = path.join(directory, `${id}.ttf`);
  await execute("hb-subset", [file, `--face-index=${index}`, "--keep-everything", `--output-file=${extracted}`], execution);
  const bytes = await readFile(extracted);
  fonts.push({ id, family, postscriptName, file, index, version: query[0], sourceSha256: hash(await readFile(file)), extractedSha256: hash(bytes), ranges, bytes: bytes.toString("base64") });
}
const shaping = [];
for (const specimen of specimens) {
  const font = fonts.find(f => f.id === specimen.id);
  const original = await shape(font.file, font.index, specimen), extracted = await shape(path.join(directory, `${specimen.id}.ttf`), 0, specimen);
  assert.deepEqual(extracted, original, `Extraction changed shaping: ${specimen.id}/${specimen.text}`);
  assert(original.every(glyph => glyph.g !== 0), `Missing actual HB glyph: ${specimen.id}`);
  shaping.push({ ...specimen, family: font.family, glyphs: original, advance: original.reduce((n, glyph) => n + glyph.ax, 0) / 1000 });
}
const moduleUrl = async file => `data:text/javascript;base64,${(await readFile(file)).toString("base64")}`;
const imports = Object.fromEntries(await Promise.all([
  ["@pptkit/core", "packages/core/dist/browser/index.js"], ["@pptkit/layout", "packages/layout/dist/browser/index.js"], ["@pptkit/layout/dom", "packages/layout/dist/browser/dom.js"],
].map(async ([name, file]) => [name, await moduleUrl(file)])));
const browser = await chromium.launch();
let browserResult, platformFonts, browserVersion;
try {
  browserVersion = browser.version();
  const page = await browser.newPage({ viewport: { width: 1100, height: 1000 }, deviceScaleFactor: 1 });
  await page.route("**/*", route => route.abort());
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script><style>body{background:white;color:black;font:16px sans-serif}.sample{display:block;white-space:pre;font-size:48px;font-synthesis:none;width:max-content;margin:10px 0 18px}</style>`);
  browserResult = await page.evaluate(async ({ fonts, shaping }) => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { createBrowserTextMeasurement } = await import("@pptkit/layout/dom");
    const inventory = [];
    for (const font of fonts) {
      const face = await new FontFace(font.family, `url(data:font/ttf;base64,${font.bytes})`).load(); document.fonts.add(face);
      inventory.push({ family: font.family, faces: [face], covers: ({ text, bold, italic }) => !bold && !italic && [...text].every(c => font.ranges.some(([start, end]) => start <= c.codePointAt(0) && c.codePointAt(0) <= end)) });
    }
    const adapter = createBrowserTextMeasurement({ document, fonts: inventory });
    const fallback = createBrowserTextMeasurement({ document, fonts: inventory, fallbackFamilies: [fonts.find(f => f.id === "sc").family] });
    const analyze = (family, text, language, provider = adapter, bold = false) => {
      const deck = createPresentation({ theme: { fonts: { body: family, heading: family } } });
      deck.addSlide({ id: "s" }).addElement({ id: "t", type: "text", box: { x: 0, y: 0, width: 1000, height: 150 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text, style: { fontFamily: family, fontSize: 48, language, bold } }] }] });
      const before = JSON.stringify(deck.slides.map(slide => slide.elements));
      const analysis = resolveLayout(deck, { textMeasurement: provider }).analysis;
      if (JSON.stringify(deck.slides.map(slide => slide.elements)) !== before) throw new Error("Measurement changed author runs");
      return analysis;
    };
    try {
      const samples = shaping.map((sample, index) => {
        const label = document.createElement("div");label.textContent = `${sample.family} / ${sample.language}`; document.body.append(label);
        const span = document.createElement("span");span.id = `sample-${index}`;span.className = "sample";span.style.fontFamily = `"${sample.family}"`;span.lang = sample.language;span.dir = sample.direction;span.textContent = sample.text;document.body.append(span);
        return { id: span.id, analysis: analyze(sample.family, sample.text, sample.language), browserAdvance: span.getBoundingClientRect().width };
      });
      const latin = fonts.find(f => f.id === "latin").family;
      return { samples, missingGlyph: analyze(latin, "骨直令返", "zh-TW"), unsupportedStyle: analyze(latin, "ABC", "en", adapter, true), permittedFallback: analyze(latin, "骨直令返", "zh-CN", fallback), emoji: analyze(latin, "👩🏽‍💻", "en"), authorRunsPreserved: true };
    } finally { adapter.dispose(); fallback.dispose(); }
  }, { fonts, shaping });
  const session = await page.context().newCDPSession(page);
  await session.send("DOM.enable"); await session.send("CSS.enable");
  const { root } = await session.send("DOM.getDocument");
  platformFonts = [];
  for (const sample of browserResult.samples) {
    const { nodeId } = await session.send("DOM.querySelector", { nodeId: root.nodeId, selector: `#${sample.id}` });
    platformFonts.push((await session.send("CSS.getPlatformFontsForNode", { nodeId })).fonts);
  }
  await page.screenshot({ path: path.join(directory, "specimens.png"), fullPage: true });
} finally { await browser.close(); }
const diagnostic = (analysis, code) => analysis.diagnostics.some(d => d.code === code);
assert(diagnostic(browserResult.missingGlyph, "font-unavailable"));
assert(diagnostic(browserResult.unsupportedStyle, "font-unavailable"));
assert(diagnostic(browserResult.emoji, "font-unavailable"));
assert(diagnostic(browserResult.permittedFallback, "font-fallback"));
const comparisons = shaping.map((sample, index) => ({ ...sample, fontId: sample.id, ...browserResult.samples[index], platformFonts: platformFonts[index], exactPlatformFace: platformFonts[index].length === 1 && platformFonts[index][0].postScriptName === fonts.find(f => f.id === sample.id).postscriptName && platformFonts[index][0].isCustomFont, advanceDifference: browserResult.samples[index].browserAdvance - sample.advance }));
for (const comparison of comparisons) {
  assert(comparison.exactPlatformFace, `Browser used a different actual face: ${comparison.fontId}`);
  assert(comparison.platformFonts.every(font => font.glyphCount > 0));
  assert.deepEqual(comparison.analysis.diagnostics, []);
  assert(Math.abs(comparison.advanceDifference) < .05, `HB/browser advance disagreement: ${comparison.fontId}`);
  assert(Math.abs(comparison.analysis.measurements[0].bounds.width - comparison.browserAdvance) < .02, `Layout/browser advance disagreement: ${comparison.fontId}`);
}
for (const analysis of [browserResult.missingGlyph, browserResult.unsupportedStyle, browserResult.emoji, browserResult.permittedFallback]) {
  assert(analysis.diagnostics.every(d => d.slideId === "s" && d.elementId === "t" && d.path === "slides.0.elements.0.content.0.runs.0.style.fontFamily"));
}
const evidence = { status: "observed", recordedAt: new Date().toISOString(), artifactsDirectory: path.relative(process.cwd(), directory), harfbuzz: (await execute("hb-shape", ["--version"], execution)).stdout.trim(), chromium: browserVersion, fonts: fonts.map(({ bytes: omitted, ...font }) => font), comparisons, controls: { ...browserResult, samples: undefined }, limitations: ["Installed regular faces only; no synthetic bold/italic", "Isolated RTL shaping is not mixed bidi certification", "Emoji is an explicit unavailable control, not supported shaping", "No native line-layout equivalence or source-font writer admission"] };
await writeFile(path.join(directory, "evidence.json"), JSON.stringify(evidence, null, 2) + "\n");
await writeFile("docs/testing/real-font-shaping-evidence.json", JSON.stringify(evidence, null, 2) + "\n");
console.log(JSON.stringify({ directory, comparisons: comparisons.map(c => ({ id: c.id, text: c.text, exactPlatformFace: c.exactPlatformFace, advanceDifference: c.advanceDifference })), controlsVerified: true }));
