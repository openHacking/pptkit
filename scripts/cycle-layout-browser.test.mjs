import assert from "node:assert/strict";
import test from "node:test";
import { readFile, mkdir } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { resolveCycleLayout } from "./research/cycle-layout.mjs";
const input = paragraphs => ({ family: "cycle", closure: "layout-generated", roots: paragraphs.map((_, i) => `n${i}`), nodes: paragraphs.map((text, i) => ({ id: `n${i}`, role: "content", children: [], paragraphs: [text] })) });
const url = async file => `data:text/javascript;base64,${(await readFile(new URL(file, import.meta.url))).toString("base64")}`;

test("explicit loaded-font measurement connects Cycle layout to the public browser measurement provider", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage();
  await page.route("**/*", route => route.abort());
  const imports = Object.fromEntries(await Promise.all([["@pptkit/core", "../packages/core/dist/browser/index.js"], ["@pptkit/layout", "../packages/layout/dist/browser/index.js"], ["@pptkit/layout/dom", "../packages/layout/dist/browser/dom.js"], ["cycle-research", "./research/cycle-layout.mjs"]].map(async ([name, file]) => [name, await url(file)])));
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  const font = (await readFile(new URL("../packages/svg-renderer/test/fixtures/roboto-mono-latin-400-normal.woff2", import.meta.url))).toString("base64");
  const result = await page.evaluate(async ({ font, candidate }) => {
    const { createPresentation } = await import("@pptkit/core"), { resolveLayout } = await import("@pptkit/layout"), { createBrowserTextMeasurement } = await import("@pptkit/layout/dom"), { resolveCycleLayout } = await import("cycle-research");
    const face = new FontFace("Cycle Pinned", `url(data:font/woff2;base64,${font})`); await face.load(); document.fonts.add(face);
    const adapter = createBrowserTextMeasurement({ document, fonts: [{ family: "Cycle Pinned", faces: [face], covers: ({ text, bold, italic }) => !bold && !italic && /^[\x00-\x7f]*$/.test(text) }] });
    let calls = 0;
    const measureText = request => {
      calls++;
      const deck = createPresentation({ theme: { fonts: { body: "Cycle Pinned", heading: "Cycle Pinned" } } });
      deck.addSlide({ id: "measure" }).addElement({ type: "text", id: request.id, box: { x: 0, y: 0, width: request.availableWidth, height: request.availableHeight }, frame: { margin: 0, wrap: true, verticalAlign: "top" }, content: request.paragraphs.map(text => ({ style: { lineSpacing: 1.2 }, runs: [{ text, style: { fontSize: 12 } }] })) });
      const analysis = resolveLayout(deck, { textMeasurement: adapter, diagnostics: { textOverflow: true } }).analysis;
      if (!analysis.measurements[0]) throw new Error(analysis.diagnostics.map(d => d.message).join(" ") || "Text measurement unavailable");
      return analysis.measurements[0].bounds;
    };
    const frame = { x: 10, y: 20, width: 520, height: 300 };
    const layout = resolveCycleLayout(candidate, { frame, measureText });
    let strictError, missingGlyphError, disposedError;
    try { resolveCycleLayout(candidate, { frame, measureText, overflow: "error" }); } catch (e) { strictError = e.message; }
    const missing = structuredClone(candidate); missing.nodes[0].paragraphs = ["中"];
    try { resolveCycleLayout(missing, { frame, measureText }); } catch (e) { missingGlyphError = e.message; }
    adapter.dispose(); try { resolveCycleLayout(candidate, { frame, measureText }); } catch (e) { disposedError = e.message; }
    return { layout, calls, strictError, missingGlyphError, disposedError, hostsAfterDispose: document.querySelectorAll("[data-pptkit-measurement]").length };
  }, { font, candidate: input(["Short", "Long label ".repeat(80), "Third"]) });
  assert.deepEqual(result.layout.nodes.map(n => n.textFit), ["fits", "overflow", "fits"], JSON.stringify(result.layout.nodes.map(n => ({ textBox: n.textBox, measured: n.measuredBounds }))));
  assert.deepEqual(result.layout.findings, [{ nodeId: "n1", code: "text-overflow" }]);
  assert.match(result.strictError, /text overflows node n1/); assert.match(result.missingGlyphError, /[Ff]ont.*unavailable/); assert.match(result.disposedError, /disposed/);
  assert.equal(result.hostsAfterDispose, 0); assert.ok(result.calls >= 7);
});

test("browser paints hand-calculated circles and separate two-node paths; generated research preview is reviewable", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage({ viewport: { width: 800, height: 430 }, deviceScaleFactor: 1 });
  await page.route("**/*", route => route.abort());
  const frame = { x: 0, y: 0, width: 300, height: 300 }, square = resolveCycleLayout(input(["A", "B", "C", "D"]), { frame, nodeDiameter: 60 });
  const pair = resolveCycleLayout(input(["A", "B"]), { frame });
  const fragment = layout => layout.edges.map(e => e.kind === "arrow" ? `<polygon points="${e.polygon.map(p => `${p.x},${p.y}`).join(" ")}" fill="#9CB0D8"/>` : `<path d="M ${e.start.x} ${e.start.y} A ${e.radius} ${e.radius} 0 0 ${e.clockwise ? 1 : 0} ${e.end.x} ${e.end.y}" fill="none" stroke="#9CB0D8" stroke-width="2"/><polygon points="${e.head.map(p => `${p.x},${p.y}`).join(" ")}" fill="#9CB0D8"/>`).join("") + layout.nodes.map(n => `<circle data-node="${n.id}" cx="${n.center.x}" cy="${n.center.y}" r="${n.diameter / 2}" fill="#4266B1"/><text x="${n.center.x}" y="${n.center.y + 5}" text-anchor="middle" fill="white" font-size="18">${n.paragraphs[0]}</text>`).join("");
  await page.setContent(`<style>body{margin:0;background:#f5f7fa;font-family:sans-serif}h3{margin:16px 0;font-size:17px;font-weight:500}.panel{display:inline-block;margin:0 30px;vertical-align:top}svg{background:white;border:1px solid #d9dfeb}</style><div class="panel"><h3>Four-node geometry</h3><svg id="square" width="320" height="320" viewBox="-10 -10 320 320">${fragment(square)}</svg></div><div class="panel"><h3>Two-node connector policy</h3><svg id="pair" width="320" height="320" viewBox="-10 -10 320 320">${fragment(pair)}</svg></div>`);
  const observed = await page.evaluate(() => {
    const svg = document.querySelector("#square"), host = svg.getBoundingClientRect();
    return { circles: [...svg.querySelectorAll("circle")].map(c => { const b = c.getBoundingClientRect(); return { id: c.getAttribute("data-node"), x: b.x - host.x - 1 - 10, y: b.y - host.y - 1 - 10, width: b.width, height: b.height }; }), paths: [...document.querySelectorAll("#pair path")].map(p => ({ d: p.getAttribute("d"), length: p.getTotalLength(), bounds: { x: p.getBBox().x, y: p.getBBox().y, width: p.getBBox().width, height: p.getBBox().height } })) };
  });
  const r = 45 * Math.SQRT2, expected = [{ x: 120, y: 120 - r }, { x: 120 + r, y: 120 }, { x: 120, y: 120 + r }, { x: 120 - r, y: 120 }];
  observed.circles.forEach((box, i) => { assert.equal(box.id, `n${i}`); for (const key of ["x", "y"]) assert.ok(Math.abs(box[key] - expected[i][key]) < 0.001); assert.ok(Math.abs(box.width - 60) < 0.001); assert.ok(Math.abs(box.height - 60) < 0.001); });
  assert.equal(observed.paths.length, 2); assert.notEqual(observed.paths[0].d, observed.paths[1].d); assert.ok(observed.paths.every(p => p.length > 10));
  assert.ok(Math.abs(observed.paths[0].bounds.x - observed.paths[1].bounds.x) > 10);
  await mkdir(".tmp/cycle-layout", { recursive: true }); await page.screenshot({ path: ".tmp/cycle-layout/browser-preview.png" });
});
