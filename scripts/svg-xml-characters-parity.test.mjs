import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { chromium } from "@playwright/test";
import { createPresentation } from "../packages/core/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { inspectXmlPart } from "../packages/pptx-parser/dist/index.js";

function fixture(create) {
  const deck = create({ metadata: { title: "Title\u0000 & 😀" } });
  deck.addSlide({ id: "page", elements: [
    { id: "text", type: "text", name: "Name\u0002", box: { x: 10, y: 10, width: 300, height: 80 }, content: "A\u0008 & <B> 😀" },
    { id: "shape", type: "shape", shape: "rect", box: { x: 10, y: 100, width: 300, height: 80 }, text: { content: "Shape\u000b" } },
    { id: "chart", type: "chart", chartType: "bar", box: { x: 350, y: 10, width: 400, height: 250 }, categories: ["Category\u0001", "Fine"], series: [{ name: "Legend\u0000", values: [1, 2] }], legend: { visible: true } },
  ] }); return deck;
}

test("XML-safe SVG and warnings agree across Node, browser ESM and global", async t => {
  const expected = await renderPresentationToSvg(fixture(createPresentation));
  assert.deepEqual(expected.warnings.filter(w => w.code.endsWith("xml-replacement") || w.code === "svg-xml-character-replacement").map(w => w.code).sort(), ["chart-text-xml-replacement", "svg-xml-character-replacement"]);
  await inspectXmlPart(new TextEncoder().encode(expected.slides[0].svg));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  const imports = {}, globals = [];
  for (const name of ["core", "layout", "svg-renderer"]) {
    imports[`@pptkit/${name}`] = `data:text/javascript;base64,${Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`, "utf8")).toString("base64")}`;
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`, "utf8"));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`); for (const source of globals) await page.addScriptTag({ content: source });
  for (const format of ["esm", "global"]) {
    const actual = await page.evaluate(async ({ format, source }) => {
      const api = format === "global" ? PPTKit : Object.assign({}, await import("@pptkit/core"), await import("@pptkit/svg-renderer"));
      const build = Function(`return (${source})`)(), result = await api.renderPresentationToSvg(build(api.createPresentation));
      const doc = new DOMParser().parseFromString(result.slides[0].svg, "image/svg+xml");
      return { result, errors: doc.getElementsByTagName("parsererror").length, labels: [...doc.getElementsByTagNameNS("http://www.w3.org/2000/svg", "text")].map(e => e.textContent) };
    }, { format, source: fixture.toString() });
    assert.deepEqual(actual.result, expected); assert.equal(actual.errors, 0); assert.ok(actual.labels.some(text => text.includes("Category�")));
  }
  assert.deepEqual(requests, []);
});
