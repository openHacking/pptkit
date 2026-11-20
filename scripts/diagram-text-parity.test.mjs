import assert from "node:assert/strict";
import test from "node:test";
import { readFile, mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { chromium } from "@playwright/test";
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
import { importPptxFile } from "../packages/pptx-parser/dist/node.js";
import { diagramTextFixture } from "../packages/pptx-parser/test/diagram-text-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
const options = { projection: "content", diagramPreview: "cache", editSaveProfile: "text-runs-v3", diagramTextComparison: "direct-paragraphs" };

test("Node, browser ESM/global agree with independent direct-paragraph DOM observations", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(); const requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  const esm = await readFile(new URL("../packages/pptx-parser/dist/browser/index.js", import.meta.url), "utf8");
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, esm);
  await page.addScriptTag({ content: await readFile(new URL("../packages/pptx-parser/dist/browser/global.js", import.meta.url), "utf8") });
  for (const strict of [false, true]) for (const different of [false, true]) {
    const f = diagramTextFixture(strict), source = ["Leading & 😀\r", "Line\nbreak", "", "e\u0301"], cached = different ? ["Old cache"] : source;
    f.setText(source, cached); const bytes = f.bytes(), expected = (await inspectPptx(bytes, options)).diagramInventory.textComparisons;
    const actual = await page.evaluate(async ({ bytes, options, dataXml, cacheXml, d, a }) => {
      const parse = xml => { const doc = new DOMParser().parseFromString(xml, "application/xml"); if (doc.querySelector("parsererror")) throw new Error("Invalid fixture XML"); return doc; };
      const data = parse(dataXml), cache = parse(cacheXml), all = doc => [...doc.getElementsByTagName("*")];
      const point = [...data.getElementsByTagNameNS(d, "pt")].find(p => p.getAttribute("modelId") === "1");
      const pres = [...data.getElementsByTagNameNS(d, "pt")].find(p => p.getAttribute("modelId") === "2");
      const shape = cache.getElementsByTagNameNS("http://schemas.microsoft.com/office/drawing/2008/diagram", "sp")[0];
      const paragraphs = body => [...body.children].filter(n => n.namespaceURI === a && n.localName === "p").map(p => [...p.children].map(r => r.localName === "br" ? "\n" : r.localName === "r" ? [...r.children].find(c => c.namespaceURI === a && c.localName === "t").textContent : "").join(""));
      const independent = { source: paragraphs([...point.children].find(n => n.localName === "t")), cache: paragraphs([...shape.children].find(n => n.localName === "txBody")), sourceIndex: all(data).indexOf(point) + 1, presIndex: all(data).indexOf(pres) + 1, cacheIndex: all(cache).indexOf(shape) + 1 };
      const reports = [];
      for (const api of [globalThis.parser, globalThis.PPTKit]) reports.push((await api.inspectPptx(Uint8Array.from(bytes), options)).diagramInventory.textComparisons);
      return { independent, reports };
    }, { bytes: [...bytes], options, dataXml: f.parts["deck/diagrams/data.xml"], cacheXml: f.parts["deck/diagrams/cache.xml"], d: f.D, a: f.A });
    for (const report of actual.reports) assert.deepEqual(report, expected);
    const s = expected[0].shapes[0]; assert.equal(s.status, different ? "different" : "equal"); assert.equal(expected[0].freshness, "unknown");
    assert.deepEqual(s.sourceParagraphs, actual.independent.source); assert.deepEqual(s.cacheParagraphs, actual.independent.cache);
    assert.deepEqual([s.sourcePointElementIndex, s.presentationElementIndex, s.cacheElementIndex], [actual.independent.sourceIndex, actual.independent.presIndex, actual.independent.cacheIndex]);
  }
  for (const [part, anchor, extra] of [["deck/diagrams/data.xml", "<d:t>", "\u00a0"], ["deck/diagrams/data.xml", "<a:r>", "<a:future/>"], ["deck/diagrams/cache.xml", "<dsp:txBody>", "\u00a0"]]) {
    const f = diagramTextFixture();
    f.parts[part] = f.parts[part].replace(anchor, `${anchor}${extra}`);
    const bytes = f.bytes(), expected = (await inspectPptx(bytes, options)).diagramInventory;
    const reports = await page.evaluate(async ({ bytes, options }) => Promise.all([globalThis.parser, globalThis.PPTKit].map(async api => (await api.inspectPptx(Uint8Array.from(bytes), options)).diagramInventory)), { bytes: [...bytes], options });
    for (const report of reports) assert.deepEqual(report, expected);
    if (part.includes("data.xml")) assert.equal(expected.textComparisons[0].status, "source-only");
    else assert.equal(expected.textComparisons[0].shapes[0].status, "unavailable");
  }
  assert.deepEqual(requests, []);
});

test("three unrelated edit/save rounds retain differing data/cache payloads and observations", async () => {
  for (const strict of [false, true]) {
    const f = diagramTextFixture(strict); f.setText(["Fresh source"], ["Stored cache"]); let bytes = f.bytes();
    const expected = (await inspectPptx(bytes, options)).diagramInventory.textComparisons;
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(bytes), session = await importPptx(bytes, options);
      try {
        assert.deepEqual(session.report.inspection.diagramInventory.textComparisons, expected);
        const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "text"), text = structuredClone(slide.elements[index]);
        text.content[0].runs[0].text = `Unrelated ${round}`; slide.removeElement(text.id); slide.insertElement(index, text);
        const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready"); bytes = (await session.save(prepared.plan)).bytes;
        const after = readGeneratedZip(bytes); assert.deepEqual([...before.keys()], [...after.keys()]);
        for (const [part, value] of before) if (part !== "deck/z.xml") assert.deepEqual(after.get(part), value, part);
        assert.deepEqual((await inspectPptx(bytes, options)).diagramInventory.textComparisons, expected);
      } finally { session.dispose(); }
    }
  }
});

test("opaque source text survives unrelated source saves without a guessed comparison", async () => {
 for (const [anchor, extra] of [["<d:t>", "\u00a0"], ["<a:r>", "<a:future/>"]]) {
  const f = diagramTextFixture(); f.parts["deck/diagrams/data.xml"] = f.parts["deck/diagrams/data.xml"].replace(anchor, `${anchor}${extra}`);
  const bytes = f.bytes(), before = readGeneratedZip(bytes), session = await importPptx(bytes, options);
  try {
    assert.equal(session.report.inspection.diagramInventory.dataModels[0].status, "partial");
    assert.equal(session.report.inspection.diagramInventory.textComparisons[0].status, "source-only");
    const noOp = await session.prepareSave({ mode: "no-op" }); assert.deepEqual((await session.save(noOp.plan)).bytes, new Uint8Array(bytes));
    const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "text"), text = structuredClone(slide.elements[index]);
    text.content[0].runs[0].text = "Unrelated"; slide.removeElement(text.id); slide.insertElement(index, text);
    const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
    const saved = (await session.save(prepared.plan)).bytes, after = readGeneratedZip(saved);
    assert.deepEqual(after.get("deck/diagrams/data.xml"), before.get("deck/diagrams/data.xml"));
    assert.equal((await inspectPptx(saved, options)).diagramInventory.textComparisons[0].status, "source-only");
  } finally { session.dispose(); }
 }
});

test("Node file import captures the comparison profile before filesystem awaits", async t => {
  const directory = await mkdtemp(path.join(tmpdir(), "pptkit-diagram-text-")); t.after(() => rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, "source.pptx"); await writeFile(file, diagramTextFixture().bytes());
  const config = { ...options }, pending = importPptxFile(file, config); config.diagramTextComparison = "invalid";
  const { session } = await pending;
  try { assert.equal(session.report.inspection.diagramInventory.textComparisons[0].shapes[0].status, "equal"); } finally { session.dispose(); }
});
