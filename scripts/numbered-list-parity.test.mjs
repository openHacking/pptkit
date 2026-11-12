import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

function fixture(create) {
  const deck = create({ id: "numbered" }), slide = deck.addSlide({ id: "slide" });
  for (const [index, value] of [1, 26, 27, 28, 52, 53, 702, 703, 32767].entries()) slide.addElement({ id: `item-${value}`, type: "text", box: { x: 30, y: index * 40, width: 300, height: 35 }, content: [{ style: { bullet: { type: "number", style: "alphaLowerPeriod", startAt: value } }, runs: [{ text: `ITEM-${value}` }] }] });
  return deck;
}

test("alphabetic numbered starts preserve source values and match preview across Node and browser", async t => {
  const deck = fixture(createPresentation), generated = await generatePptx(deck), preview = await renderPresentationToSvg(deck), values = [1, 26, 27, 28, 52, 53, 702, 703, 32767];
  const bytes = readGeneratedZip(generated.bytes), xml = bytes.get("ppt/slides/slide1.xml").toString();
  assert.deepEqual([...xml.matchAll(/<a:buAutoNum type="alphaLcPeriod" startAt="(\d+)"\/>/g)].map(m => Number(m[1])), values);
  const session = await importPptx(generated.bytes, { projection: "content" });
  try { assert.deepEqual(session.document.slides[0].elements.map(e => e.content[0].style.bullet.startAt), values); }
  finally { session.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  const imports = {}, globals = [];
  for (const name of ["core", "layout", "pptx-exporter", "svg-renderer"]) {
    imports[`@pptkit/${name}`] = `data:text/javascript;base64,${Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`, "utf8")).toString("base64")}`;
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`, "utf8"));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`); for (const source of globals) await page.addScriptTag({ content: source });
  for (const format of ["esm", "global"]) {
    const result = await page.evaluate(async ({ format, source }) => {
      const api = format === "global" ? PPTKit : Object.assign({}, await import("@pptkit/core"), await import("@pptkit/pptx-exporter"), await import("@pptkit/svg-renderer"));
      const build = Function(`return (${source})`)(), deck = build(api.createPresentation), output = await api.generatePptx(deck), preview = await api.renderPresentationToSvg(deck);
      return { bytes: [...output.bytes], preview };
    }, { format, source: fixture.toString() });
    assert.deepEqual(result.bytes, [...generated.bytes]); assert.deepEqual(result.preview, preview);
  }
  assert.deepEqual(requests, []);
});

function continuationFixture(create) {
  const deck = create({ id: "continued" }), slide = deck.addSlide({ id: "slide" });
  const content = [25, "continue", "continue", "continue", 52, "continue"].map((startAt, i) => ({
    style: { bullet: { type: "number", style: "alphaLowerPeriod", startAt }, ...(i === 0 ? { tabStops: [{ position: 36, align: "left" }, { position: 72.5, align: "center" }, { position: 144, align: "right" }, { position: 216, align: "decimal" }] } : {}) }, runs: [{ text: i === 0 ? "ITEM\t0" : `ITEM-${i}` }],
  }));
  slide.addElement({ id: "item", type: "text", box: { x: 30, y: 30, width: 300, height: 300 }, content });
  slide.setNotes(content);
  return deck;
}

test("authored continuation resolves across preview, PPTX, import and browser with explicit loss warning", async t => {
  const deck = continuationFixture(createPresentation), original = JSON.stringify(deck.slides[0].elements), values = [25, 26, 27, 28, 52, 53];
  assert.deepEqual(normalizePresentation(deck).slides[0].elements[0].content.map(p => p.style.bullet.startAt), values);
  const generated = await generatePptx(deck), preview = await renderPresentationToSvg(deck), archive = readGeneratedZip(generated.bytes);
  assert.ok(preview.warnings.some(w => w.code === "text-tab-layout-unsupported" && w.elementId === "item"));
  for (const path of ["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"]) {
    const xml = archive.get(path).toString();
    assert.deepEqual([...xml.matchAll(/<a:buAutoNum type="alphaLcPeriod" startAt="(\d+)"\/>/g)].map(m => Number(m[1])), values);
    assert.match(xml, /<a:tabLst><a:tab pos="457200" algn="l"\/><a:tab pos="920750" algn="ctr"\/><a:tab pos="1828800" algn="r"\/><a:tab pos="2743200" algn="dec"\/><\/a:tabLst>/);
  }
  const warnings = generated.warnings.filter(w => w.code === "numbered-list-continuation-materialized");
  assert.equal(warnings.length, 8);
  assert.ok(warnings.some(w => w.slideId === "slide" && w.elementId === "item" && w.path === "slides.0.elements.0.content.1.style.bullet.startAt"));
  assert.ok(warnings.some(w => w.slideId === "slide" && w.path === "slides.0.notes.1.style.bullet.startAt"));
  const session = await importPptx(generated.bytes, { projection: "content" });
  try {
    assert.deepEqual(session.document.slides[0].elements[0].content.map(p => p.style.bullet.startAt), values);
    assert.deepEqual(session.document.slides[0].elements[0].content[0].style.tabStops, [{ position: 36, align: "left" }, { position: 72.5, align: "center" }, { position: 144, align: "right" }, { position: 216, align: "decimal" }]);
    assert.deepEqual(session.document.slides[0].notes[0].style.tabStops, session.document.slides[0].elements[0].content[0].style.tabStops);
  }
  finally { session.dispose(); }
  assert.equal(JSON.stringify(deck.slides[0].elements), original);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  const imports = {}, globals = [];
  for (const name of ["core", "layout", "pptx-exporter", "svg-renderer", "pptx-parser"]) {
    imports[`@pptkit/${name}`] = `data:text/javascript;base64,${Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`, "utf8")).toString("base64")}`;
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`, "utf8"));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`); for (const source of globals) await page.addScriptTag({ content: source });
  for (const format of ["esm", "global"]) {
    const result = await page.evaluate(async ({ format, source }) => {
      const api = format === "global" ? PPTKit : Object.assign({}, await import("@pptkit/core"), await import("@pptkit/pptx-exporter"), await import("@pptkit/svg-renderer"), await import("@pptkit/pptx-parser"));
      const deck = Function(`return (${source})`)()(api.createPresentation), output = await api.generatePptx(deck);
      const session = await api.importPptx(output.bytes, { projection: "content" });
      const importedTabs = session.document.slides[0].elements[0].content[0].style.tabStops; session.dispose();
      return { bytes: [...output.bytes], warnings: output.warnings, preview: await api.renderPresentationToSvg(deck), importedTabs };
    }, { format, source: continuationFixture.toString() });
    assert.deepEqual(result.bytes, [...generated.bytes]); assert.deepEqual(result.warnings, generated.warnings); assert.deepEqual(result.preview, preview);
    assert.deepEqual(result.importedTabs, [{ position: 36, align: "left" }, { position: 72.5, align: "center" }, { position: 144, align: "right" }, { position: 216, align: "decimal" }]);
  }
  assert.deepEqual(requests, []);
});

test("continuation warnings identify grouped, shape and table owners and export rejects an overflow", async () => {
  const deck = createPresentation(), slide = deck.addSlide({ id: "s" });
  const content = [1, "continue"].map(startAt => ({ style: { bullet: { type: "number", startAt } }, runs: [{ text: "x" }] }));
  const box = { x: 0, y: 0, width: 200, height: 100 };
  slide.addElement({ id: "shape", type: "shape", shape: "rect", box, text: { content } });
  slide.addElement({ id: "table", type: "table", box, columns: [200], rows: [{ height: 100, cells: [{ content }] }] });
  slide.addElement({ id: "group", type: "group", box, coordinateSize: { width: 200, height: 100 }, children: [{ id: "child", type: "text", box, content }] });
  const output = await generatePptx(deck), warnings = output.warnings.filter(w => w.code === "numbered-list-continuation-materialized");
  assert.deepEqual(warnings.map(w => [w.elementId, w.path]), [
    ["shape", "slides.0.elements.0.text.content.1.style.bullet.startAt"],
    ["table", "slides.0.elements.1.rows.0.cells.0.content.1.style.bullet.startAt"],
    ["child", "slides.0.elements.2.children.0.content.1.style.bullet.startAt"],
  ]);
  const overflow = continuationFixture(createPresentation), item = structuredClone(overflow.slides[0].elements[0]);
  item.content[0].style.bullet.startAt = 32767; overflow.slides[0].replaceElement("item", item);
  const before = JSON.stringify(overflow.slides[0].elements);
  await assert.rejects(generatePptx(overflow));
  assert.equal(JSON.stringify(overflow.slides[0].elements), before);
});

test("continuation in master, layout and placeholder retains explicit numbering in native inheritance mode", async () => {
  const deck = createPresentation(), box = { x: 10, y: 10, width: 300, height: 100 };
  const content = [3, "continue"].map(startAt => ({ style: { bullet: { type: "number", startAt } }, runs: [{ text: "x" }] }));
  deck.defineSlideMaster({ id: "m", name: "Master", elements: [{ id: "master-item", type: "text", box, content }] });
  deck.defineSlideLayout({ id: "l", name: "Layout", masterId: "m", elements: [{ id: "layout-item", type: "text", box, content }], placeholders: [{ key: "body", kind: "body", box }] });
  deck.addSlide({ id: "s", layoutId: "l" }).addElement({ id: "placeholder", type: "text", placeholderKey: "body", content });
  const output = await generatePptx(deck, { nativePlaceholderText: "inherit" });
  const warnings = output.warnings.filter(w => w.code === "numbered-list-continuation-materialized");
  assert.deepEqual(warnings.map(w => [w.masterId ?? w.layoutId ?? w.slideId, w.elementId]), [["m", "master-item"], ["l", "layout-item"], ["s", "placeholder"]]);
  const zip = readGeneratedZip(output.bytes);
  let checked = 0;
  for (const [path, bytes] of zip) if (/^ppt\/(?:slideMasters|slideLayouts|slides)\/(?:slideMaster|slideLayout|slide)\d+\.xml$/.test(path) && bytes.toString().includes("<a:t>x</a:t>")) {
    assert.match(bytes.toString(), /<a:buAutoNum type="arabicPeriod" startAt="4"\/>/, path);
    checked++;
  }
  assert.equal(checked, 3);
});

test("explicitly cleared tab stops and table-region stops are written to PPTX", async () => {
  const deck = createPresentation({ textStylePresets: { tabs: { paragraph: { tabStops: [{ position: 36, align: "left" }] } } } }), slide = deck.addSlide();
  const box = { x: 0, y: 0, width: 200, height: 100 };
  slide.addElement({ id: "clear", type: "text", box, textStylePreset: "tabs", content: [{ style: { tabStops: [] }, runs: [{ text: "A\tB" }] }] });
  slide.addElement({ id: "table", type: "table", box, columns: [200], rows: [{ height: 100, cells: [{ content: "C\tD" }] }], style: { wholeTable: { text: { paragraph: { tabStops: [{ position: 72, align: "decimal" }] } } } } });
  const bytes = (await generatePptx(deck)).bytes, xml = readGeneratedZip(bytes).get("ppt/slides/slide1.xml").toString();
  assert.match(xml, /<a:tabLst><\/a:tabLst>/);
  assert.match(xml, /<a:tabLst><a:tab pos="914400" algn="dec"\/><\/a:tabLst>/);
  const session = await importPptx(bytes, { projection: "content" });
  try {
    assert.deepEqual(session.document.slides[0].elements[0].content[0].style.tabStops, []);
    assert.deepEqual(session.document.slides[0].elements[1].rows[0].cells[0].content[0].style.tabStops, [{ position: 72, align: "decimal" }]);
  } finally { session.dispose(); }
});

test("fresh export warns only for the observed zero-indent tab-stop combination", async () => {
  const deck = createPresentation(), slide = deck.addSlide({ id: "s" }), box = { x: 0, y: 0, width: 200, height: 80 };
  for (const [id, tabStops, hanging, text] of [
    ["observed", [{ position: 72, align: "left" }], 0, "A\tB"],
    ["nonzero", [{ position: 72, align: "left" }], 12, "A\tB"],
    ["no-tab", [{ position: 72, align: "left" }], 0, "AB"],
    ["cleared", [], 0, "A\tB"],
  ]) slide.addElement({ id, type: "text", box, content: [{ style: { tabStops, hanging }, runs: [{ text }] }] });
  const warnings = (await generatePptx(deck)).warnings.filter(w => w.code === "tab-stops-libreoffice-zero-indent");
  assert.deepEqual(warnings.map(w => [w.slideId, w.elementId]), [["s", "observed"]]);
});
