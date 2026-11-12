import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { scriptFontImportFixture, sourceFonts } from "../packages/pptx-parser/test/script-font-import-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { createScriptFontFixture } from "../tests/fixtures/script-fonts.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const options = { projection: "content", textFontProfile: "script-families-v1", editSaveProfile: "text-runs-v3" };
const moduleUrl = s => `data:text/javascript;base64,${Buffer.from(s).toString("base64")}`;
async function rounds(bytes) {
  const results = [];
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, options);
    try {
      const document = normalizePresentation(s.document), report = s.report.textStyleInheritance, slide = s.document.slides[0], next = structuredClone(slide.elements[0]);
      next.text.content[0].runs[0].text = `Scalar ${round} 中文العربية`; slide.replaceElement(next.id, next);
      const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); const saved = await s.save(plan.plan);
      results.push({ document, report, saved: [...saved.bytes] }); bytes = saved.bytes;
    } finally { s.dispose(); }
  }
  return results;
}

test("Node/ESM/global independent font fields and three source saves agree with actual source DOM paths", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const core = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8")), parser = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const strict of [false, true]) {
    const f = scriptFontImportFixture({ strict, run: sourceFonts("+mn-lt", "+mn-ea", "+mn-cs") }), bytes = encodeStructure(f), expected = await rounds(bytes);
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
      const actual = await page.evaluate(async ({ core, parser, bytes, options, parts, namespace }) => {
        const c = core ? await import(core) : PPTKit, p = parser ? await import(parser) : PPTKit, results = [];
        for (let round = 0; round < 3; round++) {
          const s = await p.importPptx(Uint8Array.from(bytes), options);
          try {
            const document = c.normalizePresentation(s.document), report = s.report.textStyleInheritance, slide = s.document.slides[0], next = structuredClone(slide.elements[0]); next.text.content[0].runs[0].text = `Scalar ${round} 中文العربية`; slide.replaceElement(next.id, next);
            const plan = await s.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw new Error("Profile was not retained for candidate validation"); const saved = await s.save(plan.plan);
            results.push({ document, report, saved: [...saved.bytes] }); bytes = [...saved.bytes];
          } finally { s.dispose(); }
        }
        const docs = Object.fromEntries(Object.entries(parts).filter(([, v]) => typeof v === "string" && v.startsWith("<")).map(([name, xml]) => [`/${name}`, new DOMParser().parseFromString(xml, "application/xml")])); let checked = 0;
        const expectedSlot = { "run.fontFamily": "lt", "run.eastAsianFontFamily": "ea", "run.complexScriptFontFamily": "cs" };
        for (const property of results[0].report.properties) {
          const slot = expectedSlot[property.property]; if (!slot || property.status !== "mapped") continue;
          const doc = docs[property.source.part], node = doc.evaluate(property.source.path, doc, prefix => prefix === "a" ? namespace : doc.documentElement.lookupNamespaceURI(prefix), XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
          if (node?.value !== `+mn-${slot}`) throw new Error("Wrong actual source slot reference");
          const theme = docs["/deck/theme.xml"], group = theme.getElementsByTagNameNS(namespace, "minorFont")[0], tag = slot === "lt" ? "latin" : slot;
          if (group.getElementsByTagNameNS(namespace, tag)[0]?.getAttribute("typeface") !== property.value) throw new Error("Independent selected theme slot differs");
          for (const dep of property.resolvedFrom) {
            const source = docs[dep.part], selected = source.evaluate(dep.path, source, prefix => prefix === "a" ? namespace : null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (selected?.value !== property.value) throw new Error("Wrong actual dependency value");
          }
          checked++;
        }
        return { results, checked };
      }, { core: format === "esm" ? core : null, parser: format === "esm" ? parser : null, bytes: [...bytes], options, parts: f.parts, namespace: f.A });
      assert.deepEqual(actual.results, expected); assert.equal(actual.checked, 6); assert.deepEqual(requests, []); await page.close();
    }
    const before = readGeneratedZip(bytes), after = readGeneratedZip(Uint8Array.from(expected.at(-1).saved));
    for (const [name, payload] of before) if (name !== "deck/z.xml") assert.deepEqual(after.get(name), payload);
  }
});

const effectiveFonts = document => {
  const p = normalizePresentation(document), rows = [], family = v => typeof v === "string" ? v : p.theme.fonts[v.theme];
  const content = (owner, data) => { for (const paragraph of data) for (const run of paragraph.runs) rows.push({ owner, text: run.text, fonts: [run.style.fontFamily, run.style.eastAsianFontFamily ?? run.style.fontFamily, run.style.complexScriptFontFamily ?? run.style.fontFamily].map(family) }); };
  const element = e => { if (e.type === "text") content("text", e.content); else if (e.type === "shape" && e.text) content("shape", e.text.content); else if (e.type === "table") for (const row of e.rows) for (const cell of row.cells) content("table", cell.content); else if (e.type === "group") e.children.forEach(element); };
  for (const slide of p.slides) { slide.elements.forEach(element); content("notes", slide.notes); }
  return rows;
};
test("mixed-owner script families retain semantic values over three fresh rounds with independent native slot DOM", async t => {
  const original = createScriptFontFixture(createPresentation), expected = effectiveFonts(original), browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  let bytes = (await generatePptx(original)).bytes;
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, options);
    try {
      assert.deepEqual(effectiveFonts(s.document), expected);
      const generated = await generatePptx(s.document), parts = readGeneratedZip(generated.bytes);
      const actual = await page.evaluate(({ slide, notes }) => {
        const ns = "http://schemas.openxmlformats.org/drawingml/2006/main", parser = new DOMParser();
        return [slide, notes].flatMap(xml => [...parser.parseFromString(xml, "application/xml").getElementsByTagNameNS(ns, "r")].map(r => ({ text: r.getElementsByTagNameNS(ns, "t")[0]?.textContent, fonts: ["latin", "ea", "cs"].map(name => r.getElementsByTagNameNS(ns, name)[0]?.getAttribute("typeface")) })));
      }, { slide: parts.get("ppt/slides/slide1.xml").toString(), notes: parts.get("ppt/notesSlides/notesSlide1.xml").toString() });
      assert.deepEqual(actual, expected.map(({ text, fonts }) => ({ text, fonts })));
      const noOp = await s.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual((await s.save(noOp.plan)).bytes, bytes); bytes = generated.bytes;
    } finally { s.dispose(); }
  }
});

test("browser profile capture, budget failure, cancellation and default retry remain independent", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(); for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const bytes = encodeStructure(scriptFontImportFixture());
  const result = await page.evaluate(async ({ input, options }) => {
    const data = Uint8Array.from(input), config = { ...options }, pending = PPTKit.importPptx(data, config); data.fill(0); config.textFontProfile = undefined;
    const captured = await pending, marker = captured.report.textFontProfile; captured.dispose(); let invalid, limited, aborted;
    try { await PPTKit.importPptx(Uint8Array.from(input), { projection: "structure", textFontProfile: "script-families-v1" }); } catch (e) { invalid = e.name; }
    try { await PPTKit.importPptx(Uint8Array.from(input), { ...options, sessionLimits: { maxTextStyleProperties: 1 } }); } catch (e) { limited = e.code; }
    const c = new AbortController(), operation = PPTKit.importPptx(Uint8Array.from(input), { ...options, signal: c.signal }); queueMicrotask(() => c.abort()); try { await operation; } catch (e) { aborted = e.code; }
    const defaultSession = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" }), old = Object.hasOwn(defaultSession.report, "textFontProfile"); defaultSession.dispose();
    const retry = await PPTKit.importPptx(Uint8Array.from(input), options), family = retry.document.slides[0].elements[0].text.content[0].runs[0].style.eastAsianFontFamily; retry.dispose();
    return { marker, invalid, limited, aborted, old, family };
  }, { input: [...bytes], options });
  assert.deepEqual(result, { marker: "script-families-v1", invalid: "TypeError", limited: "import-limit-exceeded", aborted: "operation-aborted", old: false, family: "Eastern Source" });
});
