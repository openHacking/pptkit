import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { textAppearanceFixture, fontCollection } from "../packages/pptx-parser/test/table-text-appearance-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const moduleUrl = s => `data:text/javascript;base64,${Buffer.from(s).toString("base64")}`;
async function rounds(bytes) {
  const results = [];
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
    try {
      const document = normalizePresentation(s.document), report = s.report.tableStyleInheritance, slide = s.document.slides[0], next = structuredClone(slide.elements[0]);
      next.rows[0].cells[0].content[0].runs[0].text = `Appearance ${round}`; slide.replaceElement(next.id, next);
      const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); const saved = await s.save(plan.plan);
      results.push({ document, report, saved: [...saved.bytes] }); bytes = saved.bytes;
    } finally { s.dispose(); }
  }
  return results;
}

test("Node/ESM/global table text appearance and three source saves agree with independent source DOM", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const core = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8")), parser = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const strict of [false, true]) for (const collection of [false, true]) {
    const f = textAppearanceFixture({ strict, ...(collection ? { text: fontCollection("+mn-lt", "+mn-ea", "+mn-cs") + '<a:srgbClr val="102030"/>' } : {}) }), bytes = encodeStructure(f), expected = await rounds(bytes);
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
      const result = await page.evaluate(async ({ core, parser, bytes, parts, namespace }) => {
        const c = core ? await import(core) : PPTKit, p = parser ? await import(parser) : PPTKit, results = [];
        for (let round = 0; round < 3; round++) {
          const s = await p.importPptx(Uint8Array.from(bytes), { projection: "content", editSaveProfile: "text-runs-v3" });
          try {
            const document = c.normalizePresentation(s.document), report = s.report.tableStyleInheritance, slide = s.document.slides[0], next = structuredClone(slide.elements[0]); next.rows[0].cells[0].content[0].runs[0].text = `Appearance ${round}`; slide.replaceElement(next.id, next);
            const plan = await s.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw new Error("Not ready"); const saved = await s.save(plan.plan);
            results.push({ document, report, saved: [...saved.bytes] }); bytes = [...saved.bytes];
          } finally { s.dispose(); }
        }
        const docs = Object.fromEntries(Object.entries(parts).filter(([, v]) => typeof v === "string" && v.startsWith("<")).map(([name, xml]) => [`/${name}`, new DOMParser().parseFromString(xml, "application/xml")]));
        const theme = docs["/deck/theme.xml"], slots = theme.getElementsByTagNameNS(namespace, "clrScheme")[0]; let checked = 0;
        for (const property of results[0].report.properties) {
          if (!["run.color", "run.fontFamily"].includes(property.property) || property.status !== "mapped") continue;
          const doc = docs[property.source.part], node = doc.evaluate(property.source.path, doc, prefix => prefix === "a" ? namespace : null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
          if (!node) throw new Error("Missing selected source");
          if (property.property === "run.color") {
            const expected = node.localName === "schemeClr" ? [...slots.children].find(n => n.localName === node.getAttribute("val")).firstElementChild.getAttribute("val") : node.getAttribute("val");
            if (expected !== property.value) throw new Error("Wrong independent source color");
          } else {
            const group = node.name === "idx" ? (node.value === "minor" ? "minorFont" : "majorFont") : "minorFont";
            const fonts = theme.getElementsByTagNameNS(namespace, group)[0];
            if ([...fonts.children].some(n => n.getAttribute("typeface") !== property.value)) throw new Error("Theme families do not agree");
            if (property.resolvedFrom.length < 1) throw new Error("Missing font dependencies");
          }
          checked++;
        }
        return { results, checked };
      }, { core: format === "esm" ? core : null, parser: format === "esm" ? parser : null, bytes: [...bytes], parts: f.parts, namespace: f.A });
      assert.deepEqual(result.results, expected); assert.equal(result.checked, 8); assert.deepEqual(requests, []); await page.close();
    }
    const before = readGeneratedZip(bytes), after = readGeneratedZip(Uint8Array.from(expected.at(-1).saved));
    for (const [name, value] of before) if (name !== "deck/z.xml") assert.deepEqual(after.get(name), value);
  }
});

test("fresh output retains mapped default text appearance with independently inspected native XML", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  let bytes = encodeStructure(textAppearanceFixture());
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, { projection: "content" });
    try {
      const table = normalizePresentation(s.document).slides[0].elements[0];
      assert.ok(table.rows.every(row => row.cells.every(cell => cell.content[0].runs[0].style.color === "CA6420" && cell.content[0].runs[0].style.fontFamily === "Arial")));
      bytes = (await generatePptx(s.document)).bytes;
      const xml = readGeneratedZip(bytes).get("ppt/slides/slide1.xml").toString();
      const checked = await page.evaluate(xml => {
        const doc = new DOMParser().parseFromString(xml, "application/xml"), ns = "http://schemas.openxmlformats.org/drawingml/2006/main", runs = [...doc.getElementsByTagNameNS(ns, "r")];
        for (const run of runs) {
          if (run.getElementsByTagNameNS(ns, "srgbClr")[0]?.getAttribute("val") !== "CA6420") throw new Error("Fresh color mismatch");
          if (run.getElementsByTagNameNS(ns, "latin")[0]?.getAttribute("typeface") !== "Arial") throw new Error("Fresh family mismatch");
        }
        return runs.length;
      }, xml);
      assert.equal(checked, 4);
    } finally { s.dispose(); }
  }
});
