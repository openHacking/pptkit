import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { notesStyleFixture, addSecondNotes } from "../packages/pptx-parser/test/notes-style-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const options = { projection: "content", textFontProfile: "script-families-v2", notesTextProfile: "master-style-v1", editSaveProfile: "text-runs-v3" };
const moduleUrl = source => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
async function rounds(bytes) {
  const results = [];
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, options);
    try {
      const document = normalizePresentation(session.document), report = session.report.textStyleInheritance;
      for (const slide of session.document.slides) { const notes = structuredClone(slide.notes); notes[0].runs[0].text = `Changed${round} A中عא`; slide.setNotes(notes); }
      const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); bytes = (await session.save(plan.plan)).bytes; results.push({ document, report, bytes: [...bytes] });
    } finally { session.dispose(); }
  }
  return results;
}

test("notes inheritance Node/ESM/global capture and three saves agree; independent DOM checks actual native selection paths", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const core = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8")), parser = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const strict of [false, true]) {
    const f = addSecondNotes(notesStyleFixture({ strict }), true), bytes = encodeStructure(f), expected = await rounds(bytes);
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
      const actual = await page.evaluate(async ({ core, parser, bytes, parts, namespace, options }) => {
        const c = core ? await import(core) : PPTKit, p = parser ? await import(parser) : PPTKit, results = [];
        for (let round = 0; round < 3; round++) {
          const session = await p.importPptx(Uint8Array.from(bytes), options);
          try {
            const document = c.normalizePresentation(session.document), report = session.report.textStyleInheritance;
            for (const slide of session.document.slides) { const notes = structuredClone(slide.notes); notes[0].runs[0].text = `Changed${round} A中عא`; slide.setNotes(notes); }
            const plan = await session.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw new Error("Notes profile not retained during full candidate validation"); bytes = [...(await session.save(plan.plan)).bytes]; results.push({ document, report, bytes });
          } finally { session.dispose(); }
        }
        const docs = Object.fromEntries(Object.entries(parts).filter(([, xml]) => typeof xml === "string" && xml.startsWith("<")).map(([name, xml]) => [`/${name}`, new DOMParser().parseFromString(xml, "application/xml")]));
        const select = source => { const doc = docs[source.part]; return doc.evaluate(source.path, doc, prefix => prefix === "a" ? namespace : doc.documentElement.lookupNamespaceURI(prefix), XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; };
        let maps = 0, fonts = 0, placeholderSizes = 0;
        for (const property of results[0].report.properties.filter(p => p.target.part.includes("/notes/"))) {
          if (!select(property.target) || !select(property.source)) throw new Error("Missing actual notes source or target");
          for (const dependency of property.resolvedFrom ?? []) if (!select(dependency)) throw new Error("Missing actual notes dependency");
          if (property.property.startsWith("run.scriptFontFamilies.")) {
            const node = select(property.source); if (node.value !== property.value || node.ownerElement.getAttribute("script") !== property.property.split(".").at(-1)) throw new Error("Wrong supplemental source"); maps++;
          }
          if (["run.fontFamily", "run.eastAsianFontFamily", "run.complexScriptFontFamily"].includes(property.property)) {
            if (!property.resolvedFrom.some(d => d.part.includes("notes-theme") && select(d).value === property.value)) throw new Error("Font did not come from actual notes theme"); fonts++;
          }
          if (property.property === "run.fontSize") {
            if (select(property.source).value !== "2200" || property.value !== 22 || property.level !== "notes-master-placeholder") throw new Error("Wrong inherited placeholder size");
            const types = property.resolvedFrom.map(select); if (!types.every(n => n.value === "body")) throw new Error("Notes must match by type");
            const indices = types.map(n => n.ownerElement.getAttribute("idx")).sort(); if (indices.join(",") !== "1,99") throw new Error("Fixture did not distinguish type from idx"); placeholderSizes++;
          }
        }
        // Read the actual notes-master clrMap and actual notes theme independently.
        for (const suffix of ["", "2"]) {
          const master = docs[`/deck/notes-master${suffix}.xml`], map = master.getElementsByTagNameNS(master.documentElement.namespaceURI, "clrMap")[0].getAttribute("tx1");
          const theme = docs[`/deck/notes-theme${suffix}.xml`], color = theme.getElementsByTagNameNS(namespace, map)[0].getElementsByTagNameNS(namespace, "srgbClr")[0].getAttribute("val");
          if (color !== "204060") throw new Error("Wrong independent color source");
        }
        const mapProperty = results[0].report.properties.find(p => p.property === "run.scriptFontFamilies.Hant"); select(mapProperty.source).value = "Wrong Notes Family";
        return { results, maps, fonts, placeholderSizes, rejectsMutation: select(mapProperty.source).value !== mapProperty.value };
      }, { core: format === "esm" ? core : null, parser: format === "esm" ? parser : null, bytes: [...bytes], parts: f.parts, namespace: f.A, options });
      assert.deepEqual(actual.results, expected); assert.deepEqual([actual.maps, actual.fonts, actual.placeholderSizes], [6, 6, 2]); assert.equal(actual.rejectsMutation, true); assert.deepEqual(requests, []); await page.close();
    }
  }
});

test("notes/drawing mixed scalar saves preserve masters/themes exactly; fresh notes materialize fonts with explicit rule loss", async t => {
  let bytes = Uint8Array.from(encodeStructure(addSecondNotes(notesStyleFixture()))); const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, options);
    try {
      const baseline = await session.prepareSave({ mode: "no-op" }); assert.equal(baseline.status, "ready"); assert.deepEqual((await session.save(baseline.plan)).bytes, bytes);
      for (const slide of session.document.slides) { const notes = structuredClone(slide.notes); notes[0].runs[0].text = `Round${round} A中عא`; slide.setNotes(notes); }
      const slide = session.document.slides[0], drawing = structuredClone(slide.elements[0]); (drawing.type === "text" ? drawing.content : drawing.text.content)[0].runs[0].text = `Drawing${round}`; slide.replaceElement(drawing.id, drawing);
      const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); const saved = await session.save(plan.plan), before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes);
      assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort());
      for (const [name, payload] of before) if (["deck/notes/z.xml", "deck/notes/a.xml", "deck/z.xml"].includes(name)) assert.equal(after.get(name).toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>"), payload.toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>")); else assert.deepEqual(after.get(name), payload);
      const fresh = await generatePptx(session.document); assert.ok(fresh.warnings.some(w => w.code === "script-font-map-materialized" && w.path === "notes"));
      const native = await page.evaluate(xml => {
        const doc = new DOMParser().parseFromString(xml, "application/xml"), a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        return [...doc.getElementsByTagNameNS(a, "r")].map(r => ({ text: r.getElementsByTagNameNS(a, "t")[0].textContent, size: r.getElementsByTagNameNS(a, "rPr")[0].getAttribute("sz"), color: r.getElementsByTagNameNS(a, "srgbClr")[0].getAttribute("val"), fonts: ["latin", "ea", "cs"].map(n => r.getElementsByTagNameNS(a, n)[0].getAttribute("typeface")) }));
      }, readGeneratedZip(fresh.bytes).get("ppt/notesSlides/notesSlide1.xml").toString());
      assert.deepEqual(native, [[`Round${round} A`, "Notes Latin"], ["中", "Notes Traditional"], ["ع", "Notes Arabic"], ["א", "Notes Hebrew"]].map(([text, font]) => ({ text, size: "2200", color: "204060", fonts: Array(3).fill(font) })));
      const imported = await importPptx(fresh.bytes, options); try { assert.ok(imported.document.slides[0].notes[0].runs.every(r => !r.style.scriptFontFamilies)); assert.equal(imported.document.slides[0].notes[0].runs.map(r => r.text).join(""), `Round${round} A中عא`); } finally { imported.dispose(); }
      bytes = saved.bytes;
    } finally { session.dispose(); }
  }
});

test("browser notes option capture, invalid mode, dependency limit, cancellation and style restoration remain independent", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(); for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const result = await page.evaluate(async ({ bytes, options }) => {
    const data = Uint8Array.from(bytes), config = { ...options }, pending = PPTKit.importPptx(data, config); data.fill(0); config.notesTextProfile = undefined;
    const session = await pending; let invalid, limited, aborted;
    try {
      const slide = session.document.slides[0], original = structuredClone(slide.notes), bad = structuredClone(original); bad[0].runs[0].style.fontSize = 30; slide.setNotes(bad); const blocked = (await session.prepareSave({ mode: "edit" })).status; slide.setNotes(original); const restored = (await session.prepareSave({ mode: "no-op" })).status;
      try { await PPTKit.importPptx(Uint8Array.from(bytes), { ...options, projection: "structure" }); } catch (e) { invalid = e.name; }
      try { await PPTKit.importPptx(Uint8Array.from(bytes), { ...options, sessionLimits: { maxTextStyleDependencies: 1 } }); } catch (e) { limited = e.code; }
      const c = new AbortController(), operation = PPTKit.importPptx(Uint8Array.from(bytes), { ...options, signal: c.signal }); queueMicrotask(() => c.abort()); try { await operation; } catch (e) { aborted = e.code; }
      const old = await PPTKit.importPptx(Uint8Array.from(bytes), { projection: "content", textFontProfile: "script-families-v2" }); const oldProfile = Object.hasOwn(old.report, "notesTextProfile"), oldFont = old.document.slides[0].notes[0].runs[0].style.fontFamily; old.dispose();
      return { marker: session.report.notesTextProfile, blocked, restored, invalid, limited, aborted, oldProfile, oldFont };
    } finally { session.dispose(); }
  }, { bytes: [...encodeStructure(notesStyleFixture())], options });
  assert.deepEqual(result, { marker: "master-style-v1", blocked: "blocked", restored: "ready", invalid: "TypeError", limited: "import-limit-exceeded", aborted: "operation-aborted", oldProfile: false, oldFont: undefined });
});
