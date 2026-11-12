import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { supplementalFontFixture } from "../packages/pptx-parser/test/supplemental-font-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { mixedSupplementalBytes } from "../tests/fixtures/supplemental-font-import.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const options = { projection: "content", textFontProfile: "script-families-v2", editSaveProfile: "text-runs-v3" };
const moduleUrl = source => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
async function rounds(bytes) {
  const results = [];
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, options);
    try {
      const document = normalizePresentation(session.document), report = session.report.textStyleInheritance, slide = session.document.slides[0], next = structuredClone(slide.elements[0]);
      next.text.content[0].runs[0].text = `Changed ${round} 中عאΩ`; slide.replaceElement(next.id, next);
      const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); bytes = (await session.save(plan.plan)).bytes;
      results.push({ document, report, bytes: [...bytes] });
    } finally { session.dispose(); }
  }
  return results;
}

test("v2 Node/ESM/global three scalar saves retain exact map/provenance/bytes; independent DOM validates every source path", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const core = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8")), parser = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const strict of [false, true]) {
    const fixture = supplementalFontFixture({ strict, emptySlots: true }), bytes = encodeStructure(fixture), expected = await rounds(bytes);
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
      const actual = await page.evaluate(async ({ core, parser, bytes, parts, namespace, options }) => {
        const c = core ? await import(core) : PPTKit, p = parser ? await import(parser) : PPTKit, results = [];
        for (let round = 0; round < 3; round++) {
          const session = await p.importPptx(Uint8Array.from(bytes), options);
          try {
            const document = c.normalizePresentation(session.document), report = session.report.textStyleInheritance, slide = session.document.slides[0], next = structuredClone(slide.elements[0]);
            next.text.content[0].runs[0].text = `Changed ${round} 中عאΩ`; slide.replaceElement(next.id, next);
            const plan = await session.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw new Error("Candidate did not retain v2"); bytes = [...(await session.save(plan.plan)).bytes]; results.push({ document, report, bytes });
          } finally { session.dispose(); }
        }
        const docs = Object.fromEntries(Object.entries(parts).filter(([, xml]) => typeof xml === "string" && xml.startsWith("<")).map(([name, xml]) => [`/${name}`, new DOMParser().parseFromString(xml, "application/xml")]));
        const select = source => { const doc = docs[source.part]; return doc.evaluate(source.path, doc, prefix => prefix === "a" ? namespace : doc.documentElement.lookupNamespaceURI(prefix), XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; };
        let checked = 0;
        for (const property of results[0].report.properties) if (property.property.startsWith("run.scriptFontFamilies.")) {
          const node = select(property.source), key = property.property.split(".").at(-1);
          if (node?.value !== property.value || node.ownerElement.getAttribute("script") !== key) throw new Error("Source map key/value does not match actual DOM");
          for (const dependency of property.resolvedFrom) if (!select(dependency)) throw new Error("Missing actual dependency node");
          if (!select(property.target)) throw new Error("Missing actual target node"); checked++;
        }
        // Negative control independent of Parser and the new selector.
        const theme = docs["/deck/theme.xml"], node = [...theme.getElementsByTagNameNS(namespace, "font")].find(n => n.getAttribute("script") === "Hant"); node.setAttribute("typeface", "Deliberately Wrong");
        const property = results[0].report.properties.find(p => p.property === "run.scriptFontFamilies.Hant"); const rejectsMutation = select(property.source).value !== property.value;
        return { results, checked, rejectsMutation };
      }, { core: format === "esm" ? core : null, parser: format === "esm" ? parser : null, bytes: [...bytes], parts: fixture.parts, namespace: fixture.A, options });
      assert.deepEqual(actual.results, expected); assert.equal(actual.checked, 14); assert.equal(actual.rejectsMutation, true); assert.deepEqual(requests, []); await page.close();
    }
  }
});

test("mixed text/shape/table/notes three source saves retain all rules and unrelated bytes; fresh materialization is explicit", async t => {
  let bytes = await mixedSupplementalBytes(); const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, options);
    try {
      const slide = session.document.slides[0], initial = normalizePresentation(session.document);
      assert.equal(slide.elements[0].content[0].runs.length, 1); assert.equal(slide.elements[0].content[0].runs[0].style.scriptFontFamilies.Arab, "Arabic Source");
      for (const element of [...slide.elements]) {
        const next = structuredClone(element), content = next.type === "shape" ? next.text.content : next.type === "table" ? next.rows[0].cells[0].content : next.content;
        content[0].runs[0].text = `Round${round} 中عא`; slide.replaceElement(element.id, next);
      }
      const notes = structuredClone(slide.notes); notes[0].runs[0].text = `Notes${round} 中ع`; slide.setNotes(notes);
      const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); const saved = await session.save(plan.plan), before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes);
      assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort());
      for (const [name, payload] of before) {
        if (!["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"].includes(name)) assert.deepEqual(after.get(name), payload);
        else assert.equal(after.get(name).toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>"), payload.toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>"));
      }
      assert.equal(initial.slides[0].elements[0].content[0].runs[0].text, round === 0 ? "Latin中文العربية" : `Round${round - 1} 中عא`);
      const fresh = await generatePptx(session.document); assert.ok(fresh.warnings.some(w => w.code === "script-font-map-materialized"));
      const actual = await page.evaluate(xml => {
        const doc = new DOMParser().parseFromString(xml, "application/xml"), ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const shape = [...doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "sp")][0];
        return [...shape.getElementsByTagNameNS(ns, "r")].map(r => [r.getElementsByTagNameNS(ns, "t")[0].textContent, ...["latin", "ea", "cs"].map(n => r.getElementsByTagNameNS(ns, n)[0].getAttribute("typeface"))]);
      }, readGeneratedZip(fresh.bytes).get("ppt/slides/slide1.xml").toString());
      assert.deepEqual(actual, [[`Round${round} `, "Body Face"], ["中", "Traditional Source"], ["ع", "Arabic Source"], ["א", "Hebrew Source"]].map(([text, family]) => [text, family, family, family]));
      assert.ok((await renderPresentationToSvg(session.document)).slides[0].svg.includes("Arabic Source"));
      bytes = saved.bytes;
    } finally { session.dispose(); }
  }
});

test("browser v2 validates options/budgets/cancellation and restores map edit eligibility without new permissions", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(); for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const bytes = encodeStructure(supplementalFontFixture());
  const actual = await page.evaluate(async ({ input, options }) => {
    const bytes = Uint8Array.from(input), config = { ...options }, pending = PPTKit.importPptx(bytes, config); config.textFontProfile = undefined; bytes.fill(0);
    const s = await pending; let limited, aborted, invalid;
    try {
      const slide = s.document.slides[0], old = slide.elements[0], next = structuredClone(old); next.text.content[0].runs[0].style.scriptFontFamilies.Arab = "Wrong"; slide.replaceElement(old.id, next);
      const blocked = (await s.prepareSave({ mode: "edit" })).status; slide.replaceElement(old.id, old); const restored = (await s.prepareSave({ mode: "no-op" })).status;
      try { await PPTKit.importPptx(Uint8Array.from(input), { ...options, sessionLimits: { maxTextStyleDependencies: 1 } }); } catch (e) { limited = e.code; }
      try { await PPTKit.importPptx(Uint8Array.from(input), { ...options, projection: "structure" }); } catch (e) { invalid = e.name; }
      const c = new AbortController(), operation = PPTKit.importPptx(Uint8Array.from(input), { ...options, signal: c.signal }); queueMicrotask(() => c.abort()); try { await operation; } catch (e) { aborted = e.code; }
      const retry = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content", textFontProfile: "script-families-v2" });
      const page = retry.document.slides[0], e = structuredClone(page.elements[0]); e.text.content[0].runs[0].text = "No writer permission"; page.replaceElement(e.id, e); const permission = (await retry.prepareSave({ mode: "edit" })).status; retry.dispose();
      return { marker: s.report.textFontProfile, blocked, restored, limited, aborted, invalid, permission };
    } finally { s.dispose(); }
  }, { input: [...bytes], options });
  assert.deepEqual(actual, { marker: "script-families-v2", blocked: "blocked", restored: "ready", limited: "import-limit-exceeded", aborted: "operation-aborted", invalid: "TypeError", permission: "blocked" });
});
