import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { supplementalFontFixture } from "../packages/pptx-parser/test/supplemental-font-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { mixedDeclaredSlotBytes } from "../tests/fixtures/declared-slot-font-import.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const options = { projection: "content", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" };
const moduleUrl = source => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
async function rounds(bytes) {
  const results = [];
  for (let index = 0; index < 3; index++) {
    const session = await importPptx(bytes, options);
    try {
      const document = normalizePresentation(session.document), report = session.report.textStyleInheritance, slide = session.document.slides[0], next = structuredClone(slide.elements[0]);
      next.text.content[0].runs[0].text = `Changed ${index} 中عאΩ`; slide.replaceElement(next.id, next);
      const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready"); bytes = (await session.save(prepared.plan)).bytes;
      results.push({ document, report, bytes: [...bytes] });
    } finally { session.dispose(); }
  }
  return results;
}

test("v3 both dialects/Node/ESM/global retain three full snapshots and independent mapped/blocked DOM sources", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const core = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8")), parser = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const strict of [false, true]) {
    const fixture = supplementalFontFixture({ strict, run: '<a:latin typeface="+mn-lt"/>' }), bytes = encodeStructure(fixture), expected = await rounds(bytes);
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
      const actual = await page.evaluate(async ({ core, parser, bytes, options, parts, namespace }) => {
        const c = core ? await import(core) : PPTKit, p = parser ? await import(parser) : PPTKit, results = [];
        for (let index = 0; index < 3; index++) {
          const session = await p.importPptx(Uint8Array.from(bytes), options);
          try {
            const document = c.normalizePresentation(session.document), report = session.report.textStyleInheritance, slide = session.document.slides[0], next = structuredClone(slide.elements[0]);
            next.text.content[0].runs[0].text = `Changed ${index} 中عאΩ`; slide.replaceElement(next.id, next);
            const prepared = await session.prepareSave({ mode: "edit" }); if (prepared.status !== "ready") throw Error("v3 candidate lost captured profile"); bytes = [...(await session.save(prepared.plan)).bytes]; results.push({ document, report, bytes });
          } finally { session.dispose(); }
        }
        const docs = Object.fromEntries(Object.entries(parts).filter(([, xml]) => typeof xml === "string" && xml.startsWith("<")).map(([name, xml]) => [`/${name}`, new DOMParser().parseFromString(xml, "application/xml")]));
        const select = source => { const doc = docs[source.part]; return doc.evaluate(source.path, doc, prefix => prefix === "a" ? namespace : doc.documentElement.lookupNamespaceURI(prefix), XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; };
        const sourceNames = { Hans: "Simplified Source", Hant: "Traditional Source", Jpan: "Japanese Source", Hang: "Korean Source", Arab: "Arabic Source", Hebr: "Hebrew Source", Grek: "Greek Source" };
        const rules = results[0].report.properties.filter(p => p.property.startsWith("run.scriptFontFamilies."));
        const matchesSource = property => { const attr = select(property.source), code = property.property.split(".").at(-1); return attr?.value === sourceNames[code] && attr.ownerElement.getAttribute("script") === code; };
        let blocked = 0;
        for (const property of rules) {
          if (!matchesSource(property) || !select(property.target)) throw Error("Actual source identity/value mismatch");
          for (const dependency of property.resolvedFrom) if (!select(dependency)) throw Error("Invented source dependency");
          const owner = select(property.target); if (owner.getElementsByTagNameNS(namespace, "ea").length || owner.getElementsByTagNameNS(namespace, "cs").length) throw Error("Fixture unexpectedly declares native slots");
          if (property.status === "unsupported") { if (property.value !== undefined) throw Error("Blocked source exposed a mapped value"); blocked++; }
          else if (property.property !== "run.scriptFontFamilies.Grek") throw Error("Missing bucket gained a rule");
        }
        const hant = rules.find(p => p.property.endsWith(".Hant")); select(hant.source).value = "Deliberately Wrong";
        return { results, checked: rules.length, blocked, rejectsWrongBlockedSource: !matchesSource(hant) };
      }, { core: format === "esm" ? core : null, parser: format === "esm" ? parser : null, bytes: [...bytes], options, parts: fixture.parts, namespace: fixture.A });
      assert.deepEqual(actual.results, expected); assert.equal(actual.checked, 14); assert.equal(actual.blocked, 12); assert.equal(actual.rejectsWrongBlockedSource, true); assert.deepEqual(requests, []); await page.close();
    }
  }
});

test("v3 mixed-owner notes/table/drawings retain 42 mapped/blocked rules and exact non-character bytes over three saves", async () => {
  let bytes = await mixedDeclaredSlotBytes();
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { ...options, notesTextProfile: "master-style-v1" });
    try {
      const properties = [...session.report.textStyleInheritance.properties, ...session.report.tableStyleInheritance.properties].filter(p => p.property.startsWith("run.scriptFontFamilies."));
      assert.equal(properties.length, 42); assert.equal(properties.filter(p => p.status === "unsupported").length, 36);
      assert.ok(properties.filter(p => p.status === "mapped").every(p => p.property.endsWith(".Grek")));
      const slide = session.document.slides[0];
      for (const element of [...slide.elements]) {
        const next = structuredClone(element), texts = next.type === "shape" ? [next.text.content] : next.type === "table" ? next.rows.flatMap(r => r.cells.map(c => c.content)) : [next.content];
        for (const paragraphs of texts) { assert.deepEqual(paragraphs[0].runs[0].style.scriptFontFamilies, { Grek: "Greek Source" }); paragraphs[0].runs[0].text = `Round${round} 中عאΩ`; }
        slide.replaceElement(next.id, next);
      }
      const notes = structuredClone(slide.notes); assert.deepEqual(notes[0].runs[0].style.scriptFontFamilies, { Grek: "Notes Greek Source" }); notes[0].runs[0].text = `Notes${round} 中عאΩ`; slide.setNotes(notes);
      const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready"); const saved = await session.save(prepared.plan);
      const before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes); assert.deepEqual([...after.keys()].sort(), [...before.keys()].sort());
      for (const [name, payload] of before) {
        const scalar = ["ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"].includes(name);
        assert.deepEqual(scalar ? after.get(name).toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>") : after.get(name), scalar ? payload.toString().replace(/<a:t>[^<]*<\/a:t>/g, "<a:t/>") : payload);
      }
      const fresh = await generatePptx(session.document); assert.ok(fresh.warnings.some(w => w.code === "script-font-map-materialized"));
      const freshSlide = readGeneratedZip(fresh.bytes).get("ppt/slides/slide1.xml").toString(); assert.ok(freshSlide.includes('typeface="Greek Source"')); assert.equal(freshSlide.includes('typeface="Traditional Source"'), false);
      bytes = saved.bytes;
    } finally { session.dispose(); }
  }
});

test("v3 browser capture/failure/cancel and restore preserve independent font-write guards", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const bytes = encodeStructure(supplementalFontFixture({ run: '<a:latin typeface="+mn-lt"/>' }));
  const result = await page.evaluate(async ({ input, options }) => {
    const bytes = Uint8Array.from(input), config = { ...options }, pending = PPTKit.importPptx(bytes, config); config.textFontProfile = "script-families-v2"; bytes.fill(0); const s = await pending;
    try {
      const slide = s.document.slides[0], old = slide.elements[0], next = structuredClone(old); next.text.content[0].runs[0].style.scriptFontFamilies.Hant = "Forbidden"; slide.replaceElement(old.id, next);
      const blocked = (await s.prepareSave({ mode: "edit" })).status; slide.replaceElement(old.id, old); const restored = (await s.prepareSave({ mode: "no-op" })).status;
      let limited, invalid, aborted;
      try { await PPTKit.importPptx(Uint8Array.from(input), { ...options, sessionLimits: { maxTextStyleDependencies: 1 } }); } catch (e) { limited = e.code; }
      try { await PPTKit.importPptx(Uint8Array.from(input), { ...options, projection: "structure" }); } catch (e) { invalid = e.name; }
      const controller = new AbortController(), operation = PPTKit.importPptx(Uint8Array.from(input), { ...options, signal: controller.signal }); queueMicrotask(() => controller.abort()); try { await operation; } catch (e) { aborted = e.code; }
      return { marker: s.report.textFontProfile, map: old.text.content[0].runs[0].style.scriptFontFamilies, blocked, restored, limited, invalid, aborted };
    } finally { s.dispose(); }
  }, { input: [...bytes], options });
  assert.deepEqual(result, { marker: "script-families-v3", map: { Grek: "Greek Source" }, blocked: "blocked", restored: "ready", limited: "import-limit-exceeded", invalid: "TypeError", aborted: "operation-aborted" });
});

test("native table controls retain v3 declared source rules for references and empty/filled explicit collections", async t => {
  const { nativeTableFontBytes } = await import("../tests/fixtures/native-table-fonts.mjs");
  const { resolveTextFontSegments } = await import("../packages/core/dist/index.js");
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  for (const variants of [undefined, ["collection-generic"]]) {
    const fixture = await nativeTableFontBytes(variants), parts = readGeneratedZip(fixture.bytes);
    const independent = await page.evaluate(({ slides, styles, theme }) => {
      const ns = "http://schemas.openxmlformats.org/drawingml/2006/main", parse = xml => new DOMParser().parseFromString(xml, "application/xml"), styleDoc = parse(styles), themeDoc = parse(theme);
      return slides.map(xml => {
        const doc = parse(xml), table = doc.getElementsByTagNameNS(ns, "tbl")[0], id = table.getElementsByTagNameNS(ns, "tableStyleId")[0].textContent;
        const style = [...styleDoc.getElementsByTagNameNS(ns, "tblStyle")].find(s => s.getAttribute("styleId") === id), text = style.getElementsByTagNameNS(ns, "tcTxStyle")[0];
        const ref = [...text.children].find(n => n.localName === "fontRef"), collection = ref ? themeDoc.getElementsByTagNameNS(ns, ref.getAttribute("idx") + "Font")[0] : [...text.children].find(n => n.localName === "font");
        const map = Object.fromEntries([...collection.children].filter(n => n.localName === "font").map(n => [n.getAttribute("script"), n.getAttribute("typeface")]));
        const run = table.getElementsByTagNameNS(ns, "rPr")[0];
        return { id, choice: ref ? ref.getAttribute("idx") : "collection", map, language: run.getAttribute("lang"), localFonts: ["latin", "ea", "cs"].map(n => run.getElementsByTagNameNS(ns, n).length) };
      });
    }, { slides: fixture.controls.map((_, i) => parts.get(`ppt/slides/slide${i + 1}.xml`).toString()), styles: parts.get("ppt/tableStyles.xml").toString(), theme: parts.get("ppt/theme/theme1.xml").toString() });
    const session = await importPptx(fixture.bytes, { projection: "content", textFontProfile: "script-families-v3" });
    try {
      const document = normalizePresentation(session.document);
      for (const [i, control] of fixture.controls.entries()) {
        const source = independent[i], table = document.slides[i].elements.find(e => e.name === "ProbeTable"), run = table.rows[0].cells[0].content[0].runs[0];
        assert.equal(source.id, control.styleId); assert.equal(source.choice, control.kind.startsWith("collection") ? "collection" : control.kind); assert.equal(source.language, control.language); assert.deepEqual(source.localFonts, [0, 0, 0]);
        assert.deepEqual(source.map, fixture.maps[control.kind]); assert.deepEqual(run.style.scriptFontFamilies, source.map);
        assert.deepEqual([...new Set(resolveTextFontSegments(run.text, run.style, document.theme).map(s => s.family))], [control.targetFamily]);
        const rules = session.report.tableStyleInheritance.properties.filter(p => p.target.part === `/ppt/slides/slide${i + 1}.xml` && p.property.startsWith("run.scriptFontFamilies."));
        assert.equal(rules.length, 8);
        for (const rule of rules) { assert.equal(rule.status, "mapped"); assert.equal(rule.value, source.map[rule.property.split(".").at(-1)]); assert.equal(rule.source.part, control.kind.startsWith("collection") ? "/ppt/tableStyles.xml" : "/ppt/theme/theme1.xml"); }
      }
      const before = normalizePresentation(session.document), prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready"); assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), Buffer.from(fixture.bytes)); assert.deepEqual(normalizePresentation(session.document), before);
    } finally { session.dispose(); }
  }
});

test("source-only native-observed two-run table topology retains both maps and exact scalar save boundaries", async () => {
  const { nativeTableFontBytes } = await import("../tests/fixtures/native-table-fonts.mjs");
  const { zip } = await import("../packages/pptx-parser/test/zip-fixture.mjs");
  for (const [variants, targets] of [[undefined, [3, 5, 6, 7, 12, 14, 15, 21, 25, 26]], [["collection-generic"], [3, 5, 6, 7, 8]]]) {
    const fixture = await nativeTableFontBytes(variants), parts = readGeneratedZip(fixture.bytes);
    // Named source patch reproduces the observed run topology, not a native application execution.
    for (const index of targets) {
      const name = `ppt/slides/slide${index}.xml`, control = fixture.controls[index - 1]; let count = 0;
      const source = parts.get(name).toString().replace(/<a:tbl>[\s\S]*?<\/a:tbl>/, table => table.replace(/<a:r><a:rPr\b[^>]*\/><a:t>[^<]*<\/a:t><\/a:r>/, run => {
        count++; const alternate = ["ar", "he"].includes(control.sampleId) ? "" : ` altLang="${control.language}"`;
        return run + `<a:r><a:rPr lang="el-GR"${alternate} sz="4800" dirty="0"/><a:t>Ω</a:t></a:r>`;
      }));
      assert.equal(count, 1); parts.set(name, Buffer.from(source));
    }
    let bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
    for (let round = 0; round < 3; round++) {
      const session = await importPptx(bytes, options);
      try {
        const before = normalizePresentation(session.document), noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready");
        assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), Buffer.from(bytes));
        for (const [i, control] of fixture.controls.entries()) {
          const table = before.slides[i].elements.find(e => e.name === "ProbeTable"), runs = table.rows[0].cells[0].content[0].runs, edited = targets.includes(i + 1);
          assert.equal(runs.length, edited ? 2 : 1); assert.equal(runs[0].text, control.text);
          for (const run of runs) assert.deepEqual(run.style.scriptFontFamilies, fixture.maps[control.kind]);
          const properties = session.report.tableStyleInheritance.properties.filter(p => p.target.part === `/ppt/slides/slide${i + 1}.xml` && p.property.startsWith("run.scriptFontFamilies."));
          assert.equal(properties.length, edited ? 16 : 8);
          assert.ok(properties.every(p => p.status === "mapped" && p.value === fixture.maps[control.kind][p.property.split(".").at(-1)]));
          if (edited) { assert.equal(runs[1].style.language, "el-GR"); assert.equal(runs[1].text, `Ω${"X".repeat(round)}`); }
        }
        const slide = session.document.slides[targets[0] - 1], table = slide.elements.find(e => e.name === "ProbeTable"), wrong = structuredClone(table);
        wrong.rows[0].cells[0].content[0].runs[1].style.scriptFontFamilies.Grek = "Forbidden"; slide.replaceElement(wrong.id, wrong);
        assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); slide.replaceElement(table.id, table);
        for (const index of targets) {
          const slide = session.document.slides[index - 1], table = slide.elements.find(e => e.name === "ProbeTable"), next = structuredClone(table);
          next.rows[0].cells[0].content[0].runs[1].text += "X"; slide.replaceElement(next.id, next);
        }
        const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready"); const saved = (await session.save(prepared.plan)).bytes, after = readGeneratedZip(saved), source = readGeneratedZip(bytes);
        assert.deepEqual([...source.keys()], [...after.keys()]);
        for (const [name, body] of source) {
          if (targets.some(i => name === `ppt/slides/slide${i}.xml`)) {
            const from = `<a:t>Ω${"X".repeat(round)}</a:t>`, to = `<a:t>Ω${"X".repeat(round + 1)}</a:t>`;
            assert.equal(body.toString().split(from).length, 2); assert.equal(after.get(name).toString(), body.toString().replace(from, to));
          } else assert.deepEqual(after.get(name), body);
        }
        bytes = saved;
      } finally { session.dispose(); }
    }
    const final = await importPptx(bytes, options);
    try {
      assert.equal(final.report.textFontProfile, "script-families-v3");
      for (const index of targets) assert.equal(final.document.slides[index - 1].elements.find(e => e.name === "ProbeTable").rows[0].cells[0].content[0].runs[1].text, "ΩXXX");
    } finally { final.dispose(); }
  }
});

test("native-observed table source style stages retain map/language and keep style writes blocked", async () => {
  const { nativeTableStyleSourceBytes } = await import("../tests/fixtures/native-table-font-styles.mjs");
  for (const filled of [false, true]) for (const stage of ["bold", "italic", "regular"]) {
    const fixture = await nativeTableStyleSourceBytes(stage, filled);
    let bytes = fixture.bytes;
    for (let round = 0; round < 3; round++) {
      const session = await importPptx(bytes, options);
      try {
        for (const i of fixture.targets) {
          const control = fixture.controls[i - 1], table = session.document.slides[i - 1].elements.find(e => e.name === "ProbeTable"), runs = table.rows[0].cells[0].content[0].runs;
          assert.equal(runs.length, 2); assert.equal(runs[0].text, control.text); assert.equal(runs[0].style.language, control.language);
          assert.equal(runs[1].text, `Ω${"X".repeat(round)}`); assert.equal(runs[1].style.language, "el-GR");
          for (const run of runs) { assert.deepEqual(run.style.scriptFontFamilies, fixture.maps[control.kind]); assert.equal(run.style.bold ?? false, stage === "bold"); assert.equal(run.style.italic ?? false, stage === "italic"); }
        }
        const prepared = await session.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready"); assert.deepEqual(Buffer.from((await session.save(prepared.plan)).bytes), Buffer.from(bytes));
        const first = session.document.slides[fixture.targets[0] - 1], original = first.elements.find(e => e.name === "ProbeTable"), wrong = structuredClone(original);
        wrong.rows[0].cells[0].content[0].runs[0].style.bold = stage !== "bold"; first.replaceElement(wrong.id, wrong);
        assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); first.replaceElement(original.id, original);
        for (const i of fixture.targets) {
          const slide = session.document.slides[i - 1], table = slide.elements.find(e => e.name === "ProbeTable"), next = structuredClone(table);
          next.rows[0].cells[0].content[0].runs[1].text += "X"; slide.replaceElement(next.id, next);
        }
        const edit = await session.prepareSave({ mode: "edit" }); assert.equal(edit.status, "ready"); const saved = (await session.save(edit.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(saved);
        assert.deepEqual([...after.keys()], [...before.keys()]);
        for (const [name, body] of before) {
          if (fixture.targets.some(i => name === `ppt/slides/slide${i}.xml`)) {
            const from = `<a:t>Ω${"X".repeat(round)}</a:t>`, to = `<a:t>Ω${"X".repeat(round + 1)}</a:t>`;
            assert.equal(body.toString().split(from).length, 2); assert.equal(after.get(name).toString(), body.toString().replace(from, to));
          } else assert.deepEqual(after.get(name), body);
        }
        bytes = saved;
      } finally { session.dispose(); }
    }
    const final = await importPptx(bytes, options);
    try { for (const i of fixture.targets) assert.equal(final.document.slides[i - 1].elements.find(e => e.name === "ProbeTable").rows[0].cells[0].content[0].runs[1].text, "ΩXXX"); }
    finally { final.dispose(); }
  }
});
