import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { normalizePresentation, resolveTextFontSegments } from "../packages/core/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { nativeTableThemeSourceBytes } from "../tests/fixtures/native-table-font-themes.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const options = { projection: "content", textFontProfile: "script-families-v3", editSaveProfile: "text-runs-v3" };
const moduleUrl = source => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;

async function sourceRounds(fixture) {
  let bytes = fixture.bytes; const results = [];
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, options);
    try {
      const document = normalizePresentation(session.document), report = session.report;
      assert.equal(report.textFontProfile, "script-families-v3");
      for (const [index, control] of fixture.controls.entries()) {
        const runs = document.slides[index].elements.find(e => e.name === "ProbeTable").rows[0].cells[0].content[0].runs, edited = fixture.targets.includes(index + 1);
        assert.equal(runs.length, edited ? 2 : 1); assert.equal(runs[0].text, control.text); assert.equal(runs[0].style.language, control.language);
        for (const run of runs) assert.deepEqual(run.style.scriptFontFamilies, control.localOverride ? undefined : fixture.maps[control.kind]);
        assert.deepEqual([...new Set(resolveTextFontSegments(runs[0].text, runs[0].style, document.theme).map(s => s.family))], [control.targetFamily]);
        if (edited) {
          assert.equal(runs[1].text, `Ω${"X".repeat(round)}`); assert.equal(runs[1].style.language, "el-GR");
          assert.equal(resolveTextFontSegments("Ω", runs[1].style, document.theme)[0].family, control.localOverride ? fixture.localFonts.latin : fixture.maps[control.kind].Grek);
        }
        const rules = report.tableStyleInheritance.properties.filter(p => p.target.part === `/ppt/slides/slide${index + 1}.xml` && p.property.startsWith("run.scriptFontFamilies."));
        assert.equal(rules.length, control.localOverride ? 0 : runs.length * 8);
        for (const rule of rules) { assert.equal(rule.status, "mapped"); assert.equal(rule.value, fixture.maps[control.kind][rule.property.split(".").at(-1)]); assert.equal(rule.source.part, control.kind.startsWith("collection") ? "/ppt/tableStyles.xml" : "/" + fixture.themePart); }
      }
      const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), Buffer.from(bytes));
      const first = session.document.slides[fixture.targets[0]-1], original = first.elements.find(e => e.name === "ProbeTable"), wrong = structuredClone(original);
      wrong.rows[0].cells[0].content[0].runs[0].style.fontFamily = "Forbidden"; first.replaceElement(wrong.id, wrong);
      assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); first.replaceElement(original.id, original);
      for (const i of fixture.targets) {
        const slide = session.document.slides[i-1], table = slide.elements.find(e => e.name === "ProbeTable"), next = structuredClone(table);
        next.rows[0].cells[0].content[0].runs[1].text += "X"; slide.replaceElement(next.id, next);
      }
      const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready"); const saved = (await session.save(prepared.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(saved);
      assert.deepEqual([...after.keys()], [...before.keys()]);
      for (const [name, body] of before) {
        if (fixture.targets.some(i => name === `ppt/slides/slide${i}.xml`)) {
          const from = `<a:t>Ω${"X".repeat(round)}</a:t>`, to = `<a:t>Ω${"X".repeat(round+1)}</a:t>`;
          assert.equal(body.toString().split(from).length, 2); assert.equal(after.get(name).toString(), body.toString().replace(from, to));
        } else assert.deepEqual(after.get(name), body);
      }
      const candidate = await importPptx(saved, options);
      try { assert.deepEqual(normalizePresentation(candidate.document), normalizePresentation(session.document)); } finally { candidate.dispose(); }
      results.push({ document, report, savedBytes: [...saved] }); bytes = saved;
    } finally { session.dispose(); }
  }
  return results;
}

test("source table fonts follow rebound theme relationships, explicit collections and local slots across three consumers/saves", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const core = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8")), parser = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const filled of [false, true]) for (const swapped of [false, true]) for (const localOverrides of [false, true]) {
    const fixture = await nativeTableThemeSourceBytes({ filled, swapped, localOverrides }), expected = await sourceRounds(fixture), parts = readGeneratedZip(fixture.bytes);
    const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    const independent = await page.evaluate(({ parts, properties }) => {
      const A = "http://schemas.openxmlformats.org/drawingml/2006/main", P = "http://schemas.openxmlformats.org/presentationml/2006/main", R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships", REL = "http://schemas.openxmlformats.org/package/2006/relationships";
      const docs = Object.fromEntries(Object.entries(parts).map(([name, xml]) => ["/" + name, new DOMParser().parseFromString(xml, "application/xml")]));
      for (const doc of Object.values(docs)) if (doc.getElementsByTagName("parsererror").length) throw Error("Invalid independent XML");
      const relationship = (owner, kind, id) => {
        const split = owner.lastIndexOf("/"), relPart = owner.slice(0, split+1) + "_rels/" + owner.slice(split+1) + ".rels";
        const matches = [...docs[relPart].getElementsByTagNameNS(REL, "Relationship")].filter(n => n.getAttribute("Type").endsWith("/" + kind) && (id === undefined || n.getAttribute("Id") === id));
        if (matches.length !== 1 || matches[0].getAttribute("TargetMode") === "External") throw Error("Ambiguous/external independent relationship");
        const target = new URL(matches[0].getAttribute("Target"), "https://fixture.invalid" + owner).pathname;
        if (!docs[target]) throw Error("Missing independent target"); return target;
      };
      const select = source => { const doc = docs[source.part]; if (!doc) throw Error("Invented source part"); return doc.evaluate(source.path, doc, p => ({ a: A, p: P, r: R })[p] ?? doc.documentElement.lookupNamespaceURI(p), XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; };
      const styles = docs[relationship("/ppt/presentation.xml", "tableStyles")], rows = [];
      for (const slide of docs["/ppt/presentation.xml"].getElementsByTagNameNS(P, "sldId")) {
        const part = relationship("/ppt/presentation.xml", "slide", slide.getAttributeNS(R, "id")), layout = relationship(part, "slideLayout"), master = relationship(layout, "slideMaster"), theme = relationship(master, "theme");
        const frame = [...docs[part].getElementsByTagNameNS(P, "graphicFrame")].find(n => n.getElementsByTagNameNS(P, "cNvPr")[0].getAttribute("name") === "ProbeTable"), table = frame.getElementsByTagNameNS(A, "tbl")[0], id = table.getElementsByTagNameNS(A, "tableStyleId")[0].textContent;
        const style = [...styles.getElementsByTagNameNS(A, "tblStyle")].find(n => n.getAttribute("styleId") === id), text = style.getElementsByTagNameNS(A, "tcTxStyle")[0], reference = [...text.children].find(n => n.localName === "fontRef");
        const selected = reference ? docs[theme].getElementsByTagNameNS(A, reference.getAttribute("idx") + "Font")[0] : [...text.children].find(n => n.localName === "font");
        rows.push({ part, theme, styleId: id, map: Object.fromEntries([...selected.children].filter(n => n.localName === "font").map(n => [n.getAttribute("script"), n.getAttribute("typeface")])), localFonts: [...table.getElementsByTagNameNS(A, "rPr")].map(run => ["latin", "ea", "cs"].map(slot => run.getElementsByTagNameNS(A, slot)[0]?.getAttribute("typeface") ?? null)) });
      }
      const fonts = properties.filter(p => p.property.startsWith("run.scriptFontFamilies.") || (p.level === "run" && ["run.fontFamily", "run.eastAsianFontFamily", "run.complexScriptFontFamily"].includes(p.property)));
      for (const p of fonts.filter(p => p.status === "mapped")) {
        if (select(p.source)?.value !== p.value || !select(p.target)) throw Error("Independent source property mismatch: " + JSON.stringify({ property: p.property, source: p.source, actual: select(p.source)?.value, expected: p.value }));
        for (const dependency of p.resolvedFrom ?? []) if (!select(dependency)) throw Error("Invented dependency");
      }
      const probe = fonts.find(p => p.property.startsWith("run.scriptFontFamilies.") && p.status === "mapped"), node = select(probe.source); node.value = "Wrong";
      return { rows, checkedProperties: fonts.filter(p => p.status === "mapped").length, rejectsWrongSourceValue: node.value !== probe.value };
    }, { parts: Object.fromEntries([...parts].filter(([name]) => name.endsWith(".xml") || name.endsWith(".rels")).map(([name, body]) => [name, body.toString()])), properties: expected[0].report.tableStyleInheritance.properties });
    assert.equal(independent.rows.length, fixture.controls.length); assert.ok(independent.checkedProperties > 0); assert.equal(independent.rejectsWrongSourceValue, true);
    for (const [i, c] of fixture.controls.entries()) {
      const row = independent.rows[i]; assert.equal(row.theme, "/" + fixture.themePart); assert.equal(row.styleId, c.styleId); assert.deepEqual(row.map, fixture.maps[c.kind]);
      for (const fonts of row.localFonts) assert.deepEqual(fonts, c.localOverride ? [fixture.localFonts.latin, fixture.localFonts.ea, fixture.localFonts.cs] : [null, null, null]);
    }
    for (const format of ["esm", "global"]) {
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
      const actual = await page.evaluate(async ({ bytes, targets, options, core, parser }) => {
        const c = core ? await import(core) : PPTKit, p = parser ? await import(parser) : PPTKit, results = [];
        for (let round = 0; round < 3; round++) {
          const s = await p.importPptx(Uint8Array.from(bytes), options);
          try {
            const document = c.normalizePresentation(s.document), report = s.report, noOp = await s.prepareSave({ mode: "no-op" });
            if (noOp.status !== "ready") throw Error("Browser no-op not ready");
            const noOpBytes = [...(await s.save(noOp.plan)).bytes]; if (noOpBytes.length !== bytes.length || noOpBytes.some((v, i) => v !== bytes[i])) throw Error("Browser no-op differs");
            const first = s.document.slides[targets[0]-1], original = first.elements.find(e => e.name === "ProbeTable"), wrong = structuredClone(original); wrong.rows[0].cells[0].content[0].runs[0].style.fontFamily = "Forbidden"; first.replaceElement(wrong.id, wrong);
            if ((await s.prepareSave({ mode: "edit" })).status !== "blocked") throw Error("Browser font write admitted"); first.replaceElement(original.id, original);
            for (const i of targets) { const slide = s.document.slides[i-1], old = slide.elements.find(e => e.name === "ProbeTable"), next = structuredClone(old); next.rows[0].cells[0].content[0].runs[1].text += "X"; slide.replaceElement(next.id, next); }
            const ready = await s.prepareSave({ mode: "edit" }); if (ready.status !== "ready") throw Error("Browser scalar failed"); bytes = [...(await s.save(ready.plan)).bytes]; results.push({ document, report, savedBytes: bytes });
          } finally { s.dispose(); }
        }
        return results;
      }, { bytes: [...fixture.bytes], targets: fixture.targets, options, core: format === "esm" ? core : null, parser: format === "esm" ? parser : null });
      assert.deepEqual(actual, expected);
    }
    assert.deepEqual(requests, []); await page.close();
  }
});
