import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx, inspectPptx } from "../packages/pptx-parser/dist/index.js";
import { authoredNotesPages } from "../tests/fixtures/authored-notes-pages.mjs";
import { authoredNotesPageAudit } from "../tests/support/authored-notes-page-audit.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

test("fresh default notes pages contain undistorted images, separate bodies and copied owners across sizes and three projections", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  await page.addScriptTag({ content: "globalThis.authoredNotesPages = " + authoredNotesPages.toString() });
  for (const [name, size] of [["wide", { width: 720, height: 260 }], ["portrait", { width: 360, height: 720 }], ["square", { width: 500, height: 500 }], ["thin", { width: 4032, height: 72 }]]) {
    let document = authoredNotesPages(createPresentation, size, name);
    const expected = (await generatePptx(document)).bytes;
    const browserBytes = await page.evaluate(async ({ size, name }) => [...(await PPTKit.generatePptx(authoredNotesPages(PPTKit.createPresentation, size, name))).bytes], { size, name });
    assert.deepEqual(browserBytes, [...expected]);
    let first;
    for (let round = 0; round < 3; round++) {
      const { bytes } = await generatePptx(document), parts = readGeneratedZip(bytes);
      const actual = await page.evaluate(authoredNotesPageAudit, Object.fromEntries([...parts].map(([n, b]) => [n, b.toString()])));
      assert.equal(actual.rows.length, 4);
      assert.deepEqual(actual.rows.map(r => r.text), [`Fresh notes ${name} ONE`, `Fresh notes ${name} COPY`, null, ""]);
      assert.deepEqual(actual.rows.map(r => r.bodyPresent), [true, true, false, true]);
      assert.notEqual(actual.rows[0].note, actual.rows[1].note);
      if (first) assert.deepEqual(actual, first); else first = actual;
      const session = await importPptx(bytes, { projection: "content" });
      try {
        const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready");
        assert.deepEqual((await session.save(plan.plan)).bytes, bytes);
        document = session.document;
      } finally { session.dispose(); }
    }
    // Actual geometry mutation must fail even when body text is unchanged.
    const broken = new Map(readGeneratedZip(expected)), master = "ppt/notesMasters/notesMaster1.xml";
    broken.set(master, Buffer.from(broken.get(master).toString().replace('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>', "")));
    await assert.rejects(page.evaluate(authoredNotesPageAudit, Object.fromEntries([...broken].map(([n, b]) => [n, b.toString()]))), /Missing\/ambiguous/);
    for (const mutation of ["shift", "undersize"]) {
      const edited = new Map(readGeneratedZip(expected));
      const xml = await page.evaluate(({ xml, mutation }) => {
        const p = "http://schemas.openxmlformats.org/presentationml/2006/main", a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const doc = new DOMParser().parseFromString(xml, "application/xml");
        const shape = [...doc.getElementsByTagNameNS(p, "sp")].find(s => [...s.getElementsByTagNameNS(p, "ph")].some(ph => ph.getAttribute("type") === "sldImg"));
        const off = shape.getElementsByTagNameNS(a, "off")[0], ext = shape.getElementsByTagNameNS(a, "ext")[0];
        if (mutation === "shift") off.setAttribute("y", String(Number(off.getAttribute("y")) + 12700));
        else {
          const width = Math.floor(Number(ext.getAttribute("cx")) / 2), height = Math.floor(Number(ext.getAttribute("cy")) / 2);
          ext.setAttribute("cx", String(width)); ext.setAttribute("cy", String(height)); off.setAttribute("x", String(Math.round((6858000 - width) / 2))); off.setAttribute("y", String(Math.round((360 * 12700 - height) / 2)));
        }
        return new XMLSerializer().serializeToString(doc);
      }, { xml: edited.get(master).toString(), mutation });
      edited.set(master, Buffer.from(xml));
      await assert.rejects(page.evaluate(authoredNotesPageAudit, Object.fromEntries([...edited].map(([n, b]) => [n, b.toString()]))), mutation === "shift" ? /vertically centered/ : /fit frame/);
    }
  }
  assert.deepEqual(requests, []);
});

test("slide-size limits preserve valid notes geometry and reject invalid dimensions across Node and browser", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ["core", "layout", "pptx-exporter"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  for (const [width, height, accepted] of [[72, 4032, true], [4032, 72, true], [72 - 0.4 / 12700, 72, true], [4032 + 0.4 / 12700, 72, true], [72 - 0.6 / 12700, 72, false], [4032 + 0.6 / 12700, 72, false], [1e-300, 72, false], [Number.MAX_VALUE, 72, false]]) {
    const doc = createPresentation({ size: { width, height } }); doc.addSlide();
    const make = () => generatePptx(doc);
    if (!accepted) await assert.rejects(make, RangeError);
    const result = await page.evaluate(async ({ width, height }) => { const d = PPTKit.createPresentation({ size: { width, height } }); d.addSlide(); try { return { bytes: [...(await PPTKit.generatePptx(d)).bytes] }; } catch (error) { return { name: error.name, message: error.message }; } }, { width, height });
    if (!accepted) { assert.equal(result.name, "RangeError"); assert.match(result.message, /914400–51206400 EMU/); continue; }
    const bytes = (await make()).bytes; assert.deepEqual(result.bytes, [...bytes]);
    const inspection = await inspectPptx(bytes); assert.equal(inspection.slideSize.widthEmu, Math.round(width * 12700)); assert.equal(inspection.slideSize.heightEmu, Math.round(height * 12700));
    const parts = readGeneratedZip(bytes); await page.evaluate(authoredNotesPageAudit, Object.fromEntries([...parts].map(([name, data]) => [name, data.toString()])));
  }
  assert.deepEqual(requests, []);
});
