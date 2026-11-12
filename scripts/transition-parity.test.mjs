import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

function document() {
  const p = createPresentation();
  for (const transition of [undefined, { type: "fade", durationMs: 1234, onClick: false, afterMs: 2345 }, { type: "fade", durationMs: 0, afterMs: 0 }, { type: "fade", durationMs: 4294967295, onClick: false }]) {
    p.addSlide({ ...(transition ? { transition } : {}) }).addElement({ type: "text", box: { x: 10, y: 10, width: 200, height: 40 }, content: "Fade control" });
  }
  return p;
}

test("fade timing survives Layout and three fresh rounds; static SVG is unchanged", async () => {
  const p = document(), expected = normalizePresentation(p).slides.map(s => s.transition);
  const layout = resolveLayout(p); assert.deepEqual(layout.slides.map(s => s.transition), expected);
  layout.slides[1].transition.durationMs = 99; assert.equal(p.slides[1].transition.durationMs, 1234);
  const before = await renderPresentationToSvg(p);
  let bytes = (await generatePptx(p)).bytes;
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
    assert.deepEqual(normalizePresentation(s.document).slides.map(s => s.transition), expected);
    const noOp = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(noOp.plan)).bytes, bytes);
    const old = s.document.slides[1].transition; s.document.slides[1].setTransition({ ...old, durationMs: 1 });
    assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); s.document.slides[1].setTransition(old);
    bytes = (await generatePptx(s.document)).bytes; s.dispose();
  }
  for (const s of p.slides) s.setTransition();
  assert.deepEqual((await renderPresentationToSvg(p)).slides, before.slides);
  const parts = readGeneratedZip(bytes);
  assert.doesNotMatch(parts.get("ppt/slides/slide1.xml").toString(), /<p:transition/);
  assert.match(parts.get("ppt/slides/slide2.xml").toString(), /p14:dur="1234"[^>]*advClick="0"[^>]*advTm="2345"/);
});

test("independent browser DOM verifies native timing and standalone Parser agrees with Node", async t => {
  const bytes = (await generatePptx(document())).bytes, parts = readGeneratedZip(bytes);
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(); let requests = 0;
  await page.route("**/*", route => { requests++; return route.abort(); });
  await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  const result = await page.evaluate(async ({ bytes, xml }) => {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const n = doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "transition")[0];
    const s = await PPTKit.importPptx(new Uint8Array(bytes), { projection: "content" });
    const transitions = s.document.slides.map(s => s.transition ?? null); s.dispose();
    return { transitions, duration: n.getAttributeNS("http://schemas.microsoft.com/office/powerpoint/2010/main", "dur"), click: n.getAttribute("advClick"), advance: n.getAttribute("advTm"), effect: n.firstElementChild.localName, errors: doc.querySelectorAll("parsererror").length };
  }, { bytes: [...bytes], xml: parts.get("ppt/slides/slide2.xml").toString() });
  assert.equal(result.duration, "1234"); assert.equal(result.advance, "2345"); assert.equal(result.click, "0"); assert.equal(result.effect, "fade"); assert.equal(result.errors, 0);
  assert.deepEqual(result.transitions, normalizePresentation(document()).slides.map(s => s.transition ?? null)); assert.equal(requests, 0);
});

test("unrelated text preservation retains exact known and unknown transition markup", async () => {
  const p = document(), original = (await generatePptx(p)).bytes;
  for (const unknown of [false, true]) {
  const parts = readGeneratedZip(original);
  if (unknown) parts.set("ppt/slides/slide2.xml", Buffer.from(parts.get("ppt/slides/slide2.xml").toString().replace("<p:fade/>", '<p:futureEffect><p:opaque/></p:futureEffect>')));
  const bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
  const s = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
  const slide = s.document.slides[1], element = structuredClone(slide.elements[0]);
  element.content[0].runs[0].text = "Edited control"; slide.removeElement(element.id); slide.insertElement(0, element);
  const prep = await s.prepareSave({ mode: "edit" }); assert.equal(prep.status, "ready");
  const output = readGeneratedZip((await s.save(prep.plan)).bytes);
  const transition = xml => xml.toString().match(/<p:transition[\s\S]*?<\/p:transition>/)[0];
  assert.equal(transition(output.get("ppt/slides/slide2.xml")), transition(readGeneratedZip(bytes).get("ppt/slides/slide2.xml"))); s.dispose();
  }
});


test("all eight side transitions retain direction over three rounds with independent DOM native attributes", async t => {
  const p = createPresentation(), cases = [];
  for (const type of ["push", "wipe"]) for (const [direction, native] of [["left", "l"], ["right", "r"], ["up", "u"], ["down", "d"]]) {
    const transition = { type, direction, durationMs: 765, onClick: false, afterMs: 987 }; cases.push({ transition, native });
    p.addSlide({ transition }).addElement({ type: "text", box: { x: 0, y: 0, width: 100, height: 40 }, content: "Side control" });
  }
  let bytes = (await generatePptx(p)).bytes;
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, { projection: "content", editSaveProfile: "text-runs-v3" });
    assert.deepEqual(s.document.slides.map(x => x.transition), cases.map(x => x.transition));
    const prep = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(prep.plan)).bytes, bytes);
    bytes = (await generatePptx(s.document)).bytes; s.dispose();
  }
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage();
  const parts = readGeneratedZip(bytes), xml = cases.map((_,i) => parts.get(`ppt/slides/slide${i+1}.xml`).toString());
  const inspected = await page.evaluate(xml => xml.map(text => {
    const d = new DOMParser().parseFromString(text, "application/xml"), n = d.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "transition")[0];
    return [n.firstElementChild.localName, n.firstElementChild.getAttribute("dir"), n.getAttributeNS("http://schemas.microsoft.com/office/powerpoint/2010/main", "dur")];
  }), xml);
  assert.deepEqual(inspected, cases.map(x => [x.transition.type, x.native, "765"]));
  const rendered = await renderPresentationToSvg(p); for (const s of p.slides) s.setTransition(); assert.deepEqual((await renderPresentationToSvg(p)).slides, rendered.slides);
});
