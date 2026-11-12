import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { importPptx, inspectPptxTiming } from "../packages/pptx-parser/dist/index.js";
import { paragraphProfile, paragraphTarget, paragraphBytes, editParagraph } from "../packages/pptx-parser/test/paragraph-target-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
const url = source => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;

test("paragraph-preserving saves agree in Node/ESM/global through three rounds with independent DOM target and paragraph checks", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const esm = url(await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")), global = await readFile("packages/pptx-parser/dist/browser/global.js", "utf8");
  for (const strict of [false, true]) {
    const input = paragraphBytes({ strict }), expected = []; let source = input;
    for (let i = 0; i < 3; i++) {
      const session = await importPptx(source, paragraphProfile); editParagraph(session, `Round ${i}<&>👩‍💻é`, 2);
      const prepared = await session.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready");
      source = (await session.save(prepared.plan)).bytes; expected.push([...source]); session.dispose();
    }
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") await page.addScriptTag({ content: global });
      const result = await page.evaluate(async ({ bytes, profile, esm }) => {
        const api = esm ? await import(esm) : PPTKit, outputs = []; let source = Uint8Array.from(bytes);
        const baseline = await api.inspectPptxTiming(source);
        for (let i = 0; i < 3; i++) {
          const session = await api.importPptx(source, profile), slide = session.document.slides[0], element = structuredClone(slide.elements[0]);
          element.content[2].runs[0].text = `Round ${i}<&>👩‍💻é`; slide.replaceElement(element.id, element);
          const prepared = await session.prepareSave({ mode: "edit" }); if (prepared.status !== "ready") throw new Error(JSON.stringify(prepared.report));
          source = (await session.save(prepared.plan)).bytes;
          if (JSON.stringify(await api.inspectPptxTiming(source)) !== JSON.stringify(baseline)) throw new Error("Timing changed");
          outputs.push([...source]); session.dispose();
        }
        return outputs;
      }, { bytes: [...input], profile: paragraphProfile, esm: format === "esm" ? esm : null });
      assert.deepEqual(result, expected);
      const sourceXml = readGeneratedZip(input).get("deck/z.xml").toString();
      const xmls = result.map(bytes => readGeneratedZip(Uint8Array.from(bytes)).get("deck/z.xml").toString());
      const verified = await page.evaluate(({ sourceXml, xmls }) => {
        const parse = xml => new DOMParser().parseFromString(xml, "application/xml"), before = parse(sourceXml), p = before.documentElement.namespaceURI;
        const a = before.getElementsByTagNameNS(p, "txBody")[0].children[0].namespaceURI;
        const timing = before.getElementsByTagNameNS(p, "timing")[0], serialize = n => new XMLSerializer().serializeToString(n);
        return xmls.every((xml, i) => {
          const dom = parse(xml), target = dom.getElementsByTagNameNS(p, "spTgt")[0], range = target.getElementsByTagNameNS(p, "pRg")[0];
          const owner = [...dom.getElementsByTagNameNS(p, "sp")].find(n => n.getElementsByTagNameNS(p, "cNvPr")[0].getAttribute("id") === target.getAttribute("spid"));
          const paragraphs = [...owner.getElementsByTagNameNS(p, "txBody")[0].children].filter(n => n.namespaceURI === a && n.localName === "p");
          return serialize(dom.getElementsByTagNameNS(p, "timing")[0]) === serialize(timing)
            && range.getAttribute("st") === "0" && range.getAttribute("end") === "2" && paragraphs.length === 3
            && paragraphs[0].textContent === "First tail" && paragraphs[1].textContent === ""
            && paragraphs[2].textContent === `Round ${i}<&>👩‍💻é`;
        });
      }, { sourceXml, xmls });
      assert.equal(verified, true); assert.deepEqual(requests, []); await page.close();
    }
    assert.deepEqual(await inspectPptxTiming(source), await inspectPptxTiming(input));
  }
});

test("browser paragraph opt-in retains default/character guards, mixed atomicity, cancellation and independent retry", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  const result = await page.evaluate(async ({ bytes, mixed, profile }) => {
    async function attempt(input, options) {
      const session = await PPTKit.importPptx(Uint8Array.from(input), options);
      try {
        const slide = session.document.slides[0], element = structuredClone(slide.elements[0]); element.content[0].runs[0].text = "changed"; slide.replaceElement(element.id, element);
        const prepared = await session.prepareSave({ mode: "edit" });
        return { status: prepared.status, codes: prepared.report?.diagnostics.map(d => d.code) ?? [], bytes: prepared.status === "ready" ? [...(await session.save(prepared.plan)).bytes] : null };
      } finally { session.dispose(); }
    }
    const options = { ...profile }, source = Uint8Array.from(bytes), pending = PPTKit.importPptx(source, options); options.textTargetSaveProfile = undefined; source.fill(0);
    const captured = await pending; const slide = captured.document.slides[0], next = structuredClone(slide.elements[0]); next.content[0].runs[0].text = "captured"; slide.replaceElement(next.id, next);
    const abort = new AbortController(), stopped = captured.prepareSave({ mode: "edit", signal: abort.signal }); abort.abort(); let cancelled;
    try { await stopped; } catch (e) { cancelled = e.code; }
    const retried = await captured.prepareSave({ mode: "edit" }); captured.dispose();
    const outcomes = await Promise.all([attempt(bytes, { projection: "content", editSaveProfile: "text-runs-v1" }), attempt(mixed, profile), attempt(bytes, profile)]);
    return { cancelled, retried: retried.status, outcomes };
  }, { bytes: [...paragraphBytes()], mixed: [...paragraphBytes({ target: paragraphTarget() + paragraphTarget('<p:charRg st="0" end="1"/>') })], profile: paragraphProfile });
  assert.equal(result.cancelled, "operation-aborted"); assert.equal(result.retried, "ready");
  for (const outcome of result.outcomes.slice(0, 2)) { assert.equal(outcome.status, "blocked"); assert.ok(outcome.codes.includes("save-text-range-reference")); assert.equal(outcome.bytes, null); }
  assert.equal(result.outcomes[2].status, "ready"); assert.ok(result.outcomes[2].bytes.length > 0);
});


test("paragraph target, font projection and starting number compose identically in Node, ESM and global", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  const esm = url(await readFile("packages/pptx-parser/dist/browser/index.js", "utf8"));
  const options = { ...paragraphProfile, firstSlideNumberSaveProfile: "attribute-v2", textFontProfile: "script-families-v3" };
  for (const strict of [false, true]) {
    const input = paragraphBytes({ strict }), s = await importPptx(input, options); let expected;
    try { editParagraph(s, "browser combined"); s.document.setFirstSlideNumber(17); const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready", JSON.stringify(p.report)); expected = await s.save(p.plan); } finally { s.dispose(); }
    const outputs = await page.evaluate(async ({ bytes, options, esm }) => { const out = []; for (const api of [await import(esm), PPTKit]) { const s = await api.importPptx(Uint8Array.from(bytes), options); try { const slide = s.document.slides[0], element = structuredClone(slide.elements[0]); element.content[0].runs[0].text = "browser combined"; slide.replaceElement(element.id, element); s.document.setFirstSlideNumber(17); const p = await s.prepareSave({ mode: "edit" }); if (p.status !== "ready") throw Error(JSON.stringify(p.report)); const saved = await s.save(p.plan); out.push({ bytes: [...saved.bytes], report: saved.report }); } finally { s.dispose(); } } return out; }, { bytes: [...input], options, esm });
    for (const out of outputs) { assert.deepEqual(out.bytes, [...expected.bytes]); assert.deepEqual(out.report, expected.report); }
  }
  assert.deepEqual(requests, []);
});
