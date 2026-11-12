import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { inspectPptxNotes } from "../packages/pptx-parser/dist/index.js";
import { notesInventoryFixture } from "../packages/pptx-parser/test/notes-inventory-fixture.mjs";
import { encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { notesInventorySourceAudit } from "../tests/support/notes-inventory-source-audit.mjs";
const moduleUrl = text => "data:text/javascript;base64," + Buffer.from(text).toString("base64");

test("notes inventory Node/ESM/global full reports match independently verified raw DOM sources with bounded failures", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", r => { requests.push(r.request().url()); return r.abort(); });
  const core = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8"));
  const parser = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(core)));
  for (const strict of [false, true]) for (const wrapped of [false, true]) {
    const fixture = notesInventoryFixture(strict);
    if (wrapped) fixture.parts["deck/notes/z.xml"] = fixture.parts["deck/notes/z.xml"].replace('<p:notes ', '<p:notes xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:future" mc:Ignorable="u" mc:ProcessContent="u:wrapper" ').replaceAll('<p:nvSpPr>', '<u:wrapper><p:nvSpPr>').replaceAll('</p:nvSpPr>', '</p:nvSpPr></u:wrapper>').replaceAll('<a:t>2042-01-02 &amp; stored</a:t>', '<u:wrapper><a:t>2042-01-02 &amp; stored</a:t></u:wrapper>');
    const bytes = encodeStructure(fixture), expected = await inspectPptxNotes(bytes);
    const truth = await page.evaluate(notesInventorySourceAudit, { parts: fixture.parts, report: expected });
    assert.ok(truth.checkedSourceReferences > 100); assert.equal(truth.fields, 3);
    const corrupt = structuredClone(expected); corrupt.parts.find(p => p.kind === "notes").fields[0].cachedTexts[0].value = "invented";
    await assert.rejects(page.evaluate(notesInventorySourceAudit, { parts: fixture.parts, report: corrupt }), /Wrong direct cached text/);
    const wrongParagraph = structuredClone(expected); wrongParagraph.parts.find(p => p.kind === "notes").fields[0].paragraphElementIndex = wrongParagraph.parts.find(p => p.kind === "notes").fields[1].paragraphElementIndex;
    await assert.rejects(page.evaluate(notesInventorySourceAudit, { parts: fixture.parts, report: wrongParagraph }), /Wrong enclosing paragraph source/);
    for (const format of ["esm", "global"]) {
      if (format === "global") for (const name of ["core", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
      const actual = await page.evaluate(async ({ bytes, parser }) => {
        const api = parser ? await import(parser) : PPTKit;
        const source = Uint8Array.from(bytes), options = { notesLimits: { maxRecords: 100000 } }, pending = api.inspectPptxNotes(source, options); source.fill(0); options.notesLimits.maxRecords = 1;
        const report = await pending;
        let limited, aborted;
        try { await api.inspectPptxNotes(Uint8Array.from(bytes), { notesLimits: { maxRecords: 1 } }); } catch (e) { limited = e.code; }
        const controller = new AbortController(); controller.abort();
        try { await api.inspectPptxNotes(Uint8Array.from(bytes), { signal: controller.signal }); } catch (e) { aborted = e.code; }
        return { report, limited, aborted, frozen: Object.isFrozen(report.parts[0].drawings) };
      }, { bytes: [...bytes], parser: format === "esm" ? parser : null });
      assert.deepEqual(actual, { report: expected, limited: "package-limit-exceeded", aborted: "operation-aborted", frozen: true });
    }
  }
  assert.deepEqual(requests, []);
});
