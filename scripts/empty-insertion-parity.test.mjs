import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { notesInsertionStyles } from "../tests/fixtures/notes-insertion-styles.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

const options = { projection: "content", editSaveProfile: "text-runs-v3" };
const moduleUrl = text => "data:text/javascript;base64," + Buffer.from(text).toString("base64");

test("empty insertion styles survive zero-run source projection and three fresh cycles in Node/ESM/global", async t => {
  const { document } = notesInsertionStyles(createPresentation), source = await generatePptx(document);
  assert.equal(source.warnings.filter(w => w.code === "script-font-map-materialized").length, 1);
  const parts = readGeneratedZip(source.bytes);
  // Independent source control models native deletion of empty a:r nodes;
  // it does not claim an application has performed that normalization.
  const collapsed = new Map(parts);
  for (const [part, body] of collapsed) if (part.startsWith("ppt/notesSlides/") && part.endsWith(".xml")) collapsed.set(part, Buffer.from(body.toString().replace(/<a:r>[\s\S]*?<\/a:r>/g, run => /<a:t><\/a:t>/.test(run) ? "" : run)));
  const bytes = zip([...collapsed].map(([name, body]) => ({ name, body }))).bytes;
  const expected = [];
  let next = bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(next, options);
    try {
      const notes = normalizePresentation(session.document).slides.map(s => s.notes);
      const runs = notes.map(n => n[0].runs[0]);
      assert.deepEqual(runs.map(r => [r.text, r.style.fontFamily, r.style.fontSize, r.style.bold, r.style.italic, r.style.language]), [["", "Georgia", 14, false, false, "fr-FR"], ["", "Arial", 12, true, true, "en-US"], ["", "Arial", 18, false, false, "en-US"], ["", "Georgia", 16, false, true, "en-US"]]);
      assert.ok(notes.every(n => n[0].runs.length === 1));
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(next));
      const generated = await generatePptx(session.document); next = generated.bytes;
      expected.push({ notes, bytes: [...next] });
    } finally { session.dispose(); }
  }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", r => { requests.push(r.request().url()); return r.abort(); });
  const coreUrl = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8"));
  const layoutUrl = moduleUrl((await readFile("packages/layout/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(coreUrl)));
  const parserUrl = moduleUrl((await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(coreUrl)));
  const exporterUrl = moduleUrl((await readFile("packages/pptx-exporter/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(coreUrl)).replaceAll('"@pptkit/layout"', JSON.stringify(layoutUrl)));
  for (const format of ["esm", "global"]) {
    if (format === "global") for (const name of ["core", "layout", "pptx-exporter", "pptx-parser"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
    const actual = await page.evaluate(async ({ bytes, options, urls }) => {
      const c = urls ? await import(urls.coreUrl) : PPTKit, p = urls ? await import(urls.parserUrl) : PPTKit, e = urls ? await import(urls.exporterUrl) : PPTKit, results = [];
      let next = Uint8Array.from(bytes);
      for (let round = 0; round < 3; round++) {
        const session = await p.importPptx(next, options);
        try {
          const notes = c.normalizePresentation(session.document).slides.map(s => s.notes);
          const plan = await session.prepareSave({ mode: "no-op" }); if (plan.status !== "ready") throw Error("Empty insertion no-op blocked");
          const saved = (await session.save(plan.plan)).bytes; if (saved.length !== next.length || saved.some((b, i) => b !== next[i])) throw Error("Empty insertion no-op bytes changed");
          next = (await e.generatePptx(session.document)).bytes;
          results.push({ notes, bytes: [...next] });
        } finally { session.dispose(); }
      }
      return results;
    }, { bytes: [...bytes], options, urls: format === "esm" ? { coreUrl, parserUrl, exporterUrl } : null });
    assert.deepEqual(actual, expected);
  }
  assert.deepEqual(requests, []);
});
