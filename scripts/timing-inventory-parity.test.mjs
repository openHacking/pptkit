import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { inspectPptxTiming } from "../packages/pptx-parser/dist/index.js";
import { timingBytes } from "../packages/pptx-parser/test/timing-fixture.mjs";
import { animationSpecimens } from "./research/native-timing-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
const url = source => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;

test("native timing inventories agree in Node/ESM/global and raw records match an independent Chromium DOM", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close());
  const esm = url(await readFile("packages/pptx-parser/dist/browser/index.js", "utf8")), global = await readFile("packages/pptx-parser/dist/browser/global.js", "utf8");
  for (const strict of [false, true]) for (const spec of animationSpecimens) {
    const bytes = timingBytes({ strict, name: spec.name }), expected = await inspectPptxTiming(bytes), sourceXml = readGeneratedZip(bytes).get("deck/z.xml").toString();
    for (const format of ["esm", "global"]) {
      const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
      if (format === "global") await page.addScriptTag({ content: global });
      const result = await page.evaluate(async ({ bytes, sourceXml, esm }) => {
        const api = esm ? await import(esm) : PPTKit, report = await api.inspectPptxTiming(Uint8Array.from(bytes));
        const dom = new DOMParser().parseFromString(sourceXml, "application/xml"), elements = [...dom.getElementsByTagName("*")], p = dom.documentElement.namespaceURI;
        const rows = report.slides[0].nodes, drawings = report.slides[0].drawings;
        const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
        const verified = rows.every(row => {
          const n = elements[row.source.elementIndex - 1]; if (!n || n.localName !== row.local || n.namespaceURI !== row.namespaceUri) return false;
          const attrs = [...n.attributes].map(a => ({ namespaceUri: a.namespaceURI ?? "", local: a.localName, value: a.value }));
          if (!same(attrs, row.attributes)) return false;
          const text = [...n.childNodes].filter(n => n.nodeType === Node.TEXT_NODE || n.nodeType === Node.CDATA_SECTION_NODE).map(n => n.nodeValue).join("");
          if (text !== (row.text ?? "")) return false;
          let parent = n.parentElement, time;
          while (parent) { if (parent.namespaceURI === p && parent.localName === "cTn") { time = elements.indexOf(parent) + 1; break; } parent = parent.parentElement; }
          if (time !== row.parentTimeNodeElementIndex) return false;
          if (row.drawingReference?.status === "located") {
            const target = elements[row.drawingReference.candidates[0].source.elementIndex - 1]; if (target.localName !== "cNvPr" || target.getAttribute("id") !== row.drawingReference.rawId) return false;
          }
          return true;
        }) && drawings.every(row => elements[row.source.elementIndex - 1]?.getAttribute("id") === row.rawId);
        return { report, verified, frozen: Object.isFrozen(report) && Object.isFrozen(report.slides[0].nodes) && rows.every(row => Object.isFrozen(row.source)) };
      }, { bytes: [...bytes], sourceXml, esm: format === "esm" ? esm : null });
      assert.deepEqual(result.report, expected); assert.equal(result.verified, true); assert.equal(result.frozen, true); assert.deepEqual(requests, []); await page.close();
    }
  }
});

test("browser timing captures options/bytes, enforces quotas and isolates cancellation/failure from concurrent readers", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage();
  await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  const result = await page.evaluate(async input => {
    const bytes = Uint8Array.from(input), options = { timingLimits: { maxRecords: 100000 } }, pending = PPTKit.inspectPptxTiming(bytes, options);
    bytes.fill(0); options.timingLimits.maxRecords = 1; const report = await pending;
    const limited = await Promise.allSettled([PPTKit.inspectPptxTiming(Uint8Array.from(input), { timingLimits: { maxWorkUnits: 1 } }), PPTKit.inspectPptxTiming(Uint8Array.from(input))]);
    const abort = new AbortController(); abort.abort(); let cancelled;
    try { await PPTKit.inspectPptxTiming(Uint8Array.from(input), { signal: abort.signal }); } catch (error) { cancelled = error.code; }
    return { same: JSON.stringify(report) === JSON.stringify(limited[1].value), failed: limited[0].reason.code, cancelled };
  }, [...timingBytes()]);
  assert.deepEqual(result, { same: true, failed: "package-limit-exceeded", cancelled: "operation-aborted" });
});
