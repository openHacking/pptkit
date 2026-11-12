import assert from "node:assert/strict";
import test from "node:test";
import { readFile, mkdtemp, writeFile, rm } from "node:fs/promises";
import path from "node:path";
import { tmpdir } from "node:os";
import { chromium } from "@playwright/test";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { importPptxFile } from "../packages/pptx-parser/dist/node.js";
import { chartFixture, chartXml, series, workbook, encodeChart, frame } from "../packages/pptx-parser/test/chart-fixture.mjs";
import { shape, body } from "../packages/pptx-parser/test/content-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { replaceArchiveEntries } from "../packages/pptx-parser/dist/preservation/replace-archive-entries.js";
const options = { projection: "content", chartEditSaveProfile: "existing-values-v1" };

test('v6 order saves are blocked consistently in Node and browser entries', async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8') });
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, await readFile('packages/pptx-parser/dist/browser/index.js', 'utf8'));
  const config = { projection: 'content', chartEditSaveProfile: 'existing-series-order-v6' }, bytes = await readFile('tests/fixtures/chart-label-source.pptx');
  const node = await importPptx(bytes, config); let expected;
  try { const b = node.report.chartSeriesOrders[0]; expected = await node.prepareSave({ mode: 'edit', chartSeriesOrders: [{ slideId: b.slideId, elementId: b.elementId, identities: b.series.map(item => item.sourceId).reverse() }] }); } finally { node.dispose(); }
  const actual = await page.evaluate(async ({ bytes, config }) => {
    const results = [];
    for (const api of [globalThis.parser, globalThis.PPTKit]) { const session = await api.importPptx(Uint8Array.from(bytes), config); try { const b = session.report.chartSeriesOrders[0]; results.push(await session.prepareSave({ mode: 'edit', chartSeriesOrders: [{ slideId: b.slideId, elementId: b.elementId, identities: b.series.map(item => item.sourceId).reverse() }] })); } finally { session.dispose(); } }
    return results;
  }, { bytes: [...bytes], config });
  assert.equal(expected.status, 'blocked'); assert.equal(expected.report.diagnostics[0].code, 'save-chart-order-native-unsupported');
  for (const result of actual) assert.deepEqual(result, expected);
  assert.deepEqual(requests, []);
});

test('v4 category and numeric edits match Node, browser ESM and global across three reopens', async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8') });
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, await readFile('packages/pptx-parser/dist/browser/index.js', 'utf8'));
  const config = { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' };
  let bytes = await readFile('tests/fixtures/chart-label-source.pptx');
  for (const [round, value] of ['& < _x0041_', ' 保留空格 ', '中文 😀'].entries()) {
    const node = await importPptx(bytes, config); let expected, expectedLabels;
    try {
      const slide = node.document.slides[0], chart = structuredClone(slide.elements[0]); chart.categories[1] = value; chart.series[0].values[0] = 30 + round; slide.replaceElement(chart.id, chart);
      expectedLabels = node.report.chartLabels;
      const plan = await node.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); expected = await node.save(plan.plan);
    } finally { node.dispose(); }
    const actual = await page.evaluate(async ({ bytes, config, value, round }) => {
      const results = [];
      for (const api of [globalThis.parser, globalThis.PPTKit]) {
        const session = await api.importPptx(Uint8Array.from(bytes), config);
        try {
          const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.categories[1] = value; chart.series[0].values[0] = 30 + round; slide.replaceElement(chart.id, chart);
          const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw Error(JSON.stringify(plan.report));
          const saved = await session.save(plan.plan); results.push({ bytes: [...saved.bytes], report: saved.report, labels: session.report.chartLabels });
        } finally { session.dispose(); }
      }
      return results;
    }, { bytes: [...bytes], config, value, round });
    for (const result of actual) { assert.deepEqual(result.bytes, [...expected.bytes]); assert.deepEqual(result.report, expected.report); assert.deepEqual(result.labels, expectedLabels); }
    bytes = expected.bytes;
  }
  const unsupported = encodeChart(chartFixture({ xml: chartXml({ data: series({ refs: true }), external: true }), book: workbook({ shared: true }) }));
  const blocked = await page.evaluate(async ({ bytes, config }) => {
    const outcomes = [];
    for (const api of [globalThis.parser, globalThis.PPTKit]) {
      const session = await api.importPptx(Uint8Array.from(bytes), config);
      try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.categories[0] = 'not supported'; chart.series[0].values[0] = 77; slide.replaceElement(chart.id, chart); const plan = await session.prepareSave({ mode: 'edit' }); outcomes.push({ status: plan.status, hasPlan: !!plan.plan }); } finally { session.dispose(); }
    }
    return outcomes;
  }, { bytes: [...unsupported], config });
  assert.deepEqual(blocked, [{ status: 'blocked', hasPlan: false }, { status: 'blocked', hasPlan: false }]);
  assert.deepEqual(requests, []);
});

test('v4 inserts xml:space consistently across Node and both browser entries', async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8') });
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, await readFile('packages/pptx-parser/dist/browser/index.js', 'utf8'));
  const source = await readFile('tests/fixtures/chart-label-source.pptx'), bookName = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', sheetName = 'xl/worksheets/sheet1.xml';
  const book = readGeneratedZip(source).get(bookName), sheet = readGeneratedZip(book).get(sheetName).toString(), target = '<t xml:space="preserve">literal _x005F_x0041_</t>';
  assert.ok(sheet.includes(target)); const config = { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' };
  for (const attribute of ['', ' xml:space="default"']) {
    const updated = sheet.replace(target, `<t${attribute}>literal _x005F_x0041_</t>`);
    const changedBook = await replaceArchiveEntries(book, [{ name: sheetName, bytes: new TextEncoder().encode(updated) }], { maxOutputBytes: 10_000_000 });
    const bytes = await replaceArchiveEntries(source, [{ name: bookName, bytes: changedBook }], { maxOutputBytes: 10_000_000 });
    const session = await importPptx(bytes, config); let expected;
    try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.categories[1] = ' padded '; slide.replaceElement(chart.id, chart); const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); expected = await session.save(plan.plan); } finally { session.dispose(); }
    const results = await page.evaluate(async ({ bytes, config }) => {
      const output = [];
      for (const api of [globalThis.parser, globalThis.PPTKit]) {
        const session = await api.importPptx(Uint8Array.from(bytes), config);
        try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.categories[1] = ' padded '; slide.replaceElement(chart.id, chart); const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw Error(JSON.stringify(plan.report)); const saved = await session.save(plan.plan); output.push({ bytes: [...saved.bytes], report: saved.report }); } finally { session.dispose(); }
      }
      return output;
    }, { bytes: [...bytes], config });
    for (const result of results) { assert.deepEqual(result.bytes, [...expected.bytes]); assert.deepEqual(result.report, expected.report); }
  }
  assert.deepEqual(requests, []);
});

test('v5 mixed series names, categories and numbers agree across Node and browser entries', async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile('packages/pptx-parser/dist/browser/global.js', 'utf8') });
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, await readFile('packages/pptx-parser/dist/browser/index.js', 'utf8'));
  const config = { projection: 'content', chartEditSaveProfile: 'existing-values-and-series-names-v5' };
  let bytes = await readFile('tests/fixtures/chart-label-source.pptx');
  for (const [round, value] of [' & _x0041_ ', '中文 😀', 'last'].entries()) {
    const node = await importPptx(bytes, config); let expected, sources;
    try {
      sources = node.report.chartSeriesNames;
      const slide = node.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].name = value; chart.categories[1] = `category ${round}`; chart.series[0].values[0] = 80 + round; slide.replaceElement(chart.id, chart);
      const plan = await node.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report)); expected = await node.save(plan.plan);
    } finally { node.dispose(); }
    const actual = await page.evaluate(async ({ bytes, config, value, round }) => {
      const results = [];
      for (const api of [globalThis.parser, globalThis.PPTKit]) {
        const session = await api.importPptx(Uint8Array.from(bytes), config);
        try {
          const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].name = value; chart.categories[1] = `category ${round}`; chart.series[0].values[0] = 80 + round; slide.replaceElement(chart.id, chart);
          const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw Error(JSON.stringify(plan.report));
          const saved = await session.save(plan.plan); results.push({ bytes: [...saved.bytes], report: saved.report, sources: session.report.chartSeriesNames });
        } finally { session.dispose(); }
      }
      return results;
    }, { bytes: [...bytes], config, value, round });
    for (const result of actual) { assert.deepEqual(result.bytes, [...expected.bytes]); assert.deepEqual(result.report, expected.report); assert.deepEqual(result.sources, sources); }
    bytes = expected.bytes;
  }
  assert.deepEqual(requests, []);
});

test("Xstring chart/workbook text and numeric save agree across Node and browser", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, await readFile("packages/pptx-parser/dist/browser/index.js", "utf8"));
  for (const strict of [false, true]) for (const shared of [false, true]) {
    const name = "_x005F_x0041_", categories = ["_x0041_", "line_x000D_break"];
    const bytes = encodeChart(chartFixture({ strict, xml: chartXml({ data: series({ refs: true, name, categories }), external: true }), book: workbook({ shared, name, categories }) }));
    const run = async api => { const s = await api.importPptx(bytes, options); try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]); const view = { name: chart.series[0].name, categories: chart.categories, bindings: s.report.chartValues.map(b => b.status) }; chart.series[0].values[1] = 9; slide.replaceElement(chart.id, chart); const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready"); return { view, bytes: [...(await s.save(p.plan)).bytes] }; } finally { s.dispose(); } };
    const expected = await run({ importPptx });
    const results = await page.evaluate(async ({ bytes, options }) => { const out = []; for (const api of [globalThis.parser, globalThis.PPTKit]) { const s = await api.importPptx(Uint8Array.from(bytes), options); try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]); const view = { name: chart.series[0].name, categories: chart.categories, bindings: s.report.chartValues.map(b => b.status) }; chart.series[0].values[1] = 9; slide.replaceElement(chart.id, chart); const p = await s.prepareSave({ mode: "edit" }); if (p.status !== "ready") throw Error(JSON.stringify(p.report)); out.push({ view, bytes: [...(await s.save(p.plan)).bytes] }); } finally { s.dispose(); } } return out; }, { bytes: [...bytes], options });
    assert.equal(expected.view.name, "_x0041_"); assert.deepEqual(expected.view.categories, ["A", "line\rbreak"]); for (const result of results) assert.deepEqual(result, expected);
  }
  assert.deepEqual(requests, []);
});

test('underflow and true zero preserve chart source authority across Node and browser entries',async t=>{
  const browser=await chromium.launch({headless:true});t.after(()=>browser.close());const page=await browser.newPage(),requests=[];await page.route('**/*',route=>{requests.push(route.request().url());return route.abort();});
  await page.addScriptTag({content:await readFile('packages/pptx-parser/dist/browser/global.js','utf8')});await page.evaluate(async source=>{globalThis.parser=await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`);},await readFile('packages/pptx-parser/dist/browser/index.js','utf8'));
  for(const strict of [false,true])for(const [cacheValue,bookValue] of [['1e-400','1e-400'],['0','1e-400'],['0e-9999','0e-9999'],['5e-324','5e-324'],['9007199254740992','9007199254740993'],['0.1','0.10000000000000001'],['1.00','1e0']]){
    const bytes=encodeChart(chartFixture({strict,xml:chartXml({data:series({refs:true,values:[cacheValue,-2.5]}),external:true}),book:workbook({values:[bookValue,-2.5]})}));
    const node=await importPptx(bytes,options);let expected;try{expected={types:node.document.slides[0].elements.map(e=>e.type),sources:node.report.chartValues};}finally{node.dispose();}
    const actual=await page.evaluate(async ({bytes,options})=>{const results=[];for(const api of [globalThis.parser,globalThis.PPTKit]){const s=await api.importPptx(Uint8Array.from(bytes),options);try{results.push({types:s.document.slides[0].elements.map(e=>e.type),sources:s.report.chartValues});}finally{s.dispose();}}return results;},{bytes:[...bytes],options});
    for(const item of actual)assert.deepEqual(item,expected);
  }
  assert.deepEqual(requests,[]);
});

test("public Node/ESM/global synchronize chart cache and workbook with independent DOM and three-round byte parity", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage(); const requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(new URL("../packages/pptx-parser/dist/browser/global.js", import.meta.url), "utf8") });
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, await readFile(new URL("../packages/pptx-parser/dist/browser/index.js", import.meta.url), "utf8"));
  for (const strict of [false, true]) for (const family of ["bar", "line", "pie"]) {
    let bytes = encodeChart(chartFixture({ strict, xml: chartXml({ family, data: series({ refs: true, values: [1, 2] }), external: true }), book: workbook({ values: [1, 2] }) }));
    for (let round = 0; round < 3; round++) {
      const value = 12.5 + round, node = await importPptx(bytes, options); let saved;
      try {
        const slide = node.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[1] = value; slide.removeElement(chart.id); slide.insertElement(0, chart);
        const plan = await node.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); saved = await node.save(plan.plan);
        const actual = await page.evaluate(async ({ bytes, options, value }) => {
          const outputs = [];
          for (const api of [globalThis.parser, globalThis.PPTKit]) {
            const s = await api.importPptx(Uint8Array.from(bytes), options);
            try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[1] = value; slide.removeElement(chart.id); slide.insertElement(0, chart); const p = await s.prepareSave({ mode: "edit" }); if (p.status !== "ready") throw new Error(JSON.stringify(p.report)); const saved = await s.save(p.plan); outputs.push({ bytes: [...saved.bytes], report: saved.report, sources: s.report.chartValues }); }
            finally { s.dispose(); }
          }
          return outputs;
        }, { bytes: [...bytes], options, value });
        for (const result of actual) { assert.deepEqual(result.bytes, [...saved.bytes]); assert.deepEqual(result.report, saved.report); assert.deepEqual(result.sources, node.report.chartValues); }
      } finally { node.dispose(); }
      const before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes), bookBefore = readGeneratedZip(before.get("deck/book.xlsx")), bookAfter = readGeneratedZip(after.get("deck/book.xlsx"));
      for (const [name, data] of before) if (!["deck/chart.xml", "deck/book.xlsx"].includes(name)) assert.deepEqual(after.get(name), data);
      for (const [name, data] of bookBefore) if (name !== "xl/sheets/one.xml") assert.deepEqual(bookAfter.get(name), data);
      const independentlyRead = await page.evaluate(({ chartXml, sheetXml }) => {
        const chart = new DOMParser().parseFromString(chartXml, "application/xml"), sheet = new DOMParser().parseFromString(sheetXml, "application/xml"); if (chart.querySelector("parsererror") || sheet.querySelector("parsererror")) throw new Error("Malformed native XML");
        const ns = chart.documentElement.namespaceURI, values = [...chart.getElementsByTagNameNS(ns, "val")][0], pts = [...values.getElementsByTagNameNS(ns, "pt")].map(p => ({ index: Number(p.getAttribute("idx")), value: Number(p.getElementsByTagNameNS(ns, "v")[0].textContent) })).sort((a, b) => a.index - b.index);
        const sns = sheet.documentElement.namespaceURI, cell = [...sheet.getElementsByTagNameNS(sns, "c")].find(c => c.getAttribute("r") === "B3");
        return { seriesValue: pts[1].value, cellValue: Number(cell.getElementsByTagNameNS(sns, "v")[0].textContent) };
      }, { chartXml: after.get("deck/chart.xml").toString(), sheetXml: bookAfter.get("xl/sheets/one.xml").toString() });
      assert.deepEqual(independentlyRead, { seriesValue: value, cellValue: value }); bytes = saved.bytes;
    }
  }
  assert.deepEqual(requests, []);
});

test("Node file adapter captures numeric chart profile before filesystem awaits", async t => {
  const directory = await mkdtemp(path.join(tmpdir(), "pptkit-chart-values-")); t.after(() => rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, "source.pptx"); await writeFile(file, encodeChart(chartFixture({ xml: chartXml({ data: series({ refs: true }), external: true }), book: workbook() })));
  const config = { ...options }, pending = importPptxFile(file, config); config.chartEditSaveProfile = "invalid";
  const { session } = await pending;
  try { assert.ok(session.report.chartValues.every(b => b.status === "bound")); const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[0] = 6; slide.removeElement(chart.id); slide.insertElement(0, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "ready"); } finally { session.dispose(); }
});

test('Node file adapter captures v4 label profile before filesystem awaits', async t => {
  const directory = await mkdtemp(path.join(tmpdir(), 'pptkit-chart-label-')); t.after(() => rm(directory, { recursive: true, force: true }));
  const file = path.join(directory, 'source.pptx'); await writeFile(file, await readFile('tests/fixtures/chart-label-source.pptx'));
  const config = { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' }, pending = importPptxFile(file, config);
  config.chartEditSaveProfile = 'existing-values-v1';
  const { session } = await pending;
  try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.categories[1] = 'file adapter'; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'ready'); } finally { session.dispose(); }
});

test("Exporter-produced multi-series bar/line/pie source files bind and preserve numeric changes over three rounds", async () => {
  const { createPresentation } = await import("../packages/core/dist/index.js");
  const { generatePptx } = await import("../packages/pptx-exporter/dist/index.js");
  const { createChartImportFixture } = await import("../tests/fixtures/content-import.mjs");
  const { importPptx } = await import("../packages/pptx-parser/dist/index.js");
  let bytes = (await generatePptx(createChartImportFixture(createPresentation))).bytes;
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try {
      assert.equal(s.report.chartValues.length, 12); assert.ok(s.report.chartValues.every(v => v.status === "bound"));
      const slide = s.document.slides[0], before = readGeneratedZip(bytes);
      for (const [i, element] of [...slide.elements].entries()) {
        const chart = structuredClone(element); chart.series[0].values[0] = 10 + round + i;
        slide.removeElement(chart.id); slide.insertElement(i, chart);
      }
      const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready", JSON.stringify(p.report));
      bytes = (await s.save(p.plan)).bytes; const after = readGeneratedZip(bytes);
      const changes = [...before].filter(([name, payload]) => !payload.equals(after.get(name))).map(([name]) => name);
      assert.equal(changes.length, 6); assert.ok(changes.every(name => /charts\/chart|embeddings\//.test(name)));
      for (const [i, element] of slide.elements.entries()) {
        const b = s.report.chartValues.find(v => v.elementId === element.id && v.seriesIndex === 0 && v.pointIndex === 0);
        const chart = after.get(b.source.part.slice(1)).toString();
        const sheet = readGeneratedZip(after.get(b.workbook.part.slice(1))).get(b.workbook.worksheetPart.slice(1)).toString();
        assert.ok(chart.includes(`<c:v>${10 + round + i}</c:v>`));
        assert.ok(sheet.includes(`<v>${10 + round + i}</v>`));
      }
    } finally { s.dispose(); }
  }
});


test("combined starting number and chart values match Node, browser ESM and global across dialects", async t => {
  const browser = await chromium.launch({ headless: true }); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile(new URL("../packages/pptx-parser/dist/browser/global.js", import.meta.url), "utf8") });
  await page.evaluate(async source => { globalThis.parser = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(source)))}`); }, await readFile(new URL("../packages/pptx-parser/dist/browser/index.js", import.meta.url), "utf8"));
  const config = { ...options, firstSlideNumberSaveProfile: "attribute-v2" };
  for (const strict of [false, true]) {
    const bytes = encodeChart(chartFixture({ strict, xml: chartXml({ data: series({ refs: true, values: [1, 2] }), external: true }), book: workbook({ values: [1, 2] }) }));
    const session = await importPptx(bytes, config); let expected;
    try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[1] = 12; slide.replaceElement(chart.id, chart); session.document.setFirstSlideNumber(17); const p = await session.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready", JSON.stringify(p.report)); expected = await session.save(p.plan); } finally { session.dispose(); }
    const results = await page.evaluate(async ({ bytes, config }) => { const out = []; for (const api of [globalThis.parser, globalThis.PPTKit]) { const s = await api.importPptx(Uint8Array.from(bytes), config); try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[1] = 12; slide.replaceElement(chart.id, chart); s.document.setFirstSlideNumber(17); const p = await s.prepareSave({ mode: "edit" }); if (p.status !== "ready") throw Error(JSON.stringify(p.report)); const saved = await s.save(p.plan); out.push({ bytes: [...saved.bytes], report: saved.report }); } finally { s.dispose(); } } return out; }, { bytes: [...bytes], config });
    for (const result of results) { assert.deepEqual(result.bytes, [...expected.bytes]); assert.deepEqual(result.report, expected.report); }
  }
  assert.deepEqual(requests, []);
});


test("field projection and chart edits preserve mixed-save behavior across Node and browser", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  const field = '<a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="slidenum"><a:t>42</a:t></a:fld>';
  const config = { ...options, firstSlideNumberSaveProfile: "attribute-v2", fieldTextProfile: "slide-number-v1", editSaveProfile: "text-runs-v3" };
  for (const strict of [false, true]) {
    const input = encodeChart(chartFixture({ strict, content: frame() + shape({ id: "20", txBox: true, text: body(`<a:p><a:r><a:t>before</a:t></a:r>${field}</a:p>`) }), xml: chartXml({ data: series({ refs: true }), external: true }), book: workbook() }));
    const s = await importPptx(input, config); let expected;
    try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]), text = structuredClone(slide.elements[1]); chart.series[0].values[1] = 7; text.content[0].runs[0].text = "after"; slide.replaceElement(chart.id, chart); slide.replaceElement(text.id, text); s.document.setFirstSlideNumber(17); const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready", JSON.stringify(p.report)); expected = await s.save(p.plan); } finally { s.dispose(); }
    const result = await page.evaluate(async ({ bytes, config }) => { const s = await PPTKit.importPptx(Uint8Array.from(bytes), config); try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]), text = structuredClone(slide.elements[1]); chart.series[0].values[1] = 7; text.content[0].runs[0].text = "after"; slide.replaceElement(chart.id, chart); slide.replaceElement(text.id, text); s.document.setFirstSlideNumber(17); const p = await s.prepareSave({ mode: "edit" }); if (p.status !== "ready") throw Error(JSON.stringify(p.report)); const saved = await s.save(p.plan); text.content[0].runs[1].text = "999"; slide.replaceElement(text.id, text); const invalid = await s.prepareSave({ mode: "edit" }); return { bytes: [...saved.bytes], report: saved.report, invalid: invalid.status }; } finally { s.dispose(); } }, { bytes: [...input], config });
    assert.deepEqual(result.bytes, [...expected.bytes]); assert.deepEqual(result.report, expected.report); assert.equal(result.invalid, "blocked");
  }
  assert.deepEqual(requests, []);
});
