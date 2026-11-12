import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { resolveLayout, resolveNormalizedLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
import { chartFixture, chartXml, series, workbook, encodeChart } from "../packages/pptx-parser/test/chart-fixture.mjs";

function make(chartType, values, extra = {}) {
  const deck = createPresentation({ id: "deck" });
  deck.addSlide({ id: "slide" }).addElement({ id: "chart", type: "chart", chartType, box: { x: 0, y: 0, width: 400, height: 300 }, categories: values.map((_, i) => `C${i}`), series: [{ name: "S", values }], ...extra });
  return deck;
}

test("explicit chart blanks survive cache, workbook, preview and three fresh import/export rounds", async () => {
  for (const kind of ["bar", "line", "pie"]) for (const values of [[3, null, 0, 2], [null, 1, null, null], [null, null, null, null]]) {
    const deck = make(kind, values); assert.deepEqual(validatePresentation(deck), []);
    assert.deepEqual(normalizePresentation(deck).slides[0].elements[0].series[0].values, values);
    const layout = resolveLayout(deck).slides[0].elements[0].chartLayout;
    if (kind === "bar") assert.deepEqual(layout.bars.map(bar => bar.categoryIndex), values.flatMap((value, i) => value === null ? [] : [i]));
    if (kind === "line") {
      assert.equal(layout.lines[0].points.length, values.filter(value => value !== null).length);
      assert.equal(layout.lines[0].segments.reduce((sum, segment) => sum + segment.length, 0), layout.lines[0].points.length);
      assert.equal(layout.lines[0].segments.length, values.reduce((sum, value, i) => sum + (value !== null && (i === 0 || values[i - 1] === null) ? 1 : 0), 0));
    }
    const svg = (await renderPresentationToSvg(deck)).slides[0].svg; assert.doesNotMatch(svg, /(?:NaN|Infinity)/);
    if (kind === "line") assert.equal((svg.match(/data-chart-role="line"/g) ?? []).length, layout.lines[0].segments.length);
    let bytes = (await generatePptx(deck)).bytes;
    for (let round = 0; round < 3; round++) {
      const parts = readGeneratedZip(bytes), xml = parts.get("ppt/charts/chart1.xml").toString(), workbook = readGeneratedZip(parts.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")), sheet = workbook.get("xl/worksheets/sheet1.xml").toString();
      assert.match(xml, /<c:dispBlanksAs val="gap"\/>/);
      const cache = xml.match(/<c:numCache>(.*?)<\/c:numCache>/s)[1]; assert.match(cache, /<c:ptCount val="4"\/>/);
      for (const [index, value] of values.entries()) {
        if (value === null) { assert.doesNotMatch(cache, new RegExp(`<c:pt idx="${index}">`)); assert.ok(sheet.includes(`<c r="B${index + 2}"/>`)); }
        else { assert.ok(cache.includes(`<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`)); assert.ok(sheet.includes(`<c r="B${index + 2}"><v>${value}</v></c>`)); }
      }
      const session = await importPptx(bytes, { projection: "content" });
      try {
        const chart = session.document.slides[0].elements[0]; assert.equal(chart.type, "chart"); assert.deepEqual(chart.series[0].values, values);
        const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual((await session.save(plan.plan)).bytes, bytes);
        bytes = (await generatePptx(session.document)).bytes;
      } finally { session.dispose(); }
    }
  }
});

test("sparse chart caches require a supported explicit mode and a category-bounded point count", async () => {
  const original = (await generatePptx(make("line", [3, null, 0, 2]))).bytes;
  const mutate = edit => { const parts = readGeneratedZip(original), name = "ppt/charts/chart1.xml"; parts.set(name, Buffer.from(edit(parts.get(name).toString()))); return zip([...parts].map(([name, body]) => ({ name, body }))).bytes; };
  const cases = [
    xml => xml.replace('<c:dispBlanksAs val="gap"/>', '<c:dispBlanksAs val="unknown"/>'),
    xml => xml.replace('<c:dispBlanksAs val="gap"/>', ''),
    xml => xml.replace('<c:formatCode>General</c:formatCode><c:ptCount val="4"/>', '<c:formatCode>General</c:formatCode><c:ptCount val="4294967295"/>'),
    xml => xml.replace('<c:pt idx="0"><c:v>3</c:v></c:pt>', '<c:pt idx="0"><c:v>NaN</c:v></c:pt>'),
  ];
  for (const edit of cases) {
    const session = await importPptx(mutate(edit), { projection: "content" });
    try { assert.ok(!session.document.slides[0].elements.some(element => element.type === "chart")); }
    finally { session.dispose(); }
  }
});

test("gap chart data edits remain blocked by existing-value preservation profile", async () => {
  for (const values of [[3, null, 0, 2], [null, null, null, null]]) {
    const source = (await generatePptx(make("line", values))).bytes, session = await importPptx(source, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try {
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.deepEqual(chart.series[0].values, values);
      chart.series[0].values[0] = 9; slide.replaceElement(chart.id, chart);
      const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "blocked");
    } finally { session.dispose(); }
  }
});

test("blank chart source and preview have Node and browser parity", async t => {
  const source = (await generatePptx(make("line", [3, null, 0, 2]))).bytes, node = await importPptx(source, { projection: "content" });
  let expected;
  try { const plan = await node.prepareSave({ mode: "no-op" }); expected = { values: node.document.slides[0].elements[0].series[0].values, bytes: Array.from((await generatePptx(node.document)).bytes), noOp: Array.from((await node.save(plan.plan)).bytes), svg: (await renderPresentationToSvg(node.document)).slides[0].svg }; }
  finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ["core", "layout", "pptx-exporter", "pptx-parser", "svg-renderer"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const actual = await page.evaluate(async bytes => {
    const session = await PPTKit.importPptx(Uint8Array.from(bytes), { projection: "content" });
    try { const plan = await session.prepareSave({ mode: "no-op" }); return { values: session.document.slides[0].elements[0].series[0].values, bytes: Array.from((await PPTKit.generatePptx(session.document)).bytes), noOp: Array.from((await session.save(plan.plan)).bytes), svg: (await PPTKit.renderPresentationToSvg(session.document)).slides[0].svg }; }
    finally { session.dispose(); }
  }, Array.from(source));
  assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
});

test("explicit gap literals import in both OOXML dialects without fabricating points", async () => {
  for (const strict of [false, true]) for (const family of ["bar", "line", "pie"]) {
    const data = series({ values: [0, 2] }).replace('<c:pt idx="1"><c:v>2</c:v></c:pt>', '');
    const source = encodeChart(chartFixture({ strict, xml: chartXml({ family, data, after: '<c:dispBlanksAs val="gap"/>' }) }));
    const session = await importPptx(source, { projection: "content" });
    try { const chart = session.document.slides[0].elements[0]; assert.equal(chart.chartType, family); assert.deepEqual(chart.series[0].values, [0, null]); const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source)); }
    finally { session.dispose(); }
  }
});

test("gap cache remains authoritative over conflicting workbook values and blocks edits", async () => {
  const parts = readGeneratedZip((await generatePptx(make("line", [3, null, 0, 2]))).bytes), bookPart = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx", sheetPart = "xl/worksheets/sheet1.xml", book = readGeneratedZip(parts.get(bookPart));
  book.set(sheetPart, Buffer.from(book.get(sheetPart).toString().replace('<c r="B3"/>', '<c r="B3"><v>8</v></c>')));
  parts.set(bookPart, zip([...book].map(([name, body]) => ({ name, body }))).bytes);
  const source = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
  const session = await importPptx(source, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  try {
    const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.deepEqual(chart.series[0].values, [3, null, 0, 2]);
    assert.ok(session.report.findings.some(f => f.feature === "chart.data-source" && f.reason.includes("conflicts with a complete nonblank workbook")));
    const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), Buffer.from(source));
    chart.series[0].values[0] = 9; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { session.dispose(); }
});

test("stacked bars skip blanks without inventing zero-valued bars", () => {
  for (const grouping of ["clustered", "stacked", "percentStacked"]) {
    const deck = createPresentation(); deck.addSlide().addElement({ type: "chart", chartType: "bar", grouping, box: { x: 0, y: 0, width: 300, height: 200 }, categories: ["A", "B", "C"], series: [{ name: "one", values: [null, -2, 0] }, { name: "two", values: [3, null, -3] }] });
    const layout = resolveLayout(deck).slides[0].elements[0].chartLayout;
    assert.deepEqual(layout.bars.map(b => [b.seriesIndex, b.categoryIndex, b.value]), [[1, 0, 3], [0, 1, -2], [0, 2, 0], [1, 2, -3]]);
    assert.ok(layout.bars.every(b => Object.values(b.box).every(Number.isFinite)));
    assert.ok(Number.isFinite(layout.valueScale.min) && Number.isFinite(layout.valueScale.max));
  }
});

test("blank line points do not expand the automatic scale toward zero", () => {
  const chart = resolveLayout(make("line", [null, 10, null, 11])).slides[0].elements[0];
  assert.ok(chart.chartLayout.valueScale.min > 0);
  assert.equal(chart.chartLayout.lines[0].points.length, 2);
  assert.deepEqual(chart.chartLayout.lines[0].segments.map(segment => segment.length), [1, 1]);
  const line = chart.chartLayout.lines[0], firstX = line.points[0].x; line.segments[0][0].x += 100;
  assert.equal(line.points[0].x, firstX);
});

test("unrelated text preservation edits leave gap chart cache and workbook untouched", async () => {
  const deck = make("line", [3, null, 0, 2]); deck.slides[0].addElement({ id: "text", type: "text", box: { x: 0, y: 310, width: 200, height: 30 }, content: "before" });
  const source = (await generatePptx(deck)).bytes, session = await importPptx(source, { projection: "content", editSaveProfile: "text-runs-v1", chartEditSaveProfile: "existing-values-v1" });
  try {
    const slide = session.document.slides[0], text = structuredClone(slide.elements.find(element => element.type === "text")); text.content[0].runs[0].text = "after"; slide.replaceElement(text.id, text);
    const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report));
    const output = (await session.save(plan.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(output);
    for (const name of ["ppt/charts/chart1.xml", "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]) assert.deepEqual(after.get(name), before.get(name));
    const reopened = await importPptx(output, { projection: "content" });
    try { assert.deepEqual(reopened.document.slides[0].elements.find(element => element.type === "chart").series[0].values, [3, null, 0, 2]); }
    finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

function rewriteWorkbook(source, editSheet, editChart = x => x) {
  const parts = readGeneratedZip(source), chartPart = "ppt/charts/chart1.xml", bookPart = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx", sheetPart = "xl/worksheets/sheet1.xml";
  const book = readGeneratedZip(parts.get(bookPart));
  book.set(sheetPart, Buffer.from(editSheet(book.get(sheetPart).toString())));
  parts.set(bookPart, zip([...book].map(([name, body]) => ({ name, body }))).bytes);
  parts.set(chartPart, Buffer.from(editChart(parts.get(chartPart).toString())));
  return zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
}
const removeNumericCache = xml => xml.replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, "");

test("missing caches can use category-bounded blank workbook values only under explicit gap mode", async () => {
  for (const family of ["bar", "line", "pie"]) for (const omitCell of [false, true]) {
    const original = (await generatePptx(make(family, [3, null, 0, 2]))).bytes;
    const source = rewriteWorkbook(original, xml => omitCell ? xml.replace('<c r="B3"/>', "") : xml, removeNumericCache);
    const session = await importPptx(source, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try {
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.equal(chart.chartType, family); assert.deepEqual(chart.series[0].values, [3, null, 0, 2]);
      assert.ok(session.report.findings.some(f => f.feature === "chart.data-source" && f.reason.includes("category-bounded workbook range with explicit blanks")));
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source));
      chart.series[0].values[0] = 9; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    } finally { session.dispose(); }
  }
  const spaced = rewriteWorkbook((await generatePptx(make("line", [3, null, 0, 2]))).bytes, xml => xml.replace('<c r="B3"/>', '<c r="B3" s="0"> \n </c>'), removeNumericCache);
  const spacedSession = await importPptx(spaced, { projection: "content" });
  try { assert.deepEqual(spacedSession.document.slides[0].elements[0].series[0].values, [3, null, 0, 2]); }
  finally { spacedSession.dispose(); }
  const original = (await generatePptx(make("line", [null, null, null, null]))).bytes;
  const allMissing = rewriteWorkbook(original, xml => xml.replace(/<c r="B[2-5]"\/>/g, ""), removeNumericCache);
  const session = await importPptx(allMissing, { projection: "content" });
  try { assert.deepEqual(session.document.slides[0].elements[0].series[0].values, [null, null, null, null]); }
  finally { session.dispose(); }
});

test("workbook error, text and unresolved formula cells never become blank points", async () => {
  const source = (await generatePptx(make("line", [3, null, 0, 2]))).bytes;
  const cells = [
    '<c r="B3" t="e"><v>#N/A</v></c>', '<c r="B3" t="b"><v>0</v></c>',
    '<c r="B3" t="inlineStr"><is><t></t></is></c>', '<c r="B3"><f>1/0</f></c>',
    '<c r="B3"><v/></c>', '<c r="B3"><unknown/></c>', '<c r="B3" vm="1"/>', '<c r="B3" s="invalid"/>', '<c r="B3">unexpected</c>',
  ];
  for (const cell of cells) {
    const altered = rewriteWorkbook(source, xml => xml.replace('<c r="B3"/>', cell), removeNumericCache);
    const session = await importPptx(altered, { projection: "content" });
    try { assert.ok(!session.document.slides[0].elements.some(element => element.type === "chart"), cell); }
    finally { session.dispose(); }
  }
  for (const mode of ["unknown"]) {
    const altered = rewriteWorkbook(source, xml => xml, xml => removeNumericCache(xml).replace('<c:dispBlanksAs val="gap"/>', `<c:dispBlanksAs val="${mode}"/>`));
    const session = await importPptx(altered, { projection: "content" });
    try { assert.ok(!session.document.slides[0].elements.some(element => element.type === "chart"), mode); }
    finally { session.dispose(); }
  }
  const oversized = rewriteWorkbook(source, xml => xml, xml => removeNumericCache(xml).replace('Sheet1!$B$2:$B$5', 'Sheet1!$B$2:$B$1048576'));
  const session = await importPptx(oversized, { projection: "content" });
  try { assert.ok(!session.document.slides[0].elements.some(element => element.type === "chart")); }
  finally { session.dispose(); }
});

test("chart cache and blank workbook agree or conflict without replacing the cache", async () => {
  const source = (await generatePptx(make("line", [3, null, 0, 2]))).bytes;
  const match = await importPptx(source, { projection: "content" });
  try { assert.ok(match.report.findings.some(f => f.feature === "chart.data-source" && f.reason.includes("agrees with stored workbook values including blanks"))); }
  finally { match.dispose(); }
  const shifted = rewriteWorkbook(source, xml => xml.replace('<c r="B3"/>', '<c r="B3"><v>8</v></c>').replace('<c r="B4"><v>0</v></c>', '<c r="B4"/>'));
  const shiftedSession = await importPptx(shifted, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  try {
    const slide = shiftedSession.document.slides[0], chart = structuredClone(slide.elements[0]);
    assert.deepEqual(chart.series[0].values, [3, null, 0, 2]);
    assert.ok(shiftedSession.report.findings.some(f => f.feature === "chart.data-source" && f.reason.includes("conflicts with stored workbook")));
    chart.series[0].values[0] = 9; slide.replaceElement(chart.id, chart); assert.equal((await shiftedSession.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { shiftedSession.dispose(); }
  const dense = (await generatePptx(make("line", [3, 8, 0, 2]))).bytes;
  const conflict = rewriteWorkbook(dense, xml => xml.replace('<c r="B3"><v>8</v></c>', '<c r="B3"/>'), xml => xml.replace('<c:dispBlanksAs val="zero"/>', '<c:dispBlanksAs val="gap"/>'));
  const session = await importPptx(conflict, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  try {
    const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.deepEqual(chart.series[0].values, [3, 8, 0, 2]);
    assert.ok(session.report.findings.some(f => f.feature === "chart.data-source" && f.reason.includes("conflicts with stored workbook")));
    chart.series[0].values[0] = 9; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  } finally { session.dispose(); }
});

test("workbook-only blanks match Node, browser ESM and global projections", async t => {
  const original = (await generatePptx(make("line", [3, null, 0, 2]))).bytes;
  const source = rewriteWorkbook(original, xml => xml.replace('<c r="B3"/>', ""), removeNumericCache);
  const select = session => ({ values: session.document.slides[0].elements[0].series[0].values, reasons: session.report.findings.filter(f => f.feature === "chart.data-source").map(f => f.reason) });
  const node = await importPptx(source, { projection: "content" }); let expected;
  try { const plan = await node.prepareSave({ mode: "no-op" }); expected = { ...select(node), noOp: Array.from((await node.save(plan.plan)).bytes) }; }
  finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  await page.evaluate(async sourceText => { globalThis.parserEsm = await import(`data:text/javascript;base64,${btoa(unescape(encodeURIComponent(sourceText)))}`); }, await readFile("packages/pptx-parser/dist/browser/index.js", "utf8"));
  const results = await page.evaluate(async bytes => {
    const results = [];
    for (const api of [PPTKit, parserEsm]) {
      const session = await api.importPptx(Uint8Array.from(bytes), { projection: "content" });
      try { const plan = await session.prepareSave({ mode: "no-op" }); results.push({ values: session.document.slides[0].elements[0].series[0].values, reasons: session.report.findings.filter(f => f.feature === "chart.data-source").map(f => f.reason), noOp: Array.from((await session.save(plan.plan)).bytes) }); }
      finally { session.dispose(); }
    }
    return results;
  }, Array.from(source));
  assert.deepEqual(results, [expected, expected]); assert.deepEqual(requests, []);
});

test("workbook-only blanks respect chart and spreadsheet OOXML dialects independently", async () => {
  for (const strictChart of [false, true]) for (const strictBook of [false, true]) {
    const parts = readGeneratedZip(workbook({ values: [0, 2.5] })), sheet = "xl/sheets/one.xml", original = parts.get(sheet).toString();
    const changed = original.replace('<c r="B3" ><v>2.5</v></c>', '<c r="B3"/>'); assert.notEqual(changed, original); parts.set(sheet, Buffer.from(changed));
    if (strictBook) for (const [name, body] of parts) if (name.endsWith(".xml") || name.endsWith(".rels")) parts.set(name, Buffer.from(body.toString().replaceAll("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main").replaceAll("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships")));
    const book = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
    const xml = chartXml({ family: "line", data: series({ refs: true, omit: true, values: [0, 2.5] }), external: true, after: '<c:dispBlanksAs val="gap"/>' });
    const source = encodeChart(chartFixture({ strict: strictChart, xml, book }));
    const session = await importPptx(source, { projection: "content" });
    try { assert.deepEqual(session.document.slides[0].elements[0].series[0].values, [0, null]); const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source)); }
    finally { session.dispose(); }
  }
});

test("workbook-only blanks survive three fresh projection and export rounds", async () => {
  for (const family of ["bar", "line", "pie"]) {
    let source = rewriteWorkbook((await generatePptx(make(family, [3, null, 0, 2]))).bytes, xml => xml.replace('<c r="B3"/>', ''), removeNumericCache);
    for (let round = 0; round < 3; round++) {
      const session = await importPptx(source, { projection: "content" });
      try {
        assert.deepEqual(session.document.slides[0].elements[0].series[0].values, [3, null, 0, 2]);
        const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source));
        source = (await generatePptx(session.document)).bytes;
      } finally { session.dispose(); }
    }
  }
});

test("blank workbook projection respects chart point budgets and leaves subsequent sessions usable", async () => {
  const original = (await generatePptx(make("line", [3, null, 0, 2]))).bytes, source = rewriteWorkbook(original, xml => xml, removeNumericCache);
  await assert.rejects(importPptx(source, { projection: "content", sessionLimits: { maxChartDataPoints: 3 } }), error => error.code === "import-limit-exceeded");
  const session = await importPptx(source, { projection: "content" });
  try { assert.deepEqual(session.document.slides[0].elements[0].series[0].values, [3, null, 0, 2]); }
  finally { session.dispose(); }
});

test("unrelated text edits preserve workbook-only blanks and the absent chart cache", async () => {
  const deck = make("line", [3, null, 0, 2]); deck.slides[0].addElement({ id: "text", type: "text", box: { x: 0, y: 310, width: 200, height: 30 }, content: "before" });
  const source = rewriteWorkbook((await generatePptx(deck)).bytes, xml => xml.replace('<c r="B3"/>', ''), removeNumericCache);
  const session = await importPptx(source, { projection: "content", editSaveProfile: "text-runs-v1", chartEditSaveProfile: "existing-values-v1" });
  try {
    const slide = session.document.slides[0], text = structuredClone(slide.elements.find(element => element.type === "text")); text.content[0].runs[0].text = "after"; slide.replaceElement(text.id, text);
    const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report));
    const output = (await session.save(plan.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(output);
    for (const name of ["ppt/charts/chart1.xml", "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]) assert.deepEqual(after.get(name), before.get(name));
    const reopened = await importPptx(output, { projection: "content" });
    try { assert.deepEqual(reopened.document.slides[0].elements.find(element => element.type === "chart").series[0].values, [3, null, 0, 2]); }
    finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test("an explicit gap cache remains authoritative when workbook blank cells are unsupported", async () => {
  const source = (await generatePptx(make("line", [3, null, 0, 2]))).bytes;
  for (const cell of ['<c r="B3" t="e"><v>#N/A</v></c>', '<c r="B3"><f>1/0</f></c>', '<c r="B3" vm="1"/>']) {
    const modified = rewriteWorkbook(source, xml => xml.replace('<c r="B3"/>', cell));
    const session = await importPptx(modified, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try {
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.deepEqual(chart.series[0].values, [3, null, 0, 2]);
      assert.ok(session.report.findings.some(f => f.feature === "chart.data-source" && f.reason.includes("missing, error or incompatible")));
      const noOp = await session.prepareSave({ mode: "no-op" }); assert.equal(noOp.status, "ready"); assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), Buffer.from(modified));
      chart.series[0].values[0] = 9; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    } finally { session.dispose(); }
  }
});

test("blank-aware workbook comparison preserves the sign of zero", async () => {
  const source = (await generatePptx(make("line", [3, null, -0, 2]))).bytes;
  const matching = await importPptx(source, { projection: "content" });
  try { assert.ok(Object.is(matching.document.slides[0].elements[0].series[0].values[2], -0)); assert.ok(matching.report.findings.some(f => f.reason.includes("agrees with stored workbook values including blanks"))); }
  finally { matching.dispose(); }
  const modified = rewriteWorkbook(source, xml => xml.replace('<c r="B4"><v>-0</v></c>', '<c r="B4"><v>0</v></c>'));
  const conflict = await importPptx(modified, { projection: "content" });
  try { assert.ok(Object.is(conflict.document.slides[0].elements[0].series[0].values[2], -0)); assert.ok(conflict.report.findings.some(f => f.reason.includes("conflicts with stored workbook"))); }
  finally { conflict.dispose(); }
});

test("blank display policies retain null data and produce distinct line geometry over three rounds", async () => {
  for (const mode of ["gap", "zero", "span"]) {
    const values = [5, null, 2, null, 7], deck = make("line", values, { blankDisplay: mode });
    assert.deepEqual(validatePresentation(deck), []); assert.equal(normalizePresentation(deck).slides[0].elements[0].blankDisplay, mode);
    const layout = resolveLayout(deck).slides[0].elements[0].chartLayout, line = layout.lines[0];
    assert.equal(line.points.length, mode === "zero" ? 5 : 3);
    const segments = line.segments ?? [line.points]; assert.deepEqual(segments.map(segment => segment.length), mode === "gap" ? [1, 1, 1] : [mode === "zero" ? 5 : 3]);
    if (mode === "zero") { assert.ok(layout.valueScale.min <= 0); assert.equal(line.points[1].y, line.points[3].y); }
    else assert.ok(layout.valueScale.min > 0);
    const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
    assert.equal((svg.match(/data-chart-role="line"/g) ?? []).length, mode === "gap" ? 3 : 1);
    assert.equal((svg.match(/data-chart-role="series-marker"/g) ?? []).length, mode === "zero" ? 5 : 3);
    let bytes = (await generatePptx(deck)).bytes;
    for (let round = 0; round < 3; round++) {
      const parts = readGeneratedZip(bytes), xml = parts.get("ppt/charts/chart1.xml").toString(), book = readGeneratedZip(parts.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")), sheet = book.get("xl/worksheets/sheet1.xml").toString();
      assert.ok(xml.includes(`<c:dispBlanksAs val="${mode}"/>`)); const cache = xml.match(/<c:numCache>(.*?)<\/c:numCache>/s)[1];
      for (const i of [1, 3]) { assert.ok(!cache.includes(`<c:pt idx="${i}">`)); assert.ok(sheet.includes(`<c r="B${i + 2}"/>`)); }
      const session = await importPptx(bytes, { projection: "content" });
      try { const chart = session.document.slides[0].elements[0]; assert.equal(chart.blankDisplay, mode); assert.deepEqual(chart.series[0].values, values); const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(bytes)); bytes = (await generatePptx(session.document)).bytes; }
      finally { session.dispose(); }
    }
  }
});

test("bar and pie accept gap and zero while span and unknown modes are rejected", async () => {
  for (const family of ["bar", "pie"]) {
    for (const mode of ["gap", "zero"]) {
      const deck = make(family, [3, null, 0, 2], { blankDisplay: mode }); assert.deepEqual(validatePresentation(deck), []);
      const layout = resolveLayout(deck).slides[0].elements[0].chartLayout;
      if (family === "bar") { assert.equal(layout.bars.length, mode === "zero" ? 4 : 3); if (mode === "zero") assert.deepEqual([layout.bars[1].categoryIndex, layout.bars[1].value, layout.bars[1].blank], [1, 0, true]); }
      const source = (await generatePptx(deck)).bytes, session = await importPptx(source, { projection: "content" });
      try { const chart = session.document.slides[0].elements[0]; assert.equal(chart.blankDisplay, mode); assert.deepEqual(chart.series[0].values, [3, null, 0, 2]); }
      finally { session.dispose(); }
    }
    for (const mode of ["span", "unknown"]) { const deck = make(family, [1, null], { blankDisplay: mode }); assert.ok(validatePresentation(deck).some(d => d.code === "invalid-chart-blank-display")); await assert.rejects(generatePptx(deck)); }
    const source = (await generatePptx(make(family, [1, null]))).bytes, parts = readGeneratedZip(source), part = "ppt/charts/chart1.xml";
    parts.set(part, Buffer.from(parts.get(part).toString().replace('<c:dispBlanksAs val="gap"/>', '<c:dispBlanksAs val="span"/>')));
    const session = await importPptx(zip([...parts].map(([name, body]) => ({ name, body }))).bytes, { projection: "content" });
    try { assert.ok(!session.document.slides[0].elements.some(e => e.type === "chart")); }
    finally { session.dispose(); }
  }
});

test("explicit blank policies also govern category-bounded workbook fallback", async () => {
  for (const family of ["bar", "line", "pie"]) for (const mode of family === "line" ? ["gap", "zero", "span"] : ["gap", "zero"]) {
    const source = rewriteWorkbook((await generatePptx(make(family, [3, null, 0, 2], { blankDisplay: mode }))).bytes, xml => xml.replace('<c r="B3"/>', ""), removeNumericCache);
    const session = await importPptx(source, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.equal(chart.blankDisplay, mode); assert.deepEqual(chart.series[0].values, [3, null, 0, 2]); chart.series[0].values[0] = 8; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); }
    finally { session.dispose(); }
  }
});

test("dense chart defaults remain zero while explicit policies survive fresh projection", async () => {
  for (const mode of [undefined, "gap", "zero", "span"]) {
    const deck = make("line", [1, 2], mode === undefined ? {} : { blankDisplay: mode });
    const bytes = (await generatePptx(deck)).bytes, xml = readGeneratedZip(bytes).get("ppt/charts/chart1.xml").toString();
    assert.ok(xml.includes(`<c:dispBlanksAs val="${mode ?? "zero"}"/>`));
    const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try {
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.equal(chart.blankDisplay, mode ?? "zero");
      chart.blankDisplay = chart.blankDisplay === "gap" ? "zero" : "gap"; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    } finally { session.dispose(); }
  }
});

test("explicit blank modes import correctly in strict and transitional chart dialects", async () => {
  for (const strict of [false, true]) for (const family of ["bar", "line", "pie"]) for (const mode of family === "line" ? ["gap", "zero", "span"] : ["gap", "zero"]) {
    const data = series({ values: [0, 2] }).replace('<c:pt idx="1"><c:v>2</c:v></c:pt>', '');
    const source = encodeChart(chartFixture({ strict, xml: chartXml({ family, data, after: `<c:dispBlanksAs val="${mode}"/>` }) }));
    const session = await importPptx(source, { projection: "content" });
    try { const chart = session.document.slides[0].elements[0]; assert.equal(chart.blankDisplay, mode); assert.deepEqual(chart.series[0].values, [0, null]); const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source)); }
    finally { session.dispose(); }
  }
});

test("zero and span line policies match between Node and browser", async t => {
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ["core", "layout", "pptx-exporter", "pptx-parser", "svg-renderer"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  for (const mode of ["zero", "span"]) {
    const source = (await generatePptx(make("line", [3, null, 0, 2], { blankDisplay: mode }))).bytes, node = await importPptx(source, { projection: "content" }); let expected;
    try { expected = { mode: node.document.slides[0].elements[0].blankDisplay, values: node.document.slides[0].elements[0].series[0].values, svg: (await renderPresentationToSvg(node.document)).slides[0].svg, bytes: Array.from((await generatePptx(node.document)).bytes) }; }
    finally { node.dispose(); }
    const actual = await page.evaluate(async bytes => { const s = await PPTKit.importPptx(Uint8Array.from(bytes), { projection: "content" }); try { return { mode: s.document.slides[0].elements[0].blankDisplay, values: s.document.slides[0].elements[0].series[0].values, svg: (await PPTKit.renderPresentationToSvg(s.document)).slides[0].svg, bytes: Array.from((await PPTKit.generatePptx(s.document)).bytes) }; } finally { s.dispose(); } }, Array.from(source));
    assert.deepEqual(actual, expected);
  }
  assert.deepEqual(requests, []);
});

test("legacy normalized charts without blankDisplay keep adaptive geometry", () => {
  for (const family of ["bar", "line", "pie"]) for (const values of [[1, 2, 3], [1, null, 3]]) {
    const normalized = normalizePresentation(make(family, values)), legacy = structuredClone(normalized); delete legacy.slides[0].elements[0].blankDisplay;
    assert.deepEqual(resolveNormalizedLayout(legacy).slides[0].elements[0].chartLayout, resolveNormalizedLayout(normalized).slides[0].elements[0].chartLayout);
  }
});

test("all-null and edge blanks do not fabricate span points or extrapolated segments", async () => {
  for (const mode of ["gap", "zero", "span"]) {
    const empty = make("line", [null, null, null, null], { blankDisplay: mode }), layout = resolveLayout(empty).slides[0].elements[0].chartLayout.lines[0];
    assert.equal(layout.points.length, mode === "zero" ? 4 : 0); assert.equal((layout.segments ?? [layout.points]).length, mode === "zero" ? 1 : 0);
    const svg = (await renderPresentationToSvg(empty)).slides[0].svg; assert.doesNotMatch(svg, /NaN|Infinity/);
    assert.equal((svg.match(/data-chart-role="series-marker"/g) ?? []).length, mode === "zero" ? 4 : 0);
  }
  const edge = resolveLayout(make("line", [null, 2, null, 4, null], { blankDisplay: "span" })).slides[0].elements[0].chartLayout.lines[0];
  assert.equal(edge.points.length, 2); assert.deepEqual(edge.segments.map(s => s.length), [2]);
});
