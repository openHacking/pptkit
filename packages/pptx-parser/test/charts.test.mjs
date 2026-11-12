import assert from "node:assert/strict";
import test from "node:test";
import { validatePresentation } from "@pptkit/core";
import { importPptx, DEFAULT_SESSION_LIMITS } from "../dist/index.js";
import { readFile } from "node:fs/promises";
import { chartBytes, chartXml, series, cache, frame, workbook, chartFixture, encodeChart } from "./chart-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { pack, relationships } from "./opc-fixture.mjs";
import { shape } from "./content-fixture.mjs";
import { replaceArchiveEntries } from "../dist/preservation/replace-archive-entries.js";
import { readXml } from "../dist/xml/read-xml.js";
import { chartReference } from "../dist/import/chart-data.js";
import { contentReader } from "../dist/import/content-reader.js";
import { chartWorkbooks } from "../dist/import/chart-workbook.js";
const open = (bytes = chartBytes(), options = {}) => importPptx(bytes, { projection: "content", ...options });
const first = s => s.document.slides[0].elements[0];
const reasons = s => s.report.findings.filter(f => f.feature.startsWith("chart.")).map(f => f.reason).join("\n");
const invalid = e => e.code === "import-invalid-package";

test('chart string sources retain XML-decoded lexical values separate from Xstring projection', async () => {
  for (const [c, a, p] of [
    ['http://schemas.openxmlformats.org/drawingml/2006/chart', 'http://schemas.openxmlformats.org/drawingml/2006/main', 'http://schemas.openxmlformats.org/presentationml/2006/main'],
    ['http://purl.oclc.org/ooxml/drawingml/chart', 'http://purl.oclc.org/ooxml/drawingml/main', 'http://purl.oclc.org/ooxml/presentationml/main'],
  ]) {
    const xml = `<c:cat xmlns:c="${c}"><c:strRef><c:f>Sheet1!$A$1:$A$2</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="1"><c:v>literal _x005F_x0041_ &amp; two</c:v></c:pt><c:pt idx="0"><c:v>_x0041_</c:v></c:pt></c:strCache></c:strRef></c:cat>`;
    const parsed = await readXml(new TextEncoder().encode(xml), {}, true);
    const result = await chartReference(parsed.root, false, c, contentReader('/chart.xml', p, a), async () => {}, () => {}, true);
    assert.deepEqual(result.values, ['A', 'literal _x0041_ & two']);
    assert.deepEqual(result.points.map(point => [point.index, point.lexical]), [[0, '_x0041_'], [1, 'literal _x005F_x0041_ & two']]);
    assert.deepEqual(result.points.map(point => point.node.children[0].value), ['_x0041_', 'literal _x005F_x0041_ & two']);
  }
});

test('workbook capture identifies only direct, simple inline string sources', async () => {
  const fixture = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const book = readGeneratedZip(fixture).get('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx');
  const load = bytes => chartWorkbooks({}, DEFAULT_SESSION_LIMITS, () => {}, async () => {}, true)('/book.xlsx', bytes);
  const result = await load(book), cell = result.sheets.get('Sheet1').get('A3');
  assert.equal(cell.value, 'literal _x0041_');
  assert.equal(cell.stringSource.lexical, 'literal _x005F_x0041_');
  assert.ok(cell.stringSource.elementIndex > 0); assert.equal(cell.source, undefined);
  const sheet = readGeneratedZip(book).get('xl/worksheets/sheet1.xml').toString();
  const original = '<is><t xml:space="preserve">literal _x005F_x0041_</t></is>';
  assert.ok(sheet.includes(original));
  for (const replacement of ['<is><r><t>literal _x005F_x0041_</t></r></is>', '<is><t future="1">literal _x005F_x0041_</t></is>', '<is><t>literal _x005F_x0041_</t><r><t>extra</t></r></is>']) {
    const changed = await replaceArchiveEntries(book, [{ name: 'xl/worksheets/sheet1.xml', bytes: new TextEncoder().encode(sheet.replace(original, replacement)) }], { maxOutputBytes: 10_000_000 });
    const unsafe = (await load(changed)).sheets.get('Sheet1').get('A3');
    assert.equal(unsafe.stringSource, undefined); assert.ok(unsafe.value.startsWith('literal _x0041_'));
  }
  const shared = await load(workbook({ shared: true }));
  assert.equal(shared.sheets.get('Sheet1').get('A2').stringSource, undefined);
  const without = await chartWorkbooks({}, DEFAULT_SESSION_LIMITS, () => {}, async () => {}, false)('/book.xlsx', book);
  assert.equal(without.sheets.get('Sheet1').get('A3').stringSource, undefined);
});

test("scatter sparse coordinates obey the aggregate budget and unsupported styles remain source-only", async () => {
  const data = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${cache(["XY"])}</c:tx><c:xVal>${cache([1], { numeric: true, count: 4294967295 })}</c:xVal><c:yVal>${cache([2], { numeric: true, count: 4294967295 })}</c:yVal></c:ser>`;
  const xml = chartXml({ family: "scatter", data, familyExtra: '<c:scatterStyle val="lineMarker"/>', after: '<c:dispBlanksAs val="gap"/>' }).replace('<c:catAx>', '<c:valAx>').replace('</c:catAx>', '</c:valAx>').replace('<c:axId val="11"/><c:scaling>', '<c:axId val="11"/><c:axPos val="b"/><c:scaling>').replace('<c:axId val="12"/><c:scaling>', '<c:axId val="12"/><c:axPos val="l"/><c:scaling>');
  await assert.rejects(open(chartBytes({ xml }), { sessionLimits: { maxChartDataPoints: 50 } }), e => e.code === "import-limit-exceeded");
  const unsupported = await open(chartBytes({ xml: xml.replace('val="lineMarker"', 'val="smoothMarker"') }));
  try { assert.equal(unsupported.document.slides[0].elements.length, 0); } finally { unsupported.dispose(); }
});

test("scatter signed axis IDs resolve exactly and enforce bounds, uniqueness and cross references", async () => {
  const data = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${cache(["XY"])}</c:tx><c:xVal>${cache([1, 2], { numeric: true })}</c:xVal><c:yVal>${cache([3, 4], { numeric: true })}</c:yVal></c:ser>`;
  const base = chartXml({ family: "scatter", data, familyExtra: '<c:scatterStyle val="lineMarker"/>', after: '<c:dispBlanksAs val="gap"/>' }).replace('<c:catAx>', '<c:valAx>').replace('</c:catAx>', '</c:valAx>').replace('<c:axId val="11"/><c:scaling>', '<c:axId val="11"/><c:axPos val="b"/><c:scaling>').replace('<c:axId val="12"/><c:scaling>', '<c:axId val="12"/><c:axPos val="l"/><c:scaling>');
  const xml = base.replaceAll('val="11"', 'val="-2147483648"').replaceAll('val="12"', 'val="4294967295"');
  for (const strict of [false, true]) {
    const bytes = chartBytes({ strict, xml }), session = await open(bytes);
    try { assert.deepEqual(first(session).series[0].xValues, [1, 2]); assert.deepEqual(first(session).series[0].yValues, [3, 4]); const plan = await session.prepareSave({ mode: 'no-op' }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes); } finally { session.dispose(); }
  }
  for (const bad of [xml.replaceAll('val="-2147483648"', 'val="-2147483649"'), xml.replaceAll('val="4294967295"', 'val="4294967296"'), xml.replaceAll('val="-2147483648"', 'val="1e3"')]) {
    const session = await open(chartBytes({ xml: bad })); try { assert.equal(session.document.slides[0].elements.length, 0); } finally { session.dispose(); }
  }
  for (const bad of [xml.replace('<c:crossAx val="4294967295"/>', '<c:crossAx val="-1"/>'), xml.replaceAll('val="4294967295"', 'val="-2147483648"'), xml.replace('<c:axId val="-2147483648"/>', '<c:axId val="2147483648"/>')]) await assert.rejects(open(chartBytes({ xml: bad })), invalid);
});

test("sparse scatter v2 writes existing coordinates in both OOXML dialects without filling blanks", async () => {
  const book = workbook({ name: "Revenue α😀", categories: ["one", "blank"], values: [4, 5] });
  const sheet = readGeneratedZip(book).get("xl/sheets/one.xml").toString();
  const changed = sheet.replace('<c r="A2" t="inlineStr"><is><t>one</t></is></c>', '<c r="A2"><v>1</v></c>').replace('<c r="A3" t="inlineStr"><is><t>blank</t></is></c>', '<c r="A3"/>');
  assert.notEqual(changed, sheet);
  const numericBook = await replaceArchiveEntries(book, [{ name: "xl/sheets/one.xml", bytes: new TextEncoder().encode(changed) }], { maxOutputBytes: 10_000_000 });
  const data = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx>${cache(["Revenue α😀"], { formula: "Sheet1!$B$1" })}</c:tx><c:xVal>${cache([1], { numeric: true, formula: "Sheet1!$A$2:$A$3", count: 2 })}</c:xVal><c:yVal>${cache([4, 5], { numeric: true, formula: "Sheet1!$B$2:$B$3" })}</c:yVal></c:ser>`;
  const xml = chartXml({ family: "scatter", data, familyExtra: '<c:scatterStyle val="lineMarker"/>', after: '<c:dispBlanksAs val="gap"/>', external: true }).replace('<c:catAx>', '<c:valAx>').replace('</c:catAx>', '</c:valAx>').replace('<c:axId val="11"/><c:scaling>', '<c:axId val="11"/><c:axPos val="b"/><c:scaling>').replace('<c:axId val="12"/><c:scaling>', '<c:axId val="12"/><c:axPos val="l"/><c:scaling>');
  for (const strict of [false, true]) {
    const bytes = chartBytes({ strict, xml, book: numericBook }), session = await open(bytes, { chartEditSaveProfile: "existing-values-v2" });
    try { assert.ok(session.report.chartValues.every(binding => binding.status === "bound")); const slide = session.document.slides[0], chart = structuredClone(first(session)); chart.series[0].xValues[0] = 7; chart.series[0].yValues[1] = 8; slide.replaceElement(chart.id, chart); const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report)); const reopened = await open((await session.save(plan.plan)).bytes, { chartEditSaveProfile: "existing-values-v2" }); try { assert.deepEqual(first(reopened).series[0].xValues, [7, null]); assert.deepEqual(first(reopened).series[0].yValues, [4, 8]); } finally { reopened.dispose(); } } finally { session.dispose(); }
  }
});

test("independent 2D bar/line/pie fixtures import in both dialects without invalid Core values", async () => {
  for (const strict of [false, true]) for (const family of ["bar", "line", "pie"]) {
    const s = await open(chartBytes({ strict, xml: chartXml({ family }) })), e = first(s);
    assert.equal(e.chartType, family); assert.equal(e.id, "source-slide-300-object-2");
    assert.equal(e.series[0].name, "Revenue α😀"); assert.deepEqual(e.categories, ["Q1", "Q2"]);
    assert.deepEqual(e.series[0].values, family === "pie" ? [0, 2.5] : [0, -2.5]);
    assert.deepEqual(e.transform, { rotation: 15, flipH: true, flipV: false });
    if (family === "bar") { assert.equal(e.orientation, "horizontal"); assert.deepEqual(e.axes.value.scale, { min: -5, max: 5, majorUnit: 2.5 }); }
    assert.deepEqual(validatePresentation(s.document), []);
    assert.ok(s.report.findings.some(f => f.feature === "drawing.chart.mapped-fields" && f.operation === "preview" && f.tier === "fallback"));
    s.dispose();
  }
});

test("point indices and series order determine data order independently of XML order", async () => {
  const data = series({ idx: 4, order: 9, name: "Second", values: [4, 9] }) + series({ idx: 90, order: 2, name: "First", values: [8, 3] }).replace(cache([8, 3], { numeric: true }), cache([3, 8], { numeric: true, indices: [1, 0] }));
  const s = await open(chartBytes({ xml: chartXml({ data }) })), e = first(s);
  assert.deepEqual(e.series.map(s => s.name), ["First", "Second"]); assert.deepEqual(e.series[0].values, [8, 3]); s.dispose();
});

test("complete caches win conflicts; compatible stored workbook results fill only missing caches", async () => {
  for (const shared of [false, true]) for (const omit of [false, true]) {
    const s = await open(chartBytes({ xml: chartXml({ data: series({ refs: true, omit }), external: true }), book: workbook({ values: [7, 12], shared }) }));
    assert.deepEqual(first(s).series[0].values, omit ? [7, 12] : [0, -2.5]);
    assert.match(reasons(s), omit ? /Missing chart cache was mapped/ : /cache conflicts/); s.dispose();
  }
  const s = await open(chartBytes({ xml: chartXml({ data: series({ refs: true }), external: true }), book: workbook() }));
  assert.match(reasons(s), /cache agrees/); assert.match(reasons(s), /relationships.*package/); s.dispose();
});

test("chart and inline or shared workbook text decode Xstring escapes once before comparison", async () => {
  for (const strict of [false, true]) for (const shared of [false, true]) for (const omit of [false, true]) {
    const names = "Revenue _x005F_x0041_", categories = ["_x0041_", "line_x000D_break"];
    const xml = chartXml({ data: series({ refs: true, omit, name: names, categories }), external: true });
    const book = workbook({ shared, name: names, categories });
    const s = await open(chartBytes({ strict, xml, book }), { chartEditSaveProfile: "existing-values-v1" });
    try {
      assert.equal(first(s).series[0].name, "Revenue _x0041_"); assert.deepEqual(first(s).categories, ["A", "line\rbreak"]);
      if (!omit) { assert.ok(s.report.chartValues.every(item => item.status === "bound")); assert.match(reasons(s), /cache agrees/); }
    } finally { s.dispose(); }
  }
});

test("workbook formulas use stored results and are never recalculated", async () => {
  const s = await open(chartBytes({ xml: chartXml({ data: series({ refs: true, omit: true }), external: true }), book: workbook({ formula: true, values: [6, 8] }) }));
  assert.deepEqual(first(s).series[0].values, [6, 8]); assert.match(reasons(s), /No formula is evaluated/); s.dispose();
});

test("sparse/empty/error data stays source-only and exact original bytes remain available", async () => {
  const original = series({ refs: true });
  const cases = [
    original.replace(cache([0, -2.5], { numeric: true, formula: "Sheet1!$B$2:$B$3" }), cache([5], { numeric: true, formula: "Sheet1!$B$2:$B$3", count: 2 })),
    series({ values: [] }), series({ values: ["", 2] }), series({ values: ["#N/A", 2] }), series({ values: ["NaN", 2] }), series({ values: ["INF", 2] }), series({ values: ["1e999", 2] }),
    series({ categories: ["", "Q2"] }), series({ name: "" }), series({ values: [2] }),
  ];
  for (const data of cases) {
    const bytes = chartBytes({ xml: chartXml({ data, external: true }), book: workbook() }), s = await open(bytes);
    assert.equal(s.document.slides[0].elements.length, 0);
    const p = await s.prepareSave({ mode: "no-op" }); assert.equal(p.status, "ready");
    assert.deepEqual(Buffer.from((await s.save(p.plan)).bytes), bytes); s.dispose();
  }
});

test("finite exponent/negative/zero values are preserved without lexical coercion", async () => {
  const s = await open(chartBytes({ xml: chartXml({ data: series({ values: ["-0", "  +1.25e-2  "] }) }) }));
  assert.ok(Object.is(first(s).series[0].values[0], -0)); assert.equal(first(s).series[0].values[1], .0125); s.dispose();
});

test("malformed point identities/counts, series identities/orders and axes reject", async () => {
  for (const data of [
    series().replace('idx="1"', 'idx="0"'),
    series().replace('idx="1"', 'idx="2"'),
    series().replace('val="2"', 'val="-1"'),
    series().replace('<c:ptCount val="2"/>', '<c:ptCount val="2"/><c:ptCount val="2"/>'),
    series() + series(),
    series({ idx: 1 }) + series({ idx: 2 }),
    series().replace('<c:v>0</c:v>', '<c:v><c:v/></c:v>'),
    series().replace('<c:cat>', '<c:cat><c:strLit/>'),
  ]) await assert.rejects(open(chartBytes({ xml: chartXml({ data }) })), invalid);
  for (const xml of [chartXml().replace('<c:crossAx val="12"/>', '<c:crossAx val="99"/>'), chartXml().replace('<c:axId val="12"/>', '<c:axId val="11"/>')]) await assert.rejects(open(chartBytes({ xml })), invalid);
});

test("external targets and unsupported formulas are not evaluated or loaded", async () => {
  const old = globalThis.fetch; globalThis.fetch = () => { throw Error("network must not run"); };
  try {
    const s = await open(chartBytes({ xml: chartXml({ data: series({ refs: true }), external: true }), external: true }));
    assert.deepEqual(first(s).series[0].values, [0, -2.5]); assert.match(reasons(s), /External workbook is not fetched/); s.dispose();
    for (const formula of ["SUM(Sheet1!B2:B3)", "[remote.xlsx]Sheet1!B2:B3", "NamedRange"]) {
      const xml = chartXml({ data: series({ refs: true, omit: true }), external: true }).replaceAll("Sheet1!$B$2:$B$3", formula);
      const s = await open(chartBytes({ xml, book: workbook() })); assert.equal(first(s), undefined); assert.match(reasons(s), /not a supported local A1/); s.dispose();
    }
  } finally { globalThis.fetch = old; }
});

test("pie colors use point indices and partial colors do not fabricate dense source colors", async () => {
  const point = (i, color) => `<c:dPt><c:idx val="${i}"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:dPt>`;
  const extras = point(1, "AABBCC") + point(0, "112233");
  const s = await open(chartBytes({ xml: chartXml({ family: "pie", data: series({ values: [1, 2], extras }) }) }));
  assert.deepEqual(first(s).series[0].pointColors, ["112233", "AABBCC"]); s.dispose();
  const partial = await open(chartBytes({ xml: chartXml({ family: "pie", data: series({ values: [1, 2], extras: point(1, "AABBCC") }) }) }));
  assert.equal(first(partial).series[0].pointColors, undefined); assert.match(reasons(partial), /Partial\/unsupported pie/); partial.dispose();
  await assert.rejects(open(chartBytes({ xml: chartXml({ family: "pie", data: series({ values: [1, 2], extras: extras + point(1, "FFFFFF") }) }) })), invalid);
});

test("direct title, legend, marker and line styles map with located unsupported chart properties", async () => {
  const before = '<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:rPr><a:t>Title &amp; α</a:t></a:r></a:p></c:rich></c:tx></c:title>';
  const after = '<c:legend><c:legendPos val="b"/></c:legend><c:extLst><c:ext uri="urn:unknown"><x:secret xmlns:x="urn:future"/></c:ext></c:extLst>';
  const extras = '<c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="diamond"/><c:size val="9"/></c:marker>';
  const s = await open(chartBytes({ xml: chartXml({ family: "line", before, after, data: series({ extras }) }) })), e = first(s);
  assert.equal(e.title, "Title & α"); assert.equal(e.titleStyle.fontSize, 18); assert.equal(e.titleStyle.bold, true);
  assert.equal(e.legend.position, "bottom"); assert.equal(e.series[0].line.width, 2); assert.equal(e.series[0].marker.shape, "diamond");
  assert.ok(s.report.findings.some(f => f.feature === "chart.unmapped-property" && f.source.part === "/deck/chart.xml" && f.source.path.includes("c:extLst") && f.elementId === e.id && f.source.sourceObjectId === "2")); s.dispose();
});

test("unsupported chart families and mismatched categories keep neighboring supported drawings", async () => {
  for (const xml of [chartXml({ family: "scatter" }), chartXml({ plotExtra: "<c:lineChart/>" }), chartXml({ data: series() + series({ idx: 8, order: 1, categories: ["A", "B"] }) }), chartXml({ family: "pie", data: series() }), chartXml().replace('<c:catAx>', '<c:dateAx>').replace('</c:catAx>', '</c:dateAx>')]) {
    const s = await open(chartBytes({ xml, content: frame() + shape({ id: "3" }) }));
    assert.equal(s.document.slides[0].elements.length, 1); assert.equal(first(s).type, "shape"); s.dispose();
  }
});

test("shared workbook budgets are charged once, aggregate points and malformed workbooks fail explicitly", async () => {
  const xml = chartXml({ data: series({ refs: true }), external: true }), book = workbook();
  const bytes = chartBytes({ xml, book, content: frame() + frame({ id: 3 }) });
  const expanded = [...readGeneratedZip(book).values()].reduce((n, b) => n + b.length, 0);
  const s = await open(bytes, { sessionLimits: { maxExpandedChartWorkbookBytes: expanded, maxTotalExpandedChartWorkbookBytes: expanded } });
  assert.equal(s.document.slides[0].elements.length, 2); s.dispose();
  await assert.rejects(open(bytes, { sessionLimits: { maxExpandedChartWorkbookBytes: 1 } }), e => e.code === "import-limit-exceeded" && e.source.part.includes("book.xlsx!/"));
  await assert.rejects(open(bytes, { sessionLimits: { maxChartDataPoints: 1 } }), e => e.code === "import-limit-exceeded");
  await assert.rejects(open(chartBytes({ xml, book: Buffer.from("not zip") })), invalid);
  const missing = chartFixture(); missing.parts["deck/_rels/z.xml.rels"] = missing.parts["deck/_rels/z.xml.rels"].replace('Id="chart"', 'Id="other"');
  await assert.rejects(open(encodeStructure(missing)), invalid);
  const controller = new AbortController(), pending = open(bytes, { signal: controller.signal }); controller.abort();
  await assert.rejects(pending, e => e.code === "operation-aborted");
});


function changeWorkbook(bytes, edit) {
  const parts = Object.fromEntries(readGeneratedZip(bytes)); edit(parts);
  const types = parts["[Content_Types].xml"].toString(); delete parts["[Content_Types].xml"];
  return pack(parts, types);
}

test("two distinct workbook parts share an aggregate expanded-byte budget", async () => {
  const book = workbook(), xml = chartXml({ data: series({ refs: true }), external: true });
  const f = chartFixture({ xml, book, content: frame() + frame({ id: 3, ref: "second" }) });
  f.parts["deck/second.xml"] = xml; f.types["deck/second.xml"] = "drawingml.chart";
  f.parts["deck/second.xlsx"] = book; f.types["deck/second.xlsx"] = "spreadsheetml.sheet";
  f.parts["deck/_rels/second.xml.rels"] = relationships([f.rel("book", "package", "second.xlsx")]);
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("chart", "chart", "chart.xml"), f.rel("second", "chart", "second.xml")]);
  const expanded = [...readGeneratedZip(book).values()].reduce((n, b) => n + b.length, 0), bytes = encodeChart(f);
  await assert.rejects(open(bytes, { sessionLimits: { maxTotalExpandedChartWorkbookBytes: expanded * 2 - 1 } }), e => e.code === "import-limit-exceeded");
  const s = await open(bytes, { sessionLimits: { maxTotalExpandedChartWorkbookBytes: expanded * 2 } }); assert.equal(s.document.slides[0].elements.length, 2); s.dispose();
});

test("malformed workbook cells, shared-string indices, content types and worksheet relationships fail at nested locations", async () => {
  const edits = [
    p => { p["xl/sheets/one.xml"] = p["xl/sheets/one.xml"].toString().replace('r="A2"', 'r="B2"'); },
    p => { p["xl/sheets/one.xml"] = p["xl/sheets/one.xml"].toString().replace('r="A2"', 'r="XFE2"'); },
    p => { p["xl/sheets/one.xml"] = p["xl/sheets/one.xml"].toString().replace('r="A2"', 'r="A1048577"'); },
    p => { p["xl/sheets/one.xml"] = p["xl/sheets/one.xml"].toString().replace('<v>0</v>', '<v>999999</v>'); },
    p => { p["xl/_rels/workbook.xml.rels"] = p["xl/_rels/workbook.xml.rels"].toString().replace('Id="sheet"', 'Id="unresolved"'); },
    p => { p["[Content_Types].xml"] = p["[Content_Types].xml"].toString().replace('spreadsheetml.worksheet+xml', 'spreadsheetml.sheet.main+xml'); },
  ];
  for (const edit of edits) {
    const book = changeWorkbook(workbook({ shared: true }), edit);
    await assert.rejects(open(chartBytes({ xml: chartXml({ data: series({ refs: true }), external: true }), book })), e => e.code === "import-invalid-package" && e.source.part.startsWith("/deck/book.xlsx!/") && !e.source.part.includes("book.xlsx!/deck/book.xlsx"));
  }
});

test("workbook errors and MCE remain unavailable data without replacing complete caches", async () => {
  const incompatible = workbook({ error: true });
  const mce = changeWorkbook(workbook(), p => { p["xl/sheets/one.xml"] = p["xl/sheets/one.xml"].toString().replace('<worksheet ', '<worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:q="urn:future" mc:Ignorable="q" '); });
  for (const book of [incompatible, mce]) for (const omit of [true, false]) {
    const s = await open(chartBytes({ xml: chartXml({ data: series({ refs: true, omit }), external: true }), book }));
    assert.equal(Boolean(first(s)), !omit); assert.match(reasons(s), /incompatible|Spreadsheet MCE/); s.dispose();
  }
});

test("namespace aliases, CDATA and selected chart AlternateContent use the effective retained tree", async () => {
  const f = chartFixture();
  f.parts["deck/chart.xml"] = f.parts["deck/chart.xml"].replace('<c:barChart>', '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:q="urn:future"><mc:Choice Requires="q"><c:scatterChart/></mc:Choice><mc:Fallback><c:barChart>').replace('</c:barChart>', '</c:barChart></mc:Fallback></mc:AlternateContent>').replace('<c:v>Q1</c:v>', '<c:v><![CDATA[Q1]]></c:v>').replaceAll('<c:', '<ch:').replaceAll('</c:', '</ch:').replace('xmlns:c=', 'xmlns:ch=');
  const s = await open(encodeChart(f)); assert.deepEqual(first(s).categories, ["Q1", "Q2"]); assert.ok(s.report.inspection.markupCompatibility.some(v => v.part === "/deck/chart.xml")); s.dispose();
});

test("quoted worksheet names, horizontal ranges and reversed/excess ranges have explicit data policy", async () => {
  const book = changeWorkbook(workbook(), p => { p["xl/workbook.xml"] = p["xl/workbook.xml"].toString().replace('name="Sheet1"', 'name="A &apos; B"'); });
  const quoted = chartXml({ data: series({ refs: true, omit: true }), external: true }).replaceAll("Sheet1!", "'A '' B'!");
  const s = await open(chartBytes({ xml: quoted, book })); assert.deepEqual(first(s).series[0].values, [0, -2.5]); s.dispose();
  const horizontal = changeWorkbook(workbook(), p => {
    let xml = p["xl/sheets/one.xml"].toString();
    const last = xml.match(/<c r="B3"[^>]*>[\s\S]*?<\/c>/)[0];
    xml = xml.replace(last, "").replace('</row><row r="3">', `${last.replace('r="B3"', 'r="C2"')}</row><row r="3">`);
    p["xl/sheets/one.xml"] = xml;
  });
  const horizontalXml = chartXml({ data: series({ refs: true, omit: true }), external: true }).replaceAll("Sheet1!$B$2:$B$3", "Sheet1!$B$2:$C$2");
  const h = await open(chartBytes({ xml: horizontalXml, book: horizontal })); assert.deepEqual(first(h).series[0].values, [0, -2.5]); h.dispose();
  for (const formula of ["Sheet1!$B$3:$B$2", "Sheet1!$A$2:$B$3", "Sheet1!$B$2:$B$1048576", "Sheet1!$XFE$2:$XFE$3"]) {
    const xml = chartXml({ data: series({ refs: true, omit: true }), external: true }).replaceAll("Sheet1!$B$2:$B$3", formula);
    const s = await open(chartBytes({ xml, book: workbook() })); assert.equal(first(s), undefined); s.dispose();
  }
});


test("missing-cache workbook projections cannot amplify shared ranges beyond the data-work budget", async () => {
  const bytes = chartBytes({ xml: chartXml({ data: series({ refs: true, omit: true }), external: true }), book: workbook(), content: frame() + frame({ id: 3 }) });
  // Five stored cells plus two readings of five values, with no chart cache.
  await assert.rejects(open(bytes, { sessionLimits: { maxChartDataPoints: 14 } }), e => e.code === "import-limit-exceeded");
  const s = await open(bytes, { sessionLimits: { maxChartDataPoints: 15 } }); assert.equal(s.document.slides[0].elements.length, 2); s.dispose();
});


test("chart extension islands remain opaque to nested MCE requirements", async () => {
  const after = '<c:extLst><c:ext uri="urn:future"><x:opaque xmlns:x="urn:future" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:MustUnderstand="x"><x:payload/></x:opaque></c:ext></c:extLst>';
  const s = await open(chartBytes({ xml: chartXml({ after }) }));
  assert.equal(first(s).chartType, "bar");
  assert.ok(s.report.findings.some(f => f.feature === "chart.unmapped-property" && f.source.path.includes("c:extLst"))); s.dispose();
});

test("huge declared caches remain source-only without count-driven allocation; invalid frame references reject", async () => {
  const s = await open(chartBytes({ xml: chartXml({ data: series().replaceAll('<c:ptCount val="2"/>', '<c:ptCount val="4294967295"/>') }) }));
  assert.equal(first(s), undefined); s.dispose();
  for (const edit of [
    f => { f.parts["deck/_rels/z.xml.rels"] = f.parts["deck/_rels/z.xml.rels"].replace('/relationships/chart"', '/relationships/image"'); },
    f => { f.parts["deck/_rels/z.xml.rels"] = f.parts["deck/_rels/z.xml.rels"].replace('Target="chart.xml"', 'Target="chart.xml#fragment"'); },
    f => { f.parts["deck/_rels/z.xml.rels"] = f.parts["deck/_rels/z.xml.rels"].replace('Target="chart.xml"', 'Target="https://example.invalid/chart.xml" TargetMode="External"'); },
    f => { f.parts["deck/chart.xml"] = f.parts["deck/chart.xml"].replace('<c:chart>', '<c:wrong>'); },
    f => { f.types["deck/chart.xml"] = "drawingml.chartStyle"; },
  ]) { const f = chartFixture(); edit(f); await assert.rejects(open(encodeChart(f)), invalid); }
});


test("unmapped chart properties classify every operation and source-only charts have no fabricated Core identity", async () => {
  const s = await open(chartBytes());
  const sourceOnlyProperty = s.report.findings.filter(f => f.feature === "chart.unmapped-property" && f.source.path.includes("c:scaling"));
  assert.ok(sourceOnlyProperty.some(f => f.operation === "no-op-save" && f.tier === "preserved"));
  assert.ok(sourceOnlyProperty.some(f => f.operation === "import" && f.tier === "unsupported"));
  s.dispose();
  const omitted = await open(chartBytes({ xml: chartXml({ data: series({ values: [] }) }) }));
  assert.ok(omitted.report.findings.filter(f => f.feature.startsWith("chart.")).every(f => f.elementId === undefined)); omitted.dispose();
  await assert.rejects(open(chartBytes(), { sessionLimits: { maxReportFindings: 1 } }), e => e.code === "import-limit-exceeded");
});

test("nested chart geometry rebases parent origins once and keeps source frame transforms", async () => {
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Parent"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000"><a:off x="1270000" y="1270000"/><a:ext cx="5080000" cy="3175000"/><a:chOff x="63500" y="88900"/><a:chExt cx="5080000" cy="3175000"/></a:xfrm></p:grpSpPr>${frame()}</p:grpSp>`;
  const s = await open(chartBytes({ content: group })), chart = first(s).children[0];
  assert.deepEqual(chart.box, { x: 5, y: 13, width: 400, height: 250 });
  assert.deepEqual(chart.transform, { rotation: 15, flipH: true, flipV: false });
  assert.equal(first(s).transform.rotation, 90); s.dispose();
});
