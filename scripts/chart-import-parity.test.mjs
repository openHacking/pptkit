import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { createPresentation, normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { importPptx, inspectOpcPackage, inspectXmlPart } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { createChartImportFixture } from "../tests/fixtures/content-import.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const semantics = p => normalizePresentation(p).slides.map(s => s.elements.map(({ id, ...e }) => e));
const unescape = text => text.replaceAll("&lt;", "<").replaceAll("&gt;", ">").replaceAll("&quot;", '"').replaceAll("&apos;", "'").replaceAll("&amp;", "&").replace(/_x([0-9a-fA-F]{4})_/g, (_, code) => String.fromCharCode(parseInt(code, 16)));
const column = index => { let value = ""; for (let n = index; n > 0; n = Math.floor((n - 1) / 26)) value = String.fromCharCode(65 + (n - 1) % 26) + value; return value; };
// Independent inspection of writer-owned cache and worksheet scalar streams.
// Does not use the parser's private chart/workbook readers.
function verifyData(parts, expected) {
  const charts = [...parts].filter(([n]) => /^ppt\/charts\/chart\d+\.xml$/.test(n));
  const books = [...parts].filter(([n]) => n.endsWith(".xlsx"));
  assert.equal(charts.length, expected.length); assert.equal(books.length, expected.length);
  for (const [index, [, bytes]] of charts.entries()) {
    const xml = bytes.toString(), e = expected[index];
    const series = [...xml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)];
    assert.equal(series.length, e.series.length);
    for (const [i, [, ser]] of series.entries()) {
      assert.match(ser, new RegExp(`<c:idx val="${i}"/><c:order val="${i}"/>`));
      const col = column(i + 2), last = e.categories.length + 1;
      assert.ok(ser.includes(`<c:f>Sheet1!$${col}$1</c:f>`));
      assert.ok(ser.includes(`<c:f>Sheet1!$A$2:$A$${last}</c:f>`));
      assert.ok(ser.includes(`<c:f>Sheet1!$${col}$2:$${col}$${last}</c:f>`));
      const title = ser.match(/<c:tx>[\s\S]*?<c:strCache>([\s\S]*?)<\/c:strCache>/)[1];
      const names = [...title.matchAll(/<c:pt idx="(\d+)"><c:v>([^<]*)<\/c:v><\/c:pt>/g)];
      assert.deepEqual(names.map(m => [Number(m[1]), unescape(m[2])]), [[0, e.series[i].name]]);
      const cat = ser.match(/<c:cat>[\s\S]*?<c:strCache>([\s\S]*?)<\/c:strCache>/)[1];
      const categories = [...cat.matchAll(/<c:pt idx="(\d+)"><c:v>([^<]*)<\/c:v><\/c:pt>/g)];
      assert.deepEqual(categories.map(m => Number(m[1])), e.categories.map((_, j) => j));
      assert.deepEqual(categories.map(m => unescape(m[2])), e.categories);
      const values = [...ser.match(/<c:numCache>([\s\S]*?)<\/c:numCache>/)[1].matchAll(/<c:pt idx="(\d+)"><c:v>([^<]*)<\/c:v><\/c:pt>/g)];
      assert.deepEqual(values.map(m => Number(m[1])), e.categories.map((_, i) => i));
      assert.deepEqual(values.map(m => Number(m[2])), e.series[i].values);
    }
    const workbook = readGeneratedZip(books[index][1]), sheet = workbook.get("xl/worksheets/sheet1.xml").toString();
    const cells = new Map([...sheet.matchAll(/<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g)].map(m => [m[1], { type: m[2], value: unescape((m[3].match(/<(?:v|t)(?: [^>]*)?>([^<]*)<\/(?:v|t)>/) ?? [])[1] ?? "") }]));
    assert.equal(cells.size, e.series.length + e.categories.length * (1 + e.series.length));
    assert.equal(cells.has("A1"), false);
    for (const [i, category] of e.categories.entries()) assert.equal(cells.get(`A${i + 2}`).value, category);
    for (const [i, series] of e.series.entries()) {
      const col = column(i + 2); assert.equal(cells.get(`${col}1`).value, series.name);
      for (const [j, value] of series.values.entries()) assert.equal(Number(cells.get(`${col}${j + 2}`).value), value);
    }
  }
}

test("charts survive three fresh exports with independently verified cache/workbook data and frame transforms", async () => {
  let generated = await generatePptx(createChartImportFixture(createPresentation)), expected;
  const sizes = [];
  for (let round = 0; round < 3; round++) {
    assert.deepEqual(generated.warnings, []);
    const s = await importPptx(generated.bytes, { projection: "content" });
    assert.equal(s.document.slides[0].elements.length, 3); assert.deepEqual(validatePresentation(s.document), []);
    expected ??= semantics(s.document); assert.deepEqual(semantics(s.document), expected);
    const saved = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(saved.plan)).bytes, generated.bytes);
    const fresh = await generatePptx(s.document), parts = readGeneratedZip(fresh.bytes);
    verifyData(parts, s.document.slides[0].elements);
    const xml = parts.get("ppt/slides/slide1.xml").toString();
    assert.match(xml, /<p:xfrm rot="900000" flipH="1">/); assert.match(xml, /<p:xfrm rot="-600000" flipV="1">/);
    const graph = await inspectOpcPackage(fresh.bytes);
    assert.equal(graph.relationships.filter(r => r.type.endsWith("/chart")).length, 3);
    assert.equal(graph.relationships.filter(r => r.type.endsWith("/package")).length, 3);
    const preview = await renderPresentationToSvg(s.document); assert.match(preview.slides[0].svg, /Chart α😀/);
    sizes.push([...parts.values()].reduce((n, b) => n + b.length, 0)); generated = fresh; s.dispose();
  }
  assert.equal(new Set(sizes).size, 1);
});

test("native LibreOffice cache-only edit projects, preserves original bytes, and exports new workbooks", async () => {
  const source = new Uint8Array(await readFile(new URL("../tests/fixtures/corpus-charts-libreoffice-interactive.pptx", import.meta.url)));
  const types = ["bar", "line", "pie"];
  const points = [[1.5, 4, 3], [2, 4, 3], [2, 4, 3]];
  const charts = session => session.document.slides[0].elements.filter(element => element.type === "chart");
  const verifyProjection = session => {
    const elements = charts(session);
    assert.deepEqual(elements.map(chart => chart.chartType), types);
    assert.deepEqual(elements.map(chart => chart.categories), types.map(() => ["A", "B", "C"]));
    assert.deepEqual(elements.map(chart => chart.series.map(series => series.values)), points.map(values => [values]));
    return elements;
  };
  const graph = await inspectOpcPackage(source);
  assert.equal(graph.relationships.filter(rel => rel.type.endsWith("/chart")).length, 3);
  assert.equal(graph.relationships.filter(rel => rel.type.endsWith("/package")).length, 0);
  let fresh;
  for (const chartEditSaveProfile of ["existing-values-v1", "existing-values-v2"]) {
    const session = await importPptx(source, { projection: "content", chartEditSaveProfile });
    try {
      const [bar] = verifyProjection(session);
      assert.equal(session.report.chartValues.length, 9);
      assert.ok(session.report.chartValues.every(value => value.status === "unmapped" && /internal XLSX workbook/.test(value.reason)));
      const original = await session.prepareSave({ mode: "no-op" });
      assert.equal(original.status, "ready");
      assert.deepEqual((await session.save(original.plan)).bytes, source);
      fresh ??= await generatePptx(session.document);
      const edited = structuredClone(bar);
      edited.series[0].values[0] = 1.75;
      session.document.slides[0].replaceElement(bar.id, edited);
      const edit = await session.prepareSave({ mode: "edit" });
      assert.equal(edit.status, "blocked");
      assert.ok(edit.report.diagnostics.some(d => d.code === "save-edit-unsupported" && /internal XLSX workbook/.test(d.message)));
    } finally { session.dispose(); }
  }
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(fresh.bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try {
      verifyData(readGeneratedZip(fresh.bytes), verifyProjection(session));
      assert.equal(session.report.chartValues.length, 9);
      assert.ok(session.report.chartValues.every(value => value.status === "bound"));
      if (round < 2) fresh = await generatePptx(session.document);
      else {
        const [bar] = charts(session);
        const edited = structuredClone(bar);
        edited.series[0].values[0] = 1.75;
        session.document.slides[0].replaceElement(bar.id, edited);
        const plan = await session.prepareSave({ mode: "edit" });
        assert.equal(plan.status, "ready");
        verifyData(readGeneratedZip((await session.save(plan.plan)).bytes), charts(session));
      }
    } finally { session.dispose(); }
  }
});

test("fresh chart data inspection rejects stale category caches, series titles and range references", async () => {
  const deck = createChartImportFixture(createPresentation), generated = await generatePptx(deck), session = await importPptx(generated.bytes, { projection: "content" });
  try {
    const parts = readGeneratedZip(generated.bytes), expected = session.document.slides[0].elements, original = parts.get("ppt/charts/chart1.xml").toString();
    verifyData(parts, expected);
    const changes = [
      original.replace(/(<c:cat>[\s\S]*?<c:pt idx="0"><c:v>)[^<]*/, "$1STALE-CATEGORY"),
      original.replace(/(<c:tx>[\s\S]*?<c:pt idx="0"><c:v>)[^<]*/, "$1STALE-NAME"),
      original.replace(/Sheet1!\$A\$2:\$A\$\d+/, "Sheet1!$A$3:$A$4"),
    ];
    for (const xml of changes) { assert.notEqual(xml, original); const altered = new Map(parts); altered.set("ppt/charts/chart1.xml", Buffer.from(xml)); assert.throws(() => verifyData(altered, expected)); }
  } finally { session.dispose(); }
});

test("duplicate series names and categories remain indexed through multi-letter worksheet columns and three fresh exports", async () => {
  let generated = await generatePptx(createChartImportFixture(createPresentation));
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(generated.bytes, { projection: "content" });
    try {
      const slide = session.document.slides[0], old = slide.elements[0], next = structuredClone(old);
      next.categories = round === 1 ? ["Repeated", "Repeated", "A & <B>"] : ["A & <B>", "Repeated", "Repeated"];
      const count = round === 1 ? 25 : 27;
      next.series = Array.from({ length: count }, (_, i) => ({ name: "Duplicate", values: [i + round, i + round + 1, i + round + 2] }));
      if (round === 2) next.series.reverse();
      slide.removeElement(old.id); slide.insertElement(0, next);
      generated = await generatePptx(session.document); assert.deepEqual(generated.warnings, []);
      verifyData(readGeneratedZip(generated.bytes), slide.elements);
      const reopened = await importPptx(generated.bytes, { projection: "content" });
      try { const chart = reopened.document.slides[0].elements[0]; assert.deepEqual(chart.categories, next.categories); assert.deepEqual(chart.series.map(s => [s.name, s.values]), next.series.map(s => [s.name, s.values])); }
      finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});

test("fresh bar, line and pie exports preserve signed zero in both cache and workbook over three rounds", async () => {
  let generated = await generatePptx(createChartImportFixture(createPresentation));
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(generated.bytes, { projection: "content" });
    try {
      const slide = session.document.slides[0];
      for (const [index, element] of [...slide.elements].entries()) {
        const next = structuredClone(element); next.categories = ["Negative zero", "Positive zero", "Positive value"];
        for (const series of next.series) series.values = [-0, 0, 1];
        slide.removeElement(element.id); slide.insertElement(index, next);
      }
      generated = await generatePptx(session.document); assert.deepEqual(generated.warnings, []);
      const parts = readGeneratedZip(generated.bytes); verifyData(parts, slide.elements);
      for (let index = 1; index <= 3; index++) {
        const chart = parts.get(`ppt/charts/chart${index}.xml`).toString();
        for (const [, cache] of chart.matchAll(/<c:numCache>([\s\S]*?)<\/c:numCache>/g)) assert.match(cache, /<c:pt idx="0"><c:v>-0<\/c:v><\/c:pt><c:pt idx="1"><c:v>0<\/c:v>/);
        const workbook = readGeneratedZip(parts.get(`ppt/embeddings/Microsoft_Excel_Worksheet${index}.xlsx`));
        const sheet = workbook.get("xl/worksheets/sheet1.xml").toString(); assert.match(sheet, /<c r="B2"><v>-0<\/v><\/c>/); assert.match(sheet, /<c r="B3"><v>0<\/v><\/c>/);
      }
      const reopened = await importPptx(generated.bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
      try {
        for (const chart of reopened.document.slides[0].elements) for (const series of chart.series) { assert.ok(Object.is(series.values[0], -0)); assert.ok(Object.is(series.values[1], 0)); }
        assert.ok(reopened.report.chartValues.every(binding => binding.status === "bound"));
      } finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});

test("chart Xstring literals and carriage returns survive three fresh exports and numeric binding", async () => {
  let generated = await generatePptx(createChartImportFixture(createPresentation));
  const categories = ["_x0041_", "line\rbreak", "_x005F_x0041_ & 中文"], name = "literal _x000D_ and\rCR";
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(generated.bytes, { projection: "content" });
    try {
      const slide = session.document.slides[0], old = slide.elements[0], next = structuredClone(old);
      next.categories = categories; for (const series of next.series) { series.name = name; series.values = [1, 2, 3]; }
      slide.removeElement(old.id); slide.insertElement(0, next);
      generated = await generatePptx(session.document);
      const parts = readGeneratedZip(generated.bytes); verifyData(parts, slide.elements);
      const chart = parts.get("ppt/charts/chart1.xml").toString(), sheet = readGeneratedZip(parts.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")).get("xl/worksheets/sheet1.xml").toString();
      for (const xml of [chart, sheet]) { assert.ok(xml.includes("_x005F_x0041_")); assert.ok(xml.includes("line_x000D_break")); assert.ok(xml.includes("_x005F_x000D_")); }
      const reopened = await importPptx(generated.bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
      try { const result = reopened.document.slides[0].elements[0]; assert.deepEqual(result.categories, categories); assert.ok(result.series.every(s => s.name === name)); assert.ok(reopened.report.chartValues.filter(b => b.elementId === result.id).every(b => b.status === "bound")); }
      finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});

test("Xstring chart names and categories encode XML-forbidden controls and retain literal escape text", async () => {
  const generated = await generatePptx(createChartImportFixture(createPresentation)), session = await importPptx(generated.bytes, { projection: "content" });
  try {
    const slide = session.document.slides[0], old = slide.elements[0], chart = structuredClone(old);
    chart.categories = ["A\u0008B", "_x0008_", "NUL\u0000 and 😀", Array.from({ length: 32 }, (_, i) => String.fromCharCode(i)).join("") + "\ufffe\uffff"];
    for (const series of chart.series) { series.name = "\u000b_x000B_\u001f"; series.values = [1, 2, 3, 4]; }
    slide.replaceElement(old.id, chart);
    const saved = await generatePptx(session.document), parts = readGeneratedZip(saved.bytes); verifyData(parts, slide.elements);
    const xml = parts.get("ppt/charts/chart1.xml").toString(), sheet = readGeneratedZip(parts.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")).get("xl/worksheets/sheet1.xml").toString();
    for (const value of [xml, sheet]) { assert.ok(value.includes("A_x0008_B")); assert.ok(value.includes("_x005F_x0008_")); assert.ok(value.includes("NUL_x0000_ and 😀")); assert.ok(value.includes("_x000B__x005F_x000B__x001F_")); }
    const reopened = await importPptx(saved.bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try { const result = reopened.document.slides[0].elements[0]; assert.deepEqual(result.categories, chart.categories); assert.ok(result.series.every(s => s.name === chart.series[0].name)); assert.ok(reopened.report.chartValues.filter(b => b.elementId === result.id).every(b => b.status === "bound")); }
    finally { reopened.dispose(); }
    const preview = await renderPresentationToSvg(session.document), svg = preview.slides[0].svg;
    assert.ok(preview.warnings.some(w => w.code === "chart-text-xml-replacement" && w.elementId === chart.id));
    assert.ok(svg.includes("A\ufffdB")); assert.ok(svg.includes("&#13;")); assert.ok(svg.includes("😀")); assert.ok(svg.includes("_x0008_"));
    await inspectXmlPart(new TextEncoder().encode(svg)); assert.equal(slide.elements[0].categories[0], "A\u0008B");
    chart.categories[0] = "isolated \ud800"; slide.replaceElement(chart.id, chart);
    await assert.rejects(generatePptx(session.document), /XML 1.0/);
  } finally { session.dispose(); }
});

test("edited chart series deletion/reorder and category updates rebuild clean data over three rounds", async () => {
  let generated = await generatePptx(createChartImportFixture(createPresentation));
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(generated.bytes, { projection: "content" }), slide = s.document.slides[0];
    const plan = await s.prepareSave({ mode: "no-op" });
    const old = slide.elements[0], next = structuredClone(old);
    next.categories = round % 2 ? ["B", "A"] : ["A", "B"];
    next.series = [{ ...next.series.find(s => s.name === "Keep & <literal>"), values: [round + 1, round + 2] }, { name: `Added-${round}`, values: [round + 3, round + 4] }];
    if (round % 2) next.series.reverse();
    slide.removeElement(old.id); slide.insertElement(0, next);
    await assert.rejects(s.save(plan.plan), e => e.code === "save-snapshot-stale");
    assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
    assert.equal((await s.prepareSave({ mode: "rebuild" })).status, "blocked");
    generated = await generatePptx(s.document); assert.deepEqual(generated.warnings, []);
    const parts = readGeneratedZip(generated.bytes); verifyData(parts, slide.elements);
    for (const [name, bytes] of parts) {
      if (name.endsWith(".xml")) assert.doesNotMatch(bytes.toString(), /DELETE-ME/);
      if (name.endsWith(".xlsx")) for (const value of readGeneratedZip(bytes).values()) assert.doesNotMatch(value.toString(), /DELETE-ME/);
    }
    const again = await importPptx(generated.bytes, { projection: "content" });
    assert.deepEqual(again.document.slides[0].elements[0].series.map(s => [s.name, s.values]), next.series.map(s => [s.name, s.values]));
    again.dispose(); s.dispose();
  }
});

test('duplicate series and category names retain positional data through fresh structural edits', async () => {
  const doc = createPresentation(), slide = doc.addSlide();
  slide.addElement({ type: 'chart', chartType: 'bar', box: { x: 0, y: 0, width: 400, height: 300 }, categories: ['same', 'same', 'tail'], series: [{ name: 'duplicate', values: [1, 2, 3] }, { name: 'duplicate', values: [4, 5, 6] }, { name: 'remove-me', values: [7, 8, 9] }] });
  let bytes = (await generatePptx(doc)).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: 'content', chartEditSaveProfile: 'existing-values-and-series-names-v5' });
    try {
      const page = session.document.slides[0], old = page.elements[0], chart = structuredClone(old);
      const indices = round % 2 === 0 ? [1, 0] : [2, 1];
      chart.categories = [...indices.map(index => chart.categories[index]), 'same'];
      chart.series = [chart.series[1], chart.series[0]].map((series, index) => ({ ...series, values: [...indices.map(i => series.values[i]), 100 + round * 10 + index] }));
      chart.series.push({ name: 'duplicate', values: [10 + round, 20 + round, 30 + round] });
      page.replaceElement(old.id, chart);
      const preservation = await session.prepareSave({ mode: 'edit' }); assert.equal(preservation.status, 'blocked'); assert.equal(preservation.plan, undefined);
      assert.deepEqual(validatePresentation(session.document), []);
      bytes = (await generatePptx(session.document)).bytes;
      verifyData(readGeneratedZip(bytes), [chart]);
      const reopened = await importPptx(bytes, { projection: 'content' });
      try { const result = reopened.document.slides[0].elements[0]; assert.deepEqual(result.categories, chart.categories); assert.deepEqual(result.series.map(item => [item.name, item.values]), chart.series.map(item => [item.name, item.values])); } finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});
