import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, validatePresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { importPptx, inspectOpcPackage } from "../packages/pptx-parser/dist/index.js";
import { createZip } from "../packages/pptx-exporter/dist/archive/create-zip.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

const specimen = {
  type: "chart", chartType: "combo", box: { x: 45, y: 30, width: 600, height: 345 },
  categories: ["Jan", "Feb", "Mar"], series: [
    { kind: "bar", name: "Volume", values: [2, 5, 8], color: "#4472C4" },
    { kind: "line", name: "Rate", values: [800, 200, 600], color: "#ED7D31" },
  ],
  axes: { value: { scale: { min: 0, max: 10, majorUnit: 2 } }, secondaryValue: { scale: { min: 0, max: 1000, majorUnit: 200 } } },
};
function author(element) {
  const deck = createPresentation();
  deck.addSlide({ elements: [element] });
  return deck;
}

test("combo fixed-axis warnings retain primary and secondary ownership after fresh export and import", async () => {
  const input = structuredClone(specimen); input.series[0].values[0] = 11; input.series[1].values[0] = 1001;
  const deck = author(input), before = validatePresentation(deck).filter(d => d.code === "chart-value-outside-fixed-axis");
  assert.deepEqual(before.map(d => d.path.match(/series\.\d+\.values\.\d+$/)[0]), ["series.0.values.0", "series.1.values.0"]);
  const bytes = (await generatePptx(deck)).bytes, session = await importPptx(bytes, { projection: "content" });
  try {
    const chart = session.document.slides[0].elements[0];
    assert.equal(chart.axes.value.scale.max, 10); assert.equal(chart.axes.secondaryValue.scale.max, 1000);
    const after = validatePresentation(session.document).filter(d => d.code === "chart-value-outside-fixed-axis");
    assert.deepEqual(after.map(d => d.path.match(/series\.\d+\.values\.\d+$/)[0]), ["series.0.values.0", "series.1.values.0"]);
  } finally { session.dispose(); }
});

test("combo chart rejects invalid series order and reports secondary axis errors at the authored path", () => {
  const reversed = structuredClone(specimen);
  reversed.series.reverse();
  assert.ok(validatePresentation(author(reversed)).some(d => d.code === "invalid-combo-series-order"));
  const broken = structuredClone(specimen);
  broken.axes.secondaryValue.scale.max = 0;
  assert.ok(validatePresentation(author(broken)).some(d => d.code === "invalid-chart-scale" && d.path.endsWith(".axes.secondaryValue.scale")));
});

test("combo chart uses independent scales with shared category centers and native workbook references", async t => {
  const deck = author(specimen);
  assert.deepEqual(validatePresentation(deck), []);
  const chart = resolveLayout(deck).slides[0].elements[0].chartLayout, plot = chart.plotBox;
  assert.equal(chart.valueScale.max, 10);
  assert.equal(chart.secondaryValueScale.max, 1000);
  const line = chart.lines[0];
  assert.equal(line.seriesIndex, 1);
  assert.deepEqual(line.points.map(p => Math.round(100 * (plot.y + plot.height - p.y) / plot.height)), [80, 20, 60]);
  assert.deepEqual(chart.bars.map(bar => Math.round(100 * bar.box.height / plot.height)), [20, 50, 80]);
  assert.ok(chart.bars.every((bar, i) => Math.abs(bar.box.x + bar.box.width / 2 - line.points[i].x) < 1e-8));
  const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
  assert.match(svg, /data-chart-role="secondary-label">1000<\/text>/);
  assert.match(svg, /data-chart-role="value-label">10<\/text>/);
  assert.equal([...svg.matchAll(/data-chart-role="bar"/g)].length, 3);
  assert.equal([...svg.matchAll(/data-chart-role="line"/g)].length, 1);
  const pptx = await generatePptx(deck), parts = readGeneratedZip(pptx.bytes);
  const xml = parts.get("ppt/charts/chart1.xml").toString();
  assert.match(xml, /<c:barChart>.*<c:axId val="1"\/><c:axId val="2"\/><\/c:barChart>/s);
  assert.match(xml, /<c:lineChart>.*<c:idx val="1"\/>.*Sheet1!\$C\$2:\$C\$4.*<c:axId val="3"\/><c:axId val="4"\/><\/c:lineChart>/s);
  assert.match(xml, /<c:catAx><c:axId val="3"\/>.*?<c:delete val="1"\/><c:axPos val="t"\/>.*?<c:crossAx val="4"\/>/s);
  assert.match(xml, /<c:valAx><c:axId val="4"\/>.*?<c:max val="1000"\/>.*?<c:axPos val="r"\/>.*?<c:crossAx val="3"\/>/s);
  const sheet = readGeneratedZip(parts.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")).get("xl/worksheets/sheet1.xml").toString();
  assert.match(sheet, /<c r="B2"><v>2<\/v><\/c><c r="C2"><v>800<\/v><\/c>/);

  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const name of ["core", "layout", "pptx-exporter", "svg-renderer", "pptx-parser"]) {
    await page.addScriptTag({ content: await readFile(new URL(`../packages/${name}/dist/browser/global.js`, import.meta.url), "utf8") });
  }
  const actual = await page.evaluate(async element => {
    const deck = PPTKit.createPresentation();
    deck.addSlide({ elements: [element] });
    const { bytes } = await PPTKit.generatePptx(deck);
    const session = await PPTKit.importPptx(bytes, { projection: "content" });
    try {
      return { svg: (await PPTKit.renderPresentationToSvg(deck)).slides[0].svg, pptx: [...bytes], imported: session.document.slides[0].elements[0] };
    } finally { session.dispose(); }
  }, specimen);
  const session = await importPptx(pptx.bytes, { projection: "content" });
  try {
    const imported = session.document.slides[0].elements[0];
    assert.equal(imported.chartType, "combo");
    assert.deepEqual(actual, { svg, pptx: [...pptx.bytes], imported });
  } finally { session.dispose(); }
});

test("combo auto scales are independent and a hidden secondary axis still positions its lines", async () => {
  const input = structuredClone(specimen);
  input.axes = { secondaryValue: { visible: false } };
  input.series[1].values = [800, null, 600];
  const deck = author(input), layout = resolveLayout(deck).slides[0].elements[0].chartLayout;
  assert.ok(layout.valueScale.max < 100);
  assert.ok(layout.secondaryValueScale.max >= 800);
  assert.equal(layout.lines[0].points.length, 2);
  assert.deepEqual(layout.lines[0].segments.map(segment => segment.length), [1, 1]);
  const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
  assert.doesNotMatch(svg, /data-chart-role="secondary-(?:axis|label)"/);
  assert.equal([...svg.matchAll(/data-chart-role="series-marker"/g)].length, 2);
});

test("native resaved combo without a workbook survives three fresh export/import rounds", async () => {
  let bytes = new Uint8Array(await readFile(new URL("../tests/fixtures/combination-libreoffice-resave.pptx", import.meta.url)));
  const evidence = JSON.parse(await readFile(new URL("../docs/testing/combination-public-probe-evidence.json", import.meta.url), "utf8"));
  assert.equal(createHash("sha256").update(bytes).digest("hex"), evidence.saved.sha256);
  const book = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
  assert.equal((await inspectOpcPackage(bytes)).relationships.some(rel => rel.type.endsWith("/package")), false);
  for (let round = 0; round <= 3; round++) {
    const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v3" });
    try {
      const slide = session.document.slides[0], chart = slide.elements[0];
      assert.equal(chart.chartType, "combo");
      assert.deepEqual(chart.categories, ["Jan", "Feb", "Mar"]);
      assert.deepEqual(chart.series.map(s => [s.kind, s.name, s.values]),
        [["bar", "Volume", [2, 5, 8]], ["line", "Rate", [800, 200, 600]]]);
      assert.deepEqual(chart.axes.value.scale, { min: 0, max: 10, majorUnit: 2 });
      assert.deepEqual(chart.axes.secondaryValue.scale, { min: 0, max: 1000, majorUnit: 200 });
      assert.equal(chart.series[1].marker.shape, "square");
      const noOp = await session.prepareSave({ mode: "no-op" });
      assert.equal(noOp.status, "ready");
      assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
      assert.equal(session.report.chartValues.length, 6);
      assert.ok(session.report.chartValues.every(binding => binding.status === (round === 0 ? "unmapped" : "bound")));
      if (round === 0) {
        const changed = structuredClone(chart);
        changed.series[1].values[0] = 250;
        slide.replaceElement(chart.id, changed);
        assert.notEqual((await session.prepareSave({ mode: "edit" })).status, "ready");
        slide.replaceElement(chart.id, chart);
      }
      if (round < 3) {
        bytes = (await generatePptx(session.document)).bytes;
        assert.equal(readGeneratedZip(bytes).has(book), true);
      }
    } finally { session.dispose(); }
  }
});

test("combo import preserves source-only charts when crossings, category centers or category labels disagree", async () => {
  const source = await generatePptx(author(specimen));
  const parts = readGeneratedZip(source.bytes), name = "ppt/charts/chart1.xml", chart = parts.get(name).toString();
  const patch = (tag, fn) => chart.replace(new RegExp(`<c:${tag}>[\\s\\S]*?</c:${tag}>`), fn);
  const cases = [
    [patch("lineChart", xml => xml.replace('<c:axId val="3"/><c:axId val="4"/>', '<c:axId val="1"/><c:axId val="2"/>')), /distinct reciprocal/],
    [chart.replace('<c:delete val="1"/><c:axPos val="t"/>', '<c:delete val="0"/><c:axPos val="t"/>'), /secondary category visibility/],
    [chart.replace('<c:delete val="1"/><c:axPos val="t"/>', '<c:delete val="0"/><c:axPos val="b"/>'), /secondary category visibility/],
    [chart.replace('<c:crosses val="autoZero"/>', '<c:crossesAt val="2"/>'), /Custom combo axis crossings/],
    [chart.replace('<c:crossBetween val="between"/>', '<c:crossBetween val="midCat"/>'), /shared category centers/],
    [patch("lineChart", xml => xml.replace("<c:v>Feb</c:v>", "<c:v>Apr</c:v>")), /category arrays disagree/],
    [patch("barChart", xml => xml.replace('grouping val="clustered"', 'grouping val="stacked"')), /requires default clustered columns/],
  ];
  for (const [altered, reason] of cases) {
    assert.notEqual(altered, chart, String(reason));
    const bytes = createZip([...parts].map(([part, data]) => ({ name: part, data: part === name ? new TextEncoder().encode(altered) : data })));
    const session = await importPptx(bytes, { projection: "content" });
    try {
      assert.deepEqual(session.document.slides[0].elements, []);
      assert.ok(session.report.findings.some(f => reason.test(f.reason)), String(reason));
      const noOp = await session.prepareSave({ mode: "no-op" });
      assert.equal(noOp.status, "ready");
      assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
    } finally { session.dispose(); }
  }
});

test("combo axes follow family references even when XML axes and families are reordered", async () => {
  const parts = readGeneratedZip((await generatePptx(author(specimen))).bytes);
  const name = "ppt/charts/chart1.xml", xml = parts.get(name).toString();
  const chart = xml.replace(/(<c:barChart>[\s\S]*?<\/c:barChart>)(<c:lineChart>[\s\S]*?<\/c:lineChart>)/, "$2$1")
    .replace(/(<c:axId val="3"\/>)(<c:axId val="4"\/>)(?=<\/c:lineChart>)/, "$2$1");
  assert.notEqual(chart, xml);
  assert.match(chart, /<c:lineChart>.*?<c:axId val="4"\/><c:axId val="3"\/><\/c:lineChart><c:barChart>/s);
  const bytes = createZip([...parts].map(([part, data]) => ({ name: part, data: part === name ? new TextEncoder().encode(chart) : data })));
  const session = await importPptx(bytes, { projection: "content" });
  try {
    const projected = session.document.slides[0].elements[0];
    assert.equal(projected.chartType, "combo");
    assert.deepEqual(projected.series.map(s => [s.kind, s.values]), [["bar", [2, 5, 8]], ["line", [800, 200, 600]]]);
    assert.deepEqual([projected.axes.value.scale.max, projected.axes.secondaryValue.scale.max], [10, 1000]);
  } finally { session.dispose(); }
});

test("combo edits change each family's numeric cache and worksheet atom over three source-save rounds", async () => {
  const options = { projection: "content", chartEditSaveProfile: "existing-values-v3" };
  let bytes = (await generatePptx(author(specimen))).bytes;
  const xmlPath = "ppt/charts/chart1.xml", bookPath = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx";
  const sheetPath = "xl/worksheets/sheet1.xml";
  const replace = (xml, old, next) => {
    assert.equal(xml.split(old).length, 2, old);
    return xml.replace(old, next);
  };
  const prior = [5, 200];
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, options);
    try {
      const bindings = session.report.chartValues;
      assert.equal(bindings.length, 6);
      assert.ok(bindings.every(b => b.status === "bound" && !b.coordinate));
      assert.deepEqual(bindings.filter(b => b.pointIndex === 1).map(b => b.workbook.cellReference), ["B3", "C3"]);
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]);
      assert.deepEqual(chart.series.map(series => series.values[1]), prior);
      const next = [4.25 + round, 450 + round * 40];
      const doc = session.document, changes = next.map((value, seriesIndex) => ({ type: 'chart-point', slideId: slide.id, elementId: chart.id, seriesIndex, pointIndex: 1, coordinate: 'values', expected: prior[seriesIndex], value }));
      const undo = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes }).inverse;
      const redo = doc.applyEditOperation(undo).inverse; doc.applyEditOperation(redo);
      const plan = await session.prepareSave({ mode: "edit" });
      assert.equal(plan.status, "ready", JSON.stringify(plan.report?.diagnostics));
      const saved = await session.save(plan.plan), before = readGeneratedZip(bytes), after = readGeneratedZip(saved.bytes);
      assert.deepEqual([...after.keys()], [...before.keys()]);
      for (const [name, data] of before) if (name !== xmlPath && name !== bookPath) assert.deepEqual(after.get(name), data, name);
      let expectedChart = before.get(xmlPath).toString();
      const oldBook = readGeneratedZip(before.get(bookPath)), newBook = readGeneratedZip(after.get(bookPath));
      assert.deepEqual([...newBook.keys()], [...oldBook.keys()]);
      for (const [name, data] of oldBook) if (name !== sheetPath) assert.deepEqual(newBook.get(name), data, name);
      let expectedSheet = oldBook.get(sheetPath).toString();
      for (const [i, cell] of ["B3", "C3"].entries()) {
        expectedChart = replace(expectedChart, `<c:pt idx="1"><c:v>${prior[i]}</c:v></c:pt>`, `<c:pt idx="1"><c:v>${next[i]}</c:v></c:pt>`);
        expectedSheet = replace(expectedSheet, `<c r="${cell}"><v>${prior[i]}</v></c>`, `<c r="${cell}"><v>${next[i]}</v></c>`);
      }
      assert.equal(after.get(xmlPath).toString(), expectedChart);
      assert.equal(newBook.get(sheetPath).toString(), expectedSheet);
      prior.splice(0, 2, ...next);
      bytes = saved.bytes;
    } finally { session.dispose(); }
  }
  for (const chartEditSaveProfile of ["existing-values-v1", "existing-values-v2", "existing-values-v3"]) {
    const reopened = await importPptx(bytes, { ...options, chartEditSaveProfile });
    try {
      const slide = reopened.document.slides[0], chart = structuredClone(slide.elements[0]);
      assert.deepEqual(chart.series.map(series => series.values[1]), prior);
      assert.equal(reopened.report.chartValues.length, chartEditSaveProfile === "existing-values-v3" ? 6 : 0);
      chart.series[0].values[1] = 9;
      if (chartEditSaveProfile === "existing-values-v3") chart.axes.secondaryValue.scale.max = 1200;
      slide.replaceElement(chart.id, chart);
      assert.notEqual((await reopened.prepareSave({ mode: "edit" })).status, "ready");
    } finally { reopened.dispose(); }
  }
});

test("combo preservation refuses aliased value cells and stale caches", async () => {
  const input = structuredClone(specimen);
  input.series[1].values = [...input.series[0].values];
  const parts = readGeneratedZip((await generatePptx(author(input))).bytes);
  const name = "ppt/charts/chart1.xml", xml = parts.get(name).toString();
  const line = fn => xml.replace(/<c:lineChart>[^]*?<\/c:lineChart>/, fn);
  const cases = [
    [line(group => group.replace("Sheet1!$C$2:$C$4", "Sheet1!$B$2:$B$4")), binding => binding.status === "unmapped" && /shared by multiple chart references/.test(binding.reason)],
    [line(group => group.replace("<c:v>5</c:v>", "<c:v>6</c:v>")), binding => binding.seriesIndex === 0 ? binding.status === "bound" : binding.status === "unmapped" && /must agree/.test(binding.reason)],
  ];
  for (const [altered, valid] of cases) {
    assert.notEqual(altered, xml);
    const bytes = createZip([...parts].map(([part, data]) => ({ name: part, data: part === name ? new TextEncoder().encode(altered) : data })));
    const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v3" });
    try {
      assert.equal(session.report.chartValues.length, 6);
      assert.ok(session.report.chartValues.every(valid));
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]);
      chart.series[0].values[1] = 3;
      chart.series[1].values[1] = 900;
      slide.replaceElement(chart.id, chart);
      assert.notEqual((await session.prepareSave({ mode: "edit" })).status, "ready");
    } finally { session.dispose(); }
  }
});

test("PowerPoint resave keeps the hidden bottom secondary category axis and both series writable", async () => {
  const bytes = new Uint8Array(await readFile(new URL("../tests/fixtures/combination-powerpoint-resave.pptx", import.meta.url)));
  const evidence = JSON.parse(await readFile(new URL("../docs/testing/combination-powerpoint-resave-evidence.json", import.meta.url), "utf8"));
  assert.equal(createHash("sha256").update(bytes).digest("hex"), evidence.saved.sha256);
  assert.notEqual(evidence.sourceSha256, evidence.saved.sha256);
  const options = { projection: "content", chartEditSaveProfile: "existing-values-v3" };
  const session = await importPptx(bytes, options);
  try {
    const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]);
    assert.equal(chart.chartType, "combo");
    assert.deepEqual(chart.categories, ["Jan", "Feb", "Mar"]);
    assert.deepEqual(chart.series.map(({ kind, values }) => [kind, values]), [["bar", [2, 7, 8]], ["line", [800, 400, 600]]]);
    assert.deepEqual([chart.axes.value.scale.max, chart.axes.secondaryValue.scale.max], [10, 1000]);
    assert.equal(session.report.chartValues.length, 6);
    assert.ok(session.report.chartValues.every(binding => binding.status === "bound"));
    const noOp = await session.prepareSave({ mode: "no-op" });
    assert.equal(noOp.status, "ready");
    assert.deepEqual((await session.save(noOp.plan)).bytes, bytes);
    chart.series[0].values[1] = 6;
    chart.series[1].values[1] = 350;
    slide.replaceElement(chart.id, chart);
    const plan = await session.prepareSave({ mode: "edit" });
    assert.equal(plan.status, "ready", JSON.stringify(plan.report?.diagnostics));
    const reopened = await importPptx((await session.save(plan.plan)).bytes, options);
    try {
      assert.deepEqual(reopened.document.slides[0].elements[0].series.map(series => series.values[1]), [6, 350]);
      assert.deepEqual(reopened.report.chartValues.filter(binding => binding.pointIndex === 1).map(binding => [binding.workbook.cellReference, binding.status]), [["B3", "bound"], ["C3", "bound"]]);
    } finally { reopened.dispose(); }
  } finally { session.dispose(); }
});
