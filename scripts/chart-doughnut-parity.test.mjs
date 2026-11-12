import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";

const close = (a, b) => assert.ok(Math.abs(a - b) < 1e-9, `${a} != ${b}`);
function deck(options = {}) {
  const p = createPresentation({ id: "doughnut" });
  p.addSlide({ id: "slide" }).addElement({ id: "chart", type: "chart", chartType: "doughnut", box: { x: 0, y: 0, width: 400, height: 300 }, categories: ["A", "B", "C"], series: [{ name: "inner", values: [1, 3, null], pointColors: ["#FF0000", "#0000FF", "#AAAAAA"] }, { name: "outer", values: [3, 1, 0], pointColors: ["#00AA00", "#FFAA00", "#CCCCCC"] }], holeSize: 50, firstSliceAngle: 90, ...options });
  return p;
}

test("multi-ring doughnut geometry, source data and three fresh round trips", async () => {
  const p = deck(); assert.deepEqual(validatePresentation(p), []);
  const normalized = normalizePresentation(p).slides[0].elements[0]; assert.equal(normalized.holeSize, 50); assert.equal(normalized.firstSliceAngle, 90);
  const rings = resolveLayout(p).slides[0].elements[0].chartLayout.doughnut.rings; assert.equal(rings.length, 2); close(rings[0].outerRadius, rings[1].innerRadius); close(rings[0].innerRadius / rings[1].outerRadius, .5); close(rings[0].slices[0].startAngle, 0); close(rings[0].slices[0].endAngle, Math.PI / 2); close(rings[1].slices[0].endAngle, 3 * Math.PI / 2); close(rings[0].slices[2].startAngle, rings[0].slices[2].endAngle);
  const svg = (await renderPresentationToSvg(p)).slides[0].svg; assert.equal((svg.match(/data-chart-role="doughnut-slice"/g) ?? []).length, 4); assert.doesNotMatch(svg, /data-chart-role="(?:category-axis|value-axis|pie-slice)"/); assert.doesNotMatch(svg, /NaN|Infinity/);
  let bytes = (await generatePptx(p)).bytes;
  for (let round = 0; round < 3; round++) {
    const parts = readGeneratedZip(bytes), xml = parts.get("ppt/charts/chart1.xml").toString();
    assert.match(xml, /<c:doughnutChart>/); assert.match(xml, /<c:firstSliceAng val="90"\/><c:holeSize val="50"\/>/); assert.equal((xml.match(/<c:ser>/g) ?? []).length, 2);
    const book = readGeneratedZip(parts.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")), sheet = book.get("xl/worksheets/sheet1.xml").toString(); assert.match(sheet, /<c r="B4"\/>/); assert.match(sheet, /<c r="C4"><v>0<\/v><\/c>/);
    const session = await importPptx(bytes, { projection: "content" });
    try {
      const chart = session.document.slides[0].elements[0]; assert.equal(chart.chartType, "doughnut"); assert.equal(chart.holeSize, 50); assert.equal(chart.firstSliceAngle, 90); assert.deepEqual(chart.series.map(s => s.values), [[1, 3, null], [3, 1, 0]]);
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual((await session.save(plan.plan)).bytes, bytes); bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
});

test("hole size, angle, values, colors, axes and unsupported modes are validated", () => {
  for (const [key, values, code] of [["holeSize", [0, 91, 50.5, NaN], "invalid-chart-hole-size"], ["firstSliceAngle", [-1, 361, 1.5, NaN], "invalid-chart-first-slice-angle"], ["axes", [{}], "invalid-chart-axis"], ["blankDisplay", ["span", "bad"], "invalid-chart-blank-display"]]) for (const value of values) assert.ok(validatePresentation(deck({ [key]: value })).some(d => d.code === code), `${key}=${value}`);
  for (const [key, values] of [["holeSize", [1, 9, 10, 90]], ["firstSliceAngle", [0, 360]]]) for (const value of values) assert.deepEqual(validatePresentation(deck({ [key]: value })), []);
  assert.ok(validatePresentation(deck({ series: [{ name: "bad", values: [1, -1, 0] }] })).some(d => d.code === "pie-negative-value"));
  assert.ok(validatePresentation(deck({ chartType: "pie", series: [{ name: "S", values: [1, 2, 3] }], firstSliceAngle: 10 })).some(d => d.code === "invalid-chart-geometry"));
});

test("full rings, all blanks and huge finite values produce finite geometry", async () => {
  for (const values of [[1, 0, 0], [null, null, null], [Number.MAX_VALUE, Number.MAX_VALUE, 0]]) {
    const p = deck({ series: [{ name: "S", values }], legend: { visible: false } }); const chartLayout = resolveLayout(p).slides[0].elements[0].chartLayout; assert.ok([chartLayout.valueScale.min, chartLayout.valueScale.max, chartLayout.valueScale.majorUnit, ...chartLayout.valueScale.ticks].every(Number.isFinite)); const layout = chartLayout.doughnut;
    for (const slice of layout.rings[0].slices) assert.ok(Number.isFinite(slice.startAngle) && Number.isFinite(slice.endAngle));
    const svg = (await renderPresentationToSvg(p)).slides[0].svg; assert.doesNotMatch(svg, /NaN|Infinity/);
    const count = (svg.match(/data-chart-role="doughnut-slice"/g) ?? []).length; assert.equal(count, values.every(v => v === null) ? 0 : values[0] === Number.MAX_VALUE ? 2 : 1);
    if (values[0] === 1) assert.match(svg, /A[^<]* A[^<]* A[^<]* A/);
  }
});

test("ambiguous geometry stays source-only and sparse numeric edits are blocked", async () => {
  const bytes = (await generatePptx(deck())).bytes;
  for (const token of ['<c:holeSize val="50"/>']) {
    const parts = readGeneratedZip(bytes), name = "ppt/charts/chart1.xml"; parts.set(name, Buffer.from(parts.get(name).toString().replace(token, "")));
    const session = await importPptx(zip([...parts].map(([name, body]) => ({ name, body }))).bytes, { projection: "content" });
    try { assert.ok(!session.document.slides[0].elements.some(e => e.type === "chart")); } finally { session.dispose(); }
  }
  const dense = (await generatePptx(deck({ series: [{ name: "S", values: [1, null, 3] }] }))).bytes;
  const session = await importPptx(dense, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[0] = 7; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); } finally { session.dispose(); }
});


test("browser matches Node projection and draws actual colored rings with a clear center", async t => {
  const p = deck({ firstSliceAngle: 0, legend: { visible: false } });
  const bytes = (await generatePptx(p)).bytes;
  const node = await importPptx(bytes, { projection: "content" }); let expected;
  try { expected = { chart: normalizePresentation(node.document).slides[0].elements[0], svg: (await renderPresentationToSvg(node.document)).slides[0].svg }; } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ["core", "layout", "pptx-exporter", "pptx-parser", "svg-renderer"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const actual = await page.evaluate(async input => {
    const session = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
    try { return { chart: PPTKit.normalizePresentation(session.document).slides[0].elements[0], svg: (await PPTKit.renderPresentationToSvg(session.document)).slides[0].svg }; } finally { session.dispose(); }
  }, Array.from(bytes));
  assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
  const layout = resolveLayout(p).slides[0].elements[0].chartLayout.doughnut, center = layout.center;
  const points = [center, { x: center.x + layout.rings[0].innerRadius * .8, y: center.y }, { x: center.x + (layout.rings[0].innerRadius + layout.rings[0].outerRadius) / 2 / Math.SQRT2, y: center.y - (layout.rings[0].innerRadius + layout.rings[0].outerRadius) / 2 / Math.SQRT2 }, { x: center.x + (layout.rings[1].innerRadius + layout.rings[1].outerRadius) / 2 / Math.SQRT2, y: center.y - (layout.rings[1].innerRadius + layout.rings[1].outerRadius) / 2 / Math.SQRT2 }];
  const colors = await page.evaluate(async ({ svg, points }) => {
    const img = new Image(); img.src = `data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}`; await img.decode(); const canvas = document.createElement("canvas"); canvas.width = img.naturalWidth; canvas.height = img.naturalHeight; const ctx = canvas.getContext("2d"); ctx.drawImage(img, 0, 0); return points.map(p => Array.from(ctx.getImageData(Math.round(p.x), Math.round(p.y), 1, 1).data));
  }, { svg: expected.svg, points });
  assert.deepEqual(colors[0], colors[1]); assert.deepEqual(colors[2], [255, 0, 0, 255]); assert.deepEqual(colors[3], [0, 170, 0, 255]);
});

test("strict and transitional literal doughnuts project with their declared geometry", async () => {
  const { chartFixture, chartXml, series, encodeChart } = await import("../packages/pptx-parser/test/chart-fixture.mjs");
  const xml = chartXml({ family: "pie", data: series({ values: [2, 5] }), familyExtra: '<c:firstSliceAng val="30"/><c:holeSize val="65"/>', after: '<c:dispBlanksAs val="gap"/>' }).replaceAll("pieChart", "doughnutChart");
  for (const strict of [false, true]) {
    const source = encodeChart(chartFixture({ xml, strict })), session = await importPptx(source, { projection: "content" });
    try { const chart = session.document.slides[0].elements[0]; assert.equal(chart.chartType, "doughnut"); assert.equal(chart.holeSize, 65); assert.equal(chart.firstSliceAngle, 30); assert.deepEqual(chart.series[0].values, [2, 5]); const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source)); } finally { session.dispose(); }
  }
});



test("dense multi-ring numeric edits preserve unrelated payloads and update exact workbook cells over three rounds", async () => {
  let bytes = (await generatePptx(deck({ series: [{ name: "inner", values: [1, 2, 3] }, { name: "outer", values: [4, 5, 6] }] }))).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try {
      assert.equal(session.report.chartValues.length, 6); assert.ok(session.report.chartValues.every(v => v.status === "bound"));
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]), before = readGeneratedZip(bytes);
      chart.series[0].values[1] = 10 + round; chart.series[1].values[2] = 20 + round; slide.replaceElement(chart.id, chart);
      const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report));
      bytes = (await session.save(plan.plan)).bytes; const after = readGeneratedZip(bytes);
      const changed = [...before].filter(([name, payload]) => !payload.equals(after.get(name))).map(([name]) => name);
      assert.equal(changed.length, 2); assert.ok(changed.some(name => name.startsWith("ppt/charts/"))); assert.ok(changed.some(name => name.startsWith("ppt/embeddings/")));
      const xml = after.get("ppt/charts/chart1.xml").toString(), caches = [...xml.matchAll(/<c:numCache>(.*?)<\/c:numCache>/gs)].map(v => v[1]);
      assert.match(caches[0], new RegExp(`<c:pt idx="1"><c:v>${10 + round}</c:v></c:pt>`)); assert.match(caches[1], new RegExp(`<c:pt idx="2"><c:v>${20 + round}</c:v></c:pt>`));
      const book = readGeneratedZip(after.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")), sheet = book.get("xl/worksheets/sheet1.xml").toString();
      assert.ok(sheet.includes(`<c r="B3"><v>${10 + round}</v></c>`)); assert.ok(sheet.includes(`<c r="C4"><v>${20 + round}</v></c>`));
      const reimport = await importPptx(bytes, { projection: "content" }); try { assert.deepEqual(reimport.document.slides[0].elements[0].series.map(s => s.values), [[1, 10 + round, 3], [4, 5, 20 + round]]); } finally { reimport.dispose(); }
    } finally { session.dispose(); }
  }
});

test("dense doughnut source edits agree in Node and browser and mismatched cells remain blocked", async t => {
  const source = (await generatePptx(deck({ series: [{ name: "inner", values: [1, 2, 3] }, { name: "outer", values: [4, 5, 6] }] }))).bytes, options = { projection: "content", chartEditSaveProfile: "existing-values-v1" };
  const node = await importPptx(source, options); let expected;
  try { const slide = node.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[1].values[1] = 42; slide.replaceElement(chart.id, chart); const plan = await node.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); expected = [...(await node.save(plan.plan)).bytes]; } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  const actual = await page.evaluate(async ({ source, options }) => { const session = await PPTKit.importPptx(Uint8Array.from(source), options); try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[1].values[1] = 42; slide.replaceElement(chart.id, chart); const plan = await session.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw new Error(JSON.stringify(plan.report)); return [...(await session.save(plan.plan)).bytes]; } finally { session.dispose(); } }, { source: [...source], options });
  assert.deepEqual(actual, expected); assert.deepEqual(requests, []);
  const parts = readGeneratedZip(source), bookName = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx", book = readGeneratedZip(parts.get(bookName)), sheetName = "xl/worksheets/sheet1.xml";
  book.set(sheetName, Buffer.from(book.get(sheetName).toString().replace('<c r="C3"><v>5</v></c>', '<c r="C3"><v>99</v></c>'))); parts.set(bookName, zip([...book].map(([name, body]) => ({ name, body }))).bytes);
  const session = await importPptx(zip([...parts].map(([name, body]) => ({ name, body }))).bytes, options);
  try { const target = session.report.chartValues.find(v => v.seriesIndex === 1 && v.pointIndex === 1); assert.equal(target.status, "unmapped"); const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[1].values[1] = 42; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); } finally { session.dispose(); }
});

test("an unrelated text edit preserves sparse doughnut cache and workbook bytes", async () => {
  const p = deck(); p.slides[0].addElement({ id: "label", type: "text", box: { x: 0, y: 310, width: 200, height: 30 }, content: "before" });
  const source = (await generatePptx(p)).bytes, session = await importPptx(source, { projection: "content", editSaveProfile: "text-runs-v1", chartEditSaveProfile: "existing-values-v1" });
  try {
    const slide = session.document.slides[0], label = structuredClone(slide.elements.find(e => e.type === "text")); label.content[0].runs[0].text = "after"; slide.replaceElement(label.id, label);
    const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report));
    const saved = (await session.save(plan.plan)).bytes, before = readGeneratedZip(source), after = readGeneratedZip(saved);
    for (const name of ["ppt/charts/chart1.xml", "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"]) assert.deepEqual(after.get(name), before.get(name));
    const reopened = await importPptx(saved, { projection: "content" }); try { assert.deepEqual(reopened.document.slides[0].elements.find(e => e.type === "chart").series.map(s => s.values), [[1, 3, null], [3, 1, 0]]); } finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test("existing-value profile rejects geometry changes without touching source", async () => {
  const source = (await generatePptx(deck({ series: [{ name: "S", values: [1, 2, 3] }] }))).bytes;
  for (const [key, value] of [["holeSize", 60], ["firstSliceAngle", 45]]) {
    const session = await importPptx(source, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try { const slide = session.document.slides[0], original = structuredClone(slide.elements[0]), chart = structuredClone(original); chart[key] = value; slide.replaceElement(chart.id, chart); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); slide.replaceElement(original.id, original); const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source)); } finally { session.dispose(); }
  }
});


test("omitted first slice angle projects as zero and remains absent through numeric source edits", async () => {
  const initial = (await generatePptx(deck({ series: [{ name: "S", values: [1, 2, 3] }] }))).bytes, parts = readGeneratedZip(initial), name = "ppt/charts/chart1.xml";
  parts.set(name, Buffer.from(parts.get(name).toString().replace('<c:firstSliceAng val="90"/>', "")));
  let bytes = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try {
      const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.equal(chart.firstSliceAngle, 0); assert.equal(chart.holeSize, 50); assert.ok(session.report.chartValues.every(v => v.status === "bound"));
      assert.ok(session.report.findings.some(f => f.reason.includes("Absent first slice angle")));
      const noop = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(bytes));
      chart.series[0].values[0] = 10 + round; slide.replaceElement(chart.id, chart); const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report)); bytes = (await session.save(plan.plan)).bytes;
      const output = readGeneratedZip(bytes); assert.doesNotMatch(output.get(name).toString(), /<c:firstSliceAng/); assert.ok(readGeneratedZip(output.get("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx")).get("xl/worksheets/sheet1.xml").toString().includes(`<c r="B2"><v>${10 + round}</v></c>`));
    } finally { session.dispose(); }
  }
});

test("Office-compatible out-of-range hole sizes stay source-only without preventing sibling import", async () => {
  const fixture = deck(); fixture.slides[0].addElement({ id: "neighbor", type: "text", box: { x: 0, y: 310, width: 200, height: 30 }, content: "neighbor" });
  const source = (await generatePptx(fixture)).bytes;
  for (const value of [0, 91, 255]) {
    const parts = readGeneratedZip(source), name = "ppt/charts/chart1.xml"; parts.set(name, Buffer.from(parts.get(name).toString().replace('<c:holeSize val="50"/>', `<c:holeSize val="${value}"/>`))); const candidate = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
    const session = await importPptx(candidate, { projection: "content" }); try { assert.equal(session.document.slides.length, 1); assert.ok(session.document.slides[0].elements.some(e => e.type === "text")); assert.ok(!session.document.slides[0].elements.some(e => e.type === "chart")); assert.ok(session.report.findings.some(f => f.reason.includes("outside Core's supported"))); const noop = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(candidate)); } finally { session.dispose(); }
  }
});

test("missing angle and unsupported hole size behave consistently in both OOXML dialects", async () => {
  const { chartFixture, chartXml, series, encodeChart } = await import("../packages/pptx-parser/test/chart-fixture.mjs");
  for (const strict of [false, true]) for (const holeSize of [0, 1, 9, 10, 90, 91]) {
    const xml = chartXml({ family: "pie", data: series({ values: [2, 5] }), familyExtra: `<c:holeSize val="${holeSize}"/>`, after: '<c:dispBlanksAs val="gap"/>' }).replaceAll("pieChart", "doughnutChart");
    const source = encodeChart(chartFixture({ xml, strict })), session = await importPptx(source, { projection: "content" });
    try {
      const chart = session.document.slides[0].elements.find(e => e.type === "chart");
      if (holeSize >= 1 && holeSize <= 90) { assert.equal(chart.chartType, "doughnut"); assert.equal(chart.firstSliceAngle, 0); assert.equal(chart.holeSize, holeSize); }
      else { assert.equal(chart, undefined); assert.ok(session.report.findings.some(f => f.reason.includes("outside Core's supported"))); }
      const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source));
    } finally { session.dispose(); }
  }
});

test("omitted-angle numeric edit agrees across Node and browser without inserting angle", async t => {
  const original = (await generatePptx(deck({ series: [{ name: "S", values: [1, 2, 3] }] }))).bytes, parts = readGeneratedZip(original), key = "ppt/charts/chart1.xml";
  parts.set(key, Buffer.from(parts.get(key).toString().replace('<c:firstSliceAng val="90"/>', ""))); const source = zip([...parts].map(([name, body]) => ({ name, body }))).bytes, options = { projection: "content", chartEditSaveProfile: "existing-values-v1" };
  const session = await importPptx(source, options); let expected;
  try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[1] = 12; slide.replaceElement(chart.id, chart); const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); expected = [...(await session.save(plan.plan)).bytes]; } finally { session.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  const actual = await page.evaluate(async ({ source, options }) => { const s = await PPTKit.importPptx(Uint8Array.from(source), options); try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]); if (chart.firstSliceAngle !== 0) throw Error("angle"); chart.series[0].values[1] = 12; slide.replaceElement(chart.id, chart); const plan = await s.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw Error(JSON.stringify(plan.report)); return [...(await s.save(plan.plan)).bytes]; } finally { s.dispose(); } }, { source: [...source], options });
  assert.deepEqual(actual, expected); assert.deepEqual(requests, []); assert.doesNotMatch(readGeneratedZip(Uint8Array.from(actual)).get(key).toString(), /<c:firstSliceAng/);
});

test("percentage hole-size lexical forms project in both dialects and retain original source spelling", async () => {
  const { chartFixture, chartXml, series, encodeChart } = await import("../packages/pptx-parser/test/chart-fixture.mjs");
  for (const strict of [false, true]) for (const [lexical, expected] of [["1%", 1], ["009%", 9], ["010%", 10], ["50%", 50], ["090%", 90]]) {
    const xml = chartXml({ family: "pie", data: series({ values: [2, 5] }), familyExtra: `<c:holeSize val="${lexical}"/>`, after: '<c:dispBlanksAs val="gap"/>' }).replaceAll("pieChart", "doughnutChart");
    const source = encodeChart(chartFixture({ xml, strict })), session = await importPptx(source, { projection: "content" });
    try {
      const chart = session.document.slides[0].elements.find(e => e.type === "chart");
      if (expected === undefined) assert.equal(chart, undefined); else { assert.equal(chart.holeSize, expected); assert.equal(chart.firstSliceAngle, 0); }
      const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source));
    } finally { session.dispose(); }
  }
  const original = (await generatePptx(deck({ series: [{ name: "S", values: [1, 2, 3] }] }))).bytes, parts = readGeneratedZip(original), key = "ppt/charts/chart1.xml";
  parts.set(key, Buffer.from(parts.get(key).toString().replace('<c:holeSize val="50"/>', '<c:holeSize val="050%"/>')));
  let source = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
  for (let round = 0; round < 3; round++) {
    const s = await importPptx(source, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]); assert.equal(chart.holeSize, 50); if (round === 0) { assert.ok(s.report.findings.some(f => f.reason.includes("Percentage-form hole size"))); const fresh = readGeneratedZip((await generatePptx(s.document)).bytes).get(key).toString(); assert.match(fresh, /<c:holeSize val="50"\/>/); } chart.series[0].values[0] = 5 + round; slide.replaceElement(chart.id, chart); const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report)); source = (await s.save(plan.plan)).bytes; assert.match(readGeneratedZip(source).get(key).toString(), /<c:holeSize val="050%"\/>/); } finally { s.dispose(); }
  }
});

test("empty angle value projects to zero while an empty hole value stays source-only in both dialects", async () => {
  const { chartFixture, chartXml, series, encodeChart } = await import("../packages/pptx-parser/test/chart-fixture.mjs");
  for (const strict of [false, true]) for (const emptyHole of [false, true]) {
    const xml = chartXml({ family: "pie", data: series({ values: [2, 5] }), familyExtra: `<c:firstSliceAng/>${emptyHole ? "<c:holeSize/>" : '<c:holeSize val="50"/>'}`, after: '<c:dispBlanksAs val="gap"/>' }).replaceAll("pieChart", "doughnutChart");
    const source = encodeChart(chartFixture({ xml, strict })), session = await importPptx(source, { projection: "content" });
    try { const chart = session.document.slides[0].elements.find(e => e.type === "chart"); if (emptyHole) { assert.equal(chart, undefined); assert.ok(session.report.findings.some(f => f.reason.includes("hole size has no value"))); } else { assert.equal(chart.firstSliceAngle, 0); assert.ok(session.report.findings.some(f => f.reason.includes("angle without a value"))); } const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source)); } finally { session.dispose(); }
  }
  const original = (await generatePptx(deck({ series: [{ name: "S", values: [1, 2, 3] }] }))).bytes, parts = readGeneratedZip(original), key = "ppt/charts/chart1.xml";
  parts.set(key, Buffer.from(parts.get(key).toString().replace('<c:firstSliceAng val="90"/>', '<c:firstSliceAng/>'))); let source = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(source, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
    try { const slide = session.document.slides[0], chart = structuredClone(slide.elements[0]); assert.equal(chart.firstSliceAngle, 0); chart.series[0].values[1] = round + 10; slide.replaceElement(chart.id, chart); const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report)); source = (await session.save(plan.plan)).bytes; assert.match(readGeneratedZip(source).get(key).toString(), /<c:firstSliceAng\/>/); } finally { session.dispose(); }
  }
});

test("empty angle attribute source edits agree across Node and browser", async t => {
  const original = (await generatePptx(deck({ series: [{ name: "S", values: [1, 2, 3] }] }))).bytes, parts = readGeneratedZip(original), key = "ppt/charts/chart1.xml";
  parts.set(key, Buffer.from(parts.get(key).toString().replace('<c:firstSliceAng val="90"/>', '<c:firstSliceAng/>'))); const source = zip([...parts].map(([name, body]) => ({ name, body }))).bytes, options = { projection: "content", chartEditSaveProfile: "existing-values-v1" };
  const node = await importPptx(source, options); let expected;
  try { const slide = node.document.slides[0], chart = structuredClone(slide.elements[0]); chart.series[0].values[1] = 12; slide.replaceElement(chart.id, chart); const plan = await node.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); expected = [...(await node.save(plan.plan)).bytes]; } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); }); await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  const actual = await page.evaluate(async ({ source, options }) => { const s = await PPTKit.importPptx(Uint8Array.from(source), options); try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]); if (chart.firstSliceAngle !== 0) throw Error("angle"); chart.series[0].values[1] = 12; slide.replaceElement(chart.id, chart); const plan = await s.prepareSave({ mode: "edit" }); if (plan.status !== "ready") throw new Error(JSON.stringify(plan.report)); return [...(await s.save(plan.plan)).bytes]; } finally { s.dispose(); } }, { source: [...source], options });
  assert.deepEqual(actual, expected); assert.deepEqual(requests, []); assert.match(readGeneratedZip(Uint8Array.from(actual)).get(key).toString(), /<c:firstSliceAng\/>/);
});

test("small Office hole sizes retain exact geometry and numeric sources across three rounds", async () => {
  for (const holeSize of [1, 5, 9]) {
    const fixture = deck({ holeSize, firstSliceAngle: 0, series: [{ name: "S", values: [1, 2, 3] }], legend: { visible: false } });
    const ring = resolveLayout(fixture).slides[0].elements[0].chartLayout.doughnut.rings[0]; close(ring.innerRadius / ring.outerRadius, holeSize / 100);
    let bytes = (await generatePptx(fixture)).bytes;
    for (let round = 0; round < 3; round++) {
      const parts = readGeneratedZip(bytes), xml = parts.get("ppt/charts/chart1.xml").toString(); assert.ok(xml.includes(`<c:holeSize val="${holeSize}"/>`));
      const s = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
      try { const slide = s.document.slides[0], chart = structuredClone(slide.elements[0]); assert.equal(chart.holeSize, holeSize); assert.deepEqual(chart.series[0].values, [1 + round, 2, 3]); const noop = await s.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await s.save(noop.plan)).bytes), Buffer.from(bytes)); chart.series[0].values[0]++; slide.replaceElement(chart.id, chart); const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report)); bytes = (await s.save(plan.plan)).bytes; } finally { s.dispose(); }
    }
  }
});

test("small-hole export reports Office extension with scoped element identity", async () => {
  for (const holeSize of [1, 5, 9, 10, 90]) {
    const p = deck({ holeSize }), output = await generatePptx(p), warnings = output.warnings.filter(w => w.code === "doughnut-small-hole-office-extension");
    if (holeSize < 10) { assert.equal(warnings.length, 1); assert.equal(warnings[0].slideId, "slide"); assert.equal(warnings[0].elementId, "chart"); assert.equal(warnings[0].path, "slides.0.elements.0.holeSize"); } else assert.deepEqual(warnings, []);
  }
});

test("small-hole warnings report nested chart paths and remain detached per export", async () => {
  const p = createPresentation({ id: "nested" }); p.addSlide({ id: "s" }).addElement({ id: "g", type: "group", box: { x: 0, y: 0, width: 400, height: 300 }, coordinateSize: { width: 400, height: 300 }, children: [{ id: "nested-chart", type: "chart", chartType: "doughnut", box: { x: 0, y: 0, width: 300, height: 200 }, holeSize: 1, categories: ["A"], series: [{ name: "S", values: [1] }] }] });
  const first = await generatePptx(p), second = await generatePptx(p), warnings = first.warnings.filter(w => w.code === "doughnut-small-hole-office-extension");
  assert.equal(warnings.length, 1); assert.equal(warnings[0].path, "slides.0.elements.0.children.0.holeSize"); assert.equal(warnings[0].slideId, "s"); assert.equal(warnings[0].elementId, "nested-chart"); assert.deepEqual(first.warnings, second.warnings); assert.deepEqual(first.bytes, second.bytes);
});

test("small-hole authoring and export warnings agree between Node and browser", async t => {
  const source = (await generatePptx(deck({ holeSize: 1 }))).bytes, node = await importPptx(source, { projection: "content" }); let expected;
  try { const output = await generatePptx(node.document); expected = { warnings: output.warnings, bytes: [...output.bytes], svg: (await renderPresentationToSvg(node.document)).slides[0].svg }; } finally { node.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); }); for (const name of ["core", "layout", "pptx-exporter", "pptx-parser", "svg-renderer"]) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, "utf8") });
  const actual = await page.evaluate(async source => { const session = await PPTKit.importPptx(Uint8Array.from(source), { projection: "content" }); try { const output = await PPTKit.generatePptx(session.document); return { warnings: output.warnings, bytes: [...output.bytes], svg: (await PPTKit.renderPresentationToSvg(session.document)).slides[0].svg }; } finally { session.dispose(); } }, [...source]);
  assert.deepEqual(actual, expected); assert.equal(actual.warnings.filter(w => w.code === "doughnut-small-hole-office-extension").length, 1); assert.deepEqual(requests, []);
});

test("all supported small integer and percentage holes normalize identically", async () => {
  const { chartFixture, chartXml, series, encodeChart } = await import("../packages/pptx-parser/test/chart-fixture.mjs");
  for (let size = 1; size < 10; size++) for (const strict of [false, true]) for (const lexical of [String(size), `${size}%`]) {
    const xml = chartXml({ family: "pie", data: series({ values: [2, 5] }), familyExtra: `<c:firstSliceAng val="0"/><c:holeSize val="${lexical}"/>`, after: '<c:dispBlanksAs val="gap"/>' }).replaceAll("pieChart", "doughnutChart");
    const source = encodeChart(chartFixture({ xml, strict })), session = await importPptx(source, { projection: "content" });
    try { const chart = session.document.slides[0].elements[0]; assert.equal(chart.holeSize, size); const layout = resolveLayout(session.document).slides[0].elements[0].chartLayout.doughnut; close(layout.rings[0].innerRadius / layout.rings[0].outerRadius, size / 100); const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(source)); } finally { session.dispose(); }
  }
});

test("small-hole warning identifies master and layout owners", async () => {
  const chart = id => ({ id, type: "chart", chartType: "doughnut", box: { x: 0, y: 0, width: 100, height: 100 }, holeSize: 5, categories: ["A"], series: [{ name: "S", values: [1] }] });
  const p = createPresentation({ id: "masters", masters: [{ id: "master", name: "Master", elements: [chart("master-chart")] }] });
  p.defineSlideLayout({ id: "layout", name: "Layout", masterId: "master", elements: [chart("layout-chart")] }); p.addSlide({ id: "slide", layoutId: "layout" });
  const output = await generatePptx(p), warnings = output.warnings.filter(w => w.code === "doughnut-small-hole-office-extension");
  assert.deepEqual(warnings.map(({ masterId, layoutId, elementId, path }) => ({ masterId, layoutId, elementId, path })), [{ masterId: "master", layoutId: undefined, elementId: "master-chart", path: "masters.0.elements.0.holeSize" }, { masterId: undefined, layoutId: "layout", elementId: "layout-chart", path: "layouts.0.elements.0.holeSize" }]);
});
