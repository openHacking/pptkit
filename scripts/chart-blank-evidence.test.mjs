import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
test("PowerPoint blank-chart observation binds cache, workbook and captured states", async t => {
  const verify = file => JSON.parse(execFileSync("python3", ["scripts/research/verify-blank-chart-powerpoint.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" }));
  assert.deepEqual(verify(), { verified: true, charts: 3, captureStates: 12, reopenRounds: 3 });
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-blank-chart-evidence-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/chart-blank-powerpoint-evidence.json", "utf8")), file = path.join(dir, "changed.json");
  for (const change of [e => { e.packages[1].sha256 = "0".repeat(64); }, e => { e.packages[3].sha256 = "0".repeat(64); }, e => { e.packages[0].charts[0].workbookPart = "wrong.xlsx"; }, e => { e.captures[1].points.BLANK = "0"; }, e => { e.captures[4].axSha256 = "0".repeat(64); }, e => { e.captures[11].axSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); change(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => verify(file));
  }
});

test("PowerPoint-saved blank charts import without replacing gaps with zero", async () => {
  let bytes = await readFile(".tmp/native-chart-blanks/powerpoint-AP8EmX/blank-charts-round-3.pptx");
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "content" });
    try {
      const charts = session.document.slides.map(slide => slide.elements.find(element => element.type === "chart"));
      assert.deepEqual(charts.map(chart => chart.chartType), ["bar", "line", "pie"]);
      for (const chart of charts) assert.deepEqual(chart.series[0].values, [4, null, 0, 3]);
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(bytes));
      bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
});

test("native workbook-only observation binds its missing caches and absent plot states", async t => {
  const verify = file => JSON.parse(execFileSync("python3", ["scripts/research/verify-workbook-only-chart-powerpoint.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" }));
  assert.deepEqual(verify(), { verified: true, charts: 3, numericCachesRemoved: 3 });
  const source = await readFile(".tmp/native-chart-blanks/powerpoint-AP8EmX/blank-workbook-only.pptx"), session = await importPptx(source, { projection: "content" });
  try {
    assert.deepEqual(session.document.slides.map(slide => slide.elements.find(e => e.type === "chart").series[0].values), [[4, null, 0, 3], [4, null, 0, 3], [4, null, 0, 3]]);
    assert.ok(session.report.findings.some(f => f.feature === "chart.data-source" && f.reason.includes("native first-open display may differ")));
    const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), source);
  } finally { session.dispose(); }
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-no-chart-cache-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/chart-workbook-only-powerpoint-evidence.json", "utf8")), file = path.join(dir, "bad.json");
  for (const change of [e => { e.fixtureSha256 = "0".repeat(64); }, e => { e.captures[1].axSha256 = "0".repeat(64); }, e => { e.originalSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); change(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => verify(file));
  }
});

test("PowerPoint blank display policy evidence binds native line geometry and saved workbooks", async t => {
  const verify = file => JSON.parse(execFileSync("python3", ["scripts/research/verify-blank-modes-powerpoint.py", ...(file ? [file] : [])], { encoding: "utf8", stdio: "pipe" }));
  assert.deepEqual(verify(), { verified: true, charts: 5, linePolicies: 3, pdfs: 2, reopenRounds: 1 });
  let bytes = await readFile(".tmp/native-chart-blanks/powerpoint-modes-vkUEcD/blank-modes-round-1.pptx");
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "content" });
    try {
      const charts = session.document.slides.map(slide => slide.elements.find(e => e.type === "chart"));
      assert.deepEqual(charts.map(chart => [chart.chartType, chart.blankDisplay]), [["line", "gap"], ["line", "zero"], ["line", "span"], ["bar", "zero"], ["pie", "zero"]]);
      for (const chart of charts) assert.deepEqual(chart.series[0].values, [4, null, 1, null, 3]);
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), Buffer.from(bytes)); bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
  }
  const dir = await mkdtemp(path.join(tmpdir(), "pptkit-blank-modes-evidence-")); t.after(() => rm(dir, { recursive: true, force: true }));
  const original = JSON.parse(await readFile("docs/testing/chart-blank-modes-powerpoint-evidence.json", "utf8")), file = path.join(dir, "changed.json");
  for (const change of [e => { e.packages[1].sha256 = "0".repeat(64); }, e => { e.pdfs[1].pdfSha256 = "0".repeat(64); }, e => { e.pdfs[0].linePages[1].segments[0][0][0]++; }, e => { e.captures[9].axSha256 = "0".repeat(64); }]) {
    const changed = structuredClone(original); change(changed); await writeFile(file, JSON.stringify(changed)); assert.throws(() => verify(file));
  }
});
