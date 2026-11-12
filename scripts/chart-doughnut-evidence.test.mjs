import assert from "node:assert/strict";
import test from "node:test";
import { readFile, mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";

test("native evidence verifies retained package and print", () => { assert.match(execFileSync("python3", ["scripts/research/verify-doughnut-powerpoint.py"], { encoding: "utf8" }), /"verified": true/); });

test("PowerPoint-saved doughnuts remain importable with original rings and values", async () => {
  const base = ".tmp/native-doughnut/powerpoint-TFHkIb/", bytes = await readFile(base + "doughnut-round-1.pptx");
  const session = await importPptx(bytes, { projection: "content" });
  try {
    const charts = session.document.slides.map(s => s.elements.find(e => e.type === "chart"));
    assert.deepEqual(charts.map(c => [c.chartType, c.holeSize, c.firstSliceAngle, c.series.map(s => s.values)]), [["doughnut", 50, 0, [[1, 3], [3, 1]]], ["doughnut", 30, 90, [[1, 3]]], ["doughnut", 70, 0, [[2, null]]]]);
    const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
  } finally { session.dispose(); }
});


test("actual Layout geometry agrees with independently measured PowerPoint print", async () => {
  const evidence = JSON.parse(await readFile("docs/testing/chart-doughnut-powerpoint-evidence.json", "utf8"));
  const bytes = await readFile(".tmp/native-doughnut/powerpoint-TFHkIb/doughnut.pptx"), session = await importPptx(bytes, { projection: "content" });
  try {
    const result = resolveLayout(session.document);
    for (const [index, slide] of result.slides.entries()) {
      const element = slide.elements.find(e => e.type === "chart"), layout = element.chartLayout.doughnut, sample = evidence.measurements.nativeVsPreview[index];
      assert.deepEqual([element.box.x + layout.center.x, element.box.y + layout.center.y], sample.previewCenter);
      assert.equal(layout.rings.at(-1).outerRadius, sample.previewRadius);
      assert.ok(Math.abs(layout.rings[0].innerRadius / layout.rings.at(-1).outerRadius * 100 - sample.holeSize) < 1e-9);
      assert.ok(sample.nativeCenter.every((value, i) => Math.abs(value - sample.previewCenter[i]) < .3));
      assert.ok(sample.nativeRadius.every(value => Math.abs(value - sample.previewRadius) < .3));
    }
  } finally { session.dispose(); }
});


test("native evidence verifier rejects altered measurements and hashes", async () => {
  const record = JSON.parse(await readFile("docs/testing/chart-doughnut-powerpoint-evidence.json", "utf8")), directory = await mkdtemp(path.join(tmpdir(), "pptkit-doughnut-evidence-"));
  try {
    for (const mutate of [r => { r.measurements.nativeVsPreview[0].previewRadius += 1; }, r => { r.captures[0].screenshotSha256 = "0".repeat(64); }]) {
      const bad = structuredClone(record); mutate(bad); const file = path.join(directory, "bad.json"); await writeFile(file, JSON.stringify(bad));
      assert.throws(() => execFileSync("python3", ["scripts/research/verify-doughnut-powerpoint.py", file], { stdio: "pipe" }));
    }
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test("PowerPoint missing-geometry observations are bounded and rechecked", async () => {
  assert.match(execFileSync("python3", ["scripts/research/verify-doughnut-defaults-powerpoint.py"], { encoding: "utf8" }), /"verified": true/);
  for (const file of ["defaults.pptx", "defaults-round-1.pptx"]) {
    const bytes = await readFile(`.tmp/native-doughnut-defaults/powerpoint-FE4zPv/${file}`), session = await importPptx(bytes, { projection: "content" });
    try {
      const charts = session.document.slides.map(slide => slide.elements.find(element => element.type === "chart"));
      assert.equal(charts[0].chartType, "doughnut"); assert.equal(charts[0].firstSliceAngle, 0); assert.equal(charts[0].holeSize, 50);
      assert.equal(charts[1], undefined); assert.equal(charts[2], undefined);
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes);
    } finally { session.dispose(); }
  }
});

test("omitted-geometry evidence rejects tampered saved values or PDF measurements", async () => {
  const record = JSON.parse(await readFile("docs/testing/chart-doughnut-defaults-powerpoint-evidence.json", "utf8")), directory = await mkdtemp(path.join(tmpdir(), "pptkit-doughnut-defaults-evidence-"));
  try {
    for (const mutate of [r => { r.packages[1].geometry[1].holeSize = 50; }, r => { r.centerHits[0].nearestRedPathPointDistance = 1; }]) {
      const bad = structuredClone(record); mutate(bad); const file = path.join(directory, "bad.json"); await writeFile(file, JSON.stringify(bad));
      assert.throws(() => execFileSync("python3", ["scripts/research/verify-doughnut-defaults-powerpoint.py", file], { stdio: "pipe" }));
    }
  } finally { await rm(directory, { recursive: true, force: true }); }
});

test("controlled percent-spelling PowerPoint repair warning is retained", async () => {
  assert.match(execFileSync("python3", ["scripts/research/verify-doughnut-percent-powerpoint.py"], { encoding: "utf8" }), /"powerPointRepairWarning": true/);
  const bytes = await readFile(".tmp/native-doughnut-percent/powerpoint-JYogRH/percent.pptx"), session = await importPptx(bytes, { projection: "content" });
  try { assert.equal(session.document.slides[0].elements.find(e => e.type === "chart").holeSize, 50); assert.ok(session.report.findings.some(f => f.reason.includes("Percentage-form hole size"))); const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes); } finally { session.dispose(); }
});

test("LibreOffice percent and integer print comparison preserves its observed geometry limits", () => {
  assert.match(execFileSync("python3", ["scripts/research/verify-doughnut-percent-libreoffice.py"], { encoding: "utf8" }), /"verified": true/);
});

test("PowerPoint empty-geometry sample is independently bounded and source-preserving", async () => {
  assert.match(execFileSync("python3", ["scripts/research/verify-doughnut-empty-geometry-powerpoint.py"], { encoding: "utf8" }), /"verified": true/);
  for (const file of ["empty.pptx", "empty-round-1.pptx"]) {
    const bytes = await readFile(`.tmp/native-doughnut-empty-geometry/powerpoint-aPY1ED/${file}`), session = await importPptx(bytes, { projection: "content" });
    try { const charts = session.document.slides.map(s => s.elements.find(e => e.type === "chart")); assert.equal(charts[0].firstSliceAngle, 0); assert.equal(charts[0].holeSize, 50); assert.equal(charts[1], undefined); assert.equal(charts[2], undefined); const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes); } finally { session.dispose(); }
  }
});

test("empty-geometry evidence rejects altered source or saved geometry records", async () => {
  const record = JSON.parse(await readFile("docs/testing/chart-doughnut-empty-geometry-powerpoint-evidence.json", "utf8")), directory = await mkdtemp(path.join(tmpdir(), "pptkit-empty-doughnut-evidence-"));
  try { for (const mutate of [r => { r.sourceSha256 = "0".repeat(64); }, r => { r.savedGeometry[1].holeSize = 10; }]) { const bad = structuredClone(record); mutate(bad); const file = path.join(directory, "bad.json"); await writeFile(file, JSON.stringify(bad)); assert.throws(() => execFileSync("python3", ["scripts/research/verify-doughnut-empty-geometry-powerpoint.py", file], { stdio: "pipe" })); } } finally { await rm(directory, { recursive: true, force: true }); }
});

test("PowerPoint small-hole source and saved packages import with exact values", async () => {
  assert.match(execFileSync("python3", ["scripts/research/verify-doughnut-small-hole-powerpoint.py"], { encoding: "utf8" }), /"verified": true/);
  for (const name of ["small-holes.pptx", "small-holes-round-1.pptx"]) {
    const bytes = await readFile(`.tmp/native-doughnut-small-hole/powerpoint-Kx43uN/${name}`), session = await importPptx(bytes, { projection: "content" });
    try { const charts = session.document.slides.map(s => s.elements.find(e => e.type === "chart")); assert.deepEqual(charts.map(c => c.holeSize), [1, 5, 9, 10]); assert.deepEqual(charts.map(c => c.series[0].values), [[1,3],[1,3],[1,3],[1,3]]); const plan = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(plan.plan)).bytes), bytes); } finally { session.dispose(); }
  }
});

test("small-hole native evidence rejects altered measurements and saved geometry", async () => {
  const record = JSON.parse(await readFile("docs/testing/chart-doughnut-small-hole-powerpoint-evidence.json", "utf8")), directory = await mkdtemp(path.join(tmpdir(), "pptkit-small-hole-evidence-"));
  try { for (const mutate of [r => { r.measurements[0].nativeInnerOuterRadiusRatio = .1; }, r => { r.packages[1].charts[0].holeSize = 10; }]) { const bad = structuredClone(record); mutate(bad); const file = path.join(directory, "bad.json"); await writeFile(file, JSON.stringify(bad)); assert.throws(() => execFileSync("python3", ["scripts/research/verify-doughnut-small-hole-powerpoint.py", file], { stdio: "pipe" })); } } finally { await rm(directory, { recursive: true, force: true }); }
});

test("small-hole layout ratios match independently measured native vector paths", async () => {
  const record = JSON.parse(await readFile("docs/testing/chart-doughnut-small-hole-powerpoint-evidence.json", "utf8")), bytes = await readFile(".tmp/native-doughnut-small-hole/powerpoint-Kx43uN/small-holes.pptx"), session = await importPptx(bytes, { projection: "content" });
  try { const slides = resolveLayout(session.document).slides; for (const [index, slide] of slides.entries()) { const chart = slide.elements.find(e => e.type === "chart"), ring = chart.chartLayout.doughnut.rings[0], native = record.measurements[index]; assert.equal(chart.holeSize, native.holeSize); assert.ok(Math.abs(ring.innerRadius / ring.outerRadius - native.nativeInnerOuterRadiusRatio) < .0001); } } finally { session.dispose(); }
});
