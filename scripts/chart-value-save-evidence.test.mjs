// Optional independent evidence checks: requires Python 3, never launches native applications.
import assert from "node:assert/strict";
import test from "node:test";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, readFile, writeFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { createChartSaveFixture } from "../tests/fixtures/chart-save.mjs";
import { chartFixture, chartXml, series, workbook, encodeChart } from "../packages/pptx-parser/test/chart-fixture.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { zip } from "../packages/pptx-parser/test/zip-fixture.mjs";
const execute = promisify(execFile), script = "scripts/research/chart-value-save-oracle.py";
const pack = parts => zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
async function saveValues(bytes, edits) {
  const session = await importPptx(bytes, { projection: "content", chartEditSaveProfile: "existing-values-v1" });
  try {
    const charts = session.document.slides.flatMap(slide => slide.elements.filter(e => e.type === "chart").map(chart => ({ slide, chart })));
    for (const edit of edits) {
      const { slide, chart: original } = charts[edit.chartIndex], chart = slide.elements.find(e => e.id === original.id), changed = structuredClone(chart), index = slide.elements.indexOf(chart);
      changed.series[edit.seriesIndex].values[edit.pointIndex] = edit.value;
      slide.removeElement(chart.id); slide.insertElement(index, changed);
    }
    const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); return (await session.save(plan.plan)).bytes;
  } finally { session.dispose(); }
}
async function harness(t) {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptkit-chart-oracle-")); t.after(() => rm(dir, { recursive: true, force: true }));
  let counter = 0;
  return {
    async inspect(bytes) {
      const file = path.join(dir, `inspect-${counter++}.pptx`); await writeFile(file, bytes);
      return JSON.parse((await execute("python3", [script, "inspect", file], { maxBuffer: 4 * 1024 * 1024 })).stdout);
    },
    async verify(before, after, edits) {
      const prefix = path.join(dir, `compare-${counter++}`);
      await Promise.all([writeFile(prefix + "-before.pptx", before), writeFile(prefix + "-after.pptx", after), writeFile(prefix + ".json", JSON.stringify(Array.isArray(edits) ? edits.map(e => ({ ...e, value: Object.is(e.value, -0) ? "-0" : e.value })) : edits))]);
      return JSON.parse((await execute("python3", [script, "verify", prefix + "-before.pptx", prefix + "-after.pptx", prefix + ".json"], { maxBuffer: 4 * 1024 * 1024 })).stdout);
    },
  };
}

test("independent ZIP/XML oracle verifies both dialects and all families over three source-save rounds", async t => {
  const oracle = await harness(t);
  for (const strict of [false, true]) for (const family of ["bar", "line", "pie"]) {
    let bytes = encodeChart(chartFixture({ strict, xml: chartXml({ family, data: series({ refs: true, values: [0, 2.5] }), external: true }), book: workbook({ values: [0, 2.5] }) }));
    for (const value of [7.25, 0, 16.5]) {
      const edits = [{ chartIndex: 0, seriesIndex: 0, pointIndex: 1, value }], after = await saveValues(bytes, edits), result = await oracle.verify(bytes, after, edits);
      assert.equal(result.verified, true); assert.equal(result.charts[0].family, family); assert.equal(result.charts[0].series[0].values[1], value);
      assert.equal(result.changedOuterParts.length, 2); assert.ok(result.unchangedOuterPayloads > 0); bytes = after;
    }
  }
});

test("three-family multi-slide saves retain independent control chart/series and entire package inventory", async t => {
  const oracle = await harness(t); let bytes = (await generatePptx(createChartSaveFixture(createPresentation))).bytes;
  const first = await oracle.inspect(bytes); assert.deepEqual(first.charts.map(c => c.family), ["bar", "line", "pie", "bar"]);
  for (let round = 0; round < 3; round++) {
    const edits = [0, 1, 2].map(chartIndex => ({ chartIndex, seriesIndex: 0, pointIndex: 1, value: 10 + round * 3 + chartIndex }));
    const after = await saveValues(bytes, edits), result = await oracle.verify(bytes, after, edits);
    assert.equal(result.partCount, first.partCount); assert.equal(result.changedOuterParts.length, 6);
    assert.deepEqual(result.charts[3], first.charts[3]);
    for (const ci of [0, 1]) assert.deepEqual(result.charts[ci].series[1], first.charts[ci].series[1]); bytes = after;
  }
});

test("oracle rejects stale XLSX, wrong proposals, unapproved charts/styles/cells/parts, source mismatch and invalid identities", async t => {
  const oracle = await harness(t), before = (await generatePptx(createChartSaveFixture(createPresentation))).bytes;
  const edits = [{ chartIndex: 0, seriesIndex: 0, pointIndex: 1, value: 42 }], after = await saveValues(before, edits);
  const inventory = await oracle.inspect(before), chart = inventory.charts[0], sheet = chart.series[0].cells[1].worksheetPart, cell = chart.series[0].cells[1].reference;
  const mutate = edit => { const parts = readGeneratedZip(after); edit(parts); return pack(parts); };
  const old = readGeneratedZip(before);
  const cases = [
    mutate(parts => parts.set(chart.workbookPart, old.get(chart.workbookPart))),
    mutate(parts => parts.set(chart.part, Buffer.from(parts.get(chart.part).toString().replace('<c:chart>', '<c:chart><c:autoTitleDeleted val="1"/>')))),
    mutate(parts => parts.set(inventory.charts[3].part, Buffer.from(parts.get(inventory.charts[3].part).toString().replace('<c:v>5</c:v>', '<c:v>44</c:v>')))),
    mutate(parts => { const book = readGeneratedZip(parts.get(chart.workbookPart)); book.set(sheet, Buffer.from(book.get(sheet).toString().replace(`r="${cell}"`, `r="${cell}" s="7"`))); parts.set(chart.workbookPart, pack(book)); }),
    mutate(parts => { const book = readGeneratedZip(parts.get(chart.workbookPart)); book.set('unapproved.bin', Buffer.from([1, 2])); parts.set(chart.workbookPart, pack(book)); }),
    mutate(parts => parts.set('unapproved.bin', Buffer.from([1, 2]))),
  ];
  for (const bytes of cases) await assert.rejects(oracle.verify(before, bytes, edits), e => /ValueError/.test(e.stderr));
  await assert.rejects(oracle.verify(before, after, [{ ...edits[0], value: 43 }]), e => /Saved cache\/workbook disagree/.test(e.stderr));
  await assert.rejects(oracle.verify(before, after, [edits[0], edits[0]]), e => /Duplicate edit/.test(e.stderr));
  for (const spec of [[{ ...edits[0], chartIndex: -1 }], [{ ...edits[0], pointIndex: true }], [{ ...edits[0], chartIndex: 9999999 }], [{ ...edits[0], value: "NaN" }], []]) await assert.rejects(oracle.verify(before, after, spec), e => /ValueError/.test(e.stderr));
  const share = bytes => {
    const parts = readGeneratedZip(bytes), second = inventory.charts[1];
    const relationPart = path.posix.join(path.posix.dirname(second.part), "_rels", path.posix.basename(second.part) + ".rels");
    const target = path.posix.relative(path.posix.dirname(second.part), chart.workbookPart);
    parts.set(relationPart, Buffer.from(parts.get(relationPart).toString().replace(/Target="[^"]+\.xlsx"/, `Target="${target}"`)));
    return pack(parts);
  };
  await assert.rejects(oracle.verify(share(before), share(after), edits), e => /Edited workbook source is shared/.test(e.stderr));
  const badSource = readGeneratedZip(before); badSource.set(chart.workbookPart, readGeneratedZip(after).get(chart.workbookPart));
  await assert.rejects(oracle.verify(pack(badSource), after, edits), e => /Source cache\/workbook disagree/.test(e.stderr));
});

test("native evidence retains three numeric rounds and explicitly separates cache retention from write eligibility", async () => {
  const evidence = JSON.parse(await readFile("docs/testing/chart-value-save-evidence.json", "utf8"));
  assert.equal(evidence.schemaVersion, 1); assert.equal(evidence.rounds.length, 3);
  assert.match(evidence.application, /LibreOffice/);
  for (const round of evidence.rounds) {
    assert.equal(round.independent.verified, true); assert.equal(round.independent.charts.length, 4); assert.equal(round.independent.changedOuterParts.length, 6);
    assert.equal(round.native.inventory.charts.length, 4);
    assert.equal(round.native.noOpOriginalBytes, true);
    assert.equal(round.native.numericEditTrial.status, "blocked");
    assert.ok(round.native.numericEditTrial.diagnostics.includes("save-edit-unsupported"));
    for (const edit of round.edits) assert.equal(round.native.inventory.charts[edit.chartIndex].series[edit.seriesIndex].values[edit.pointIndex], edit.value);
    assert.ok(round.native.bindingEligibility.every(b => b.status === "unmapped" && /XLSX/.test(b.reason)));
    assert.match(round.sourceSha256, /^[a-f0-9]{64}$/); assert.match(round.savedSha256, /^[a-f0-9]{64}$/);
  }
  assert.equal(evidence.pdf.observation.pages, 4);
  assert.equal(evidence.interactiveEditData.powerpoint, "untested"); assert.equal(evidence.interactiveEditData.libreOffice, "untested");
});

test("oracle preserves signed zero agreement and rejects duplicate ZIP names and XML declarations outside scope", async t => {
  const oracle = await harness(t);
  const before = encodeChart(chartFixture({ xml: chartXml({ data: series({ refs: true }), external: true }), book: workbook() }));
  const edits = [{ chartIndex: 0, seriesIndex: 0, pointIndex: 1, value: -0 }], after = await saveValues(before, edits);
  const verified = await oracle.verify(before, after, edits); assert.ok(Object.is(verified.charts[0].series[0].values[1], -0));
  const mismatched = readGeneratedZip(after), book = readGeneratedZip(mismatched.get("deck/book.xlsx"));
  book.set("xl/sheets/one.xml", Buffer.from(book.get("xl/sheets/one.xml").toString().replace("<v>-0</v>", "<v>0</v>"))); mismatched.set("deck/book.xlsx", pack(book));
  await assert.rejects(oracle.verify(before, pack(mismatched), edits), e => /Saved cache\/workbook disagree/.test(e.stderr));
  const parts = readGeneratedZip(before), dtd = new Map(parts);
  dtd.set("deck/chart.xml", Buffer.from('<!DOCTYPE c:chartSpace [<!ENTITY x "unsafe">]>' + dtd.get("deck/chart.xml").toString()));
  await assert.rejects(oracle.inspect(pack(dtd)), e => /DTD\/entity declarations/.test(e.stderr));
  const deep = new Map(parts);
  deep.set("deck/chart.xml", Buffer.from(deep.get("deck/chart.xml").toString().replace("</c:chartSpace>", '<c:extLst><c:ext uri="depth">' + '<x xmlns="urn:test">'.repeat(130) + '</x>'.repeat(130) + '</c:ext></c:extLst></c:chartSpace>')));
  await assert.rejects(oracle.inspect(pack(deep)), e => /XML exceeds observation depth budget/.test(e.stderr));
  const duplicate = zip([...parts].map(([name, body]) => ({ name, body })).concat([{ name: "deck/chart.xml", body: parts.get("deck/chart.xml") }])).bytes;
  await assert.rejects(oracle.inspect(duplicate), e => /Duplicate ZIP names/.test(e.stderr));
});

test("one candidate updates several series and points in the same worksheet without losing earlier proposals", async t => {
  const oracle = await harness(t), before = (await generatePptx(createChartSaveFixture(createPresentation))).bytes;
  const edits = [{ chartIndex: 0, seriesIndex: 0, pointIndex: 0, value: -7.5 }, { chartIndex: 0, seriesIndex: 0, pointIndex: 2, value: 21.25 }, { chartIndex: 0, seriesIndex: 1, pointIndex: 1, value: 15.5 }];
  const after = await saveValues(before, edits), result = await oracle.verify(before, after, edits);
  assert.equal(result.changedOuterParts.length, 2);
  assert.deepEqual(result.charts[0].series.map(s => s.values), [[-7.5, 5, 21.25], [3, 15.5, 6]]);
  const original = await oracle.inspect(before);
  assert.deepEqual(result.charts.slice(1), original.charts.slice(1));
});

test("oracle compares original decimal lexemes instead of rounded inventory floats", async t => {
  const oracle = await harness(t), before = (await generatePptx(createChartSaveFixture(createPresentation))).bytes;
  const edit = { chartIndex: 0, seriesIndex: 0, pointIndex: 1, value: 42 }, after = await saveValues(before, [edit]);
  const chart = (await oracle.inspect(before)).charts[0], cell = chart.series[0].cells[1];
  const replace = (text, from, to) => { assert.ok(text.includes(from), from); return text.replace(from, to); };
  const mutate = (input, oldValue, cacheValue, cellValue) => {
    const parts = readGeneratedZip(input), book = readGeneratedZip(parts.get(chart.workbookPart));
    parts.set(chart.part, Buffer.from(replace(parts.get(chart.part).toString(), `<c:v>${oldValue}</c:v>`, `<c:v>${cacheValue}</c:v>`)));
    const prefix = `<c r="${cell.reference}"><v>`;
    book.set(cell.worksheetPart, Buffer.from(replace(book.get(cell.worksheetPart).toString(), `${prefix}${oldValue}</v>`, `${prefix}${cellValue}</v>`)));
    parts.set(chart.workbookPart, pack(book)); return pack(parts);
  };
  const collision = mutate(before, 5, '9007199254740993', '9007199254740992');
  await assert.rejects(oracle.verify(collision, after, [edit]), e => /Source cache\/workbook disagree/.test(e.stderr));
  const equivalent = mutate(before, 5, '9007199254740993', '9.007199254740993e15');
  assert.equal((await oracle.verify(equivalent, after, [edit])).verified, true);
  const badSaved = mutate(after, 42, '9007199254740993', '9007199254740992');
  await assert.rejects(oracle.verify(before, badSaved, [{ ...edit, value: '9007199254740992' }]), e => /Saved cache\/workbook disagree/.test(e.stderr));
});
