import assert from "node:assert/strict";
import test from "node:test";
import { createHash } from "node:crypto";
import { normalizePresentation } from "../../core/dist/index.js";
import { importPptx, compareOpcPackages, DEFAULT_SESSION_LIMITS } from "../dist/index.js";
import { analyzeTextEdits } from "../dist/import/analyze-text-edits.js";
import { table, cell, tableBody, tableRow, tableFixture, tableTransform } from "./table-fixture.mjs";
import { shape, body } from "./content-fixture.mjs";
import { notesFixture, notesShape } from "./notes-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { zip } from "./zip-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";

const profile = { projection: "content", editSaveProfile: "text-runs-v3" };
const code = expected => error => { assert.equal(error.code, expected); return true; };
const digest = bytes => createHash("sha256").update(bytes).digest("hex");
const open = (f = tableFixture(), options = {}) => importPptx(encodeStructure(f), { ...profile, ...options });
const all = elements => elements.flatMap(e => [e, ...(e.type === "group" ? all(e.children) : [])]);
function modify(session, sourceId, mutate) {
  const slide = session.document.slides[0], id = `${slide.id}-object-${sourceId}`;
  const top = slide.elements.find(e => all([e]).some(child => child.id === id)), index = slide.elements.indexOf(top);
  assert.ok(top);
  const copied = structuredClone(top); mutate(all([copied]).find(child => child.id === id));
  slide.removeElement(top.id); slide.insertElement(index, copied);
}
const edit = (session, row, cellIndex, text, sourceId = 2, p = 0, r = 0) => modify(session, sourceId, e => {
  const cell = e.rows[row].cells[cellIndex];
  if (typeof cell.content === "string") cell.content = text;
  else cell.content[p].runs[r].text = text;
});
async function save(session) {
  const plan = await session.prepareSave({ mode: "edit" });
  assert.equal(plan.status, "ready", JSON.stringify(plan.report));
  return { ...await session.save(plan.plan), plan: plan.plan };
}
function unchanged(before, after, selected = ["deck/z.xml"]) {
  const a = readGeneratedZip(before), b = readGeneratedZip(after);
  assert.deepEqual([...a.keys()].sort(), [...b.keys()].sort());
  for (const [name, bytes] of a) if (!selected.includes(name)) assert.equal(digest(b.get(name)), digest(bytes), name);
}
const merged = () => table({ columns: [30, 30, 40], rows: [
  tableRow([cell({ text: "Anchor", attributes: 'gridSpan="2" rowSpan="2"' }), cell({ body: "", attributes: 'hMerge="1"' }), cell({ text: "Top right" })]),
  tableRow([cell({ body: "", attributes: 'vMerge="1"' }), cell({ body: "", attributes: 'vMerge="1" hMerge="1"' }), cell({ text: "Bottom right" })], 30),
] });
const group = content => `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="900000" flipH="1"><a:off x="127000" y="254000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="63500" y="88900"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></p:grpSpPr>${content}</p:grpSp>`;

test("v3 maps compact anchor cells to physical source columns and saves both dialects without touching merge continuations", async () => {
  for (const strict of [false, true]) {
    const f = tableFixture(merged(), strict), input = encodeStructure(f), session = await open(f);
    try {
      assert.deepEqual(session.report.tableTextRuns.map(b => [b.rowIndex, b.cellIndex, b.columnIndex]), [[0, 0, 0], [0, 1, 2], [1, 0, 2]]);
      const run = session.report.tableTextRuns[2];
      assert.ok(Object.isFrozen(run) && Object.isFrozen(run.source));
      assert.equal(run.elementId, session.document.slides[0].elements[0].id);
      assert.match(run.source.path, /a:tr\[2\]\/a:tc\[3\]$/);
      assert.equal("owner" in run, false); assert.equal("baselineText" in run, false);
      edit(session, 0, 0, "Anchor<&🙂\r\n"); edit(session, 1, 0, "é changed");
      const saved = await save(session), expected = f.parts["deck/z.xml"].replace(">Anchor</a:t>", ">Anchor&lt;&amp;🙂&#xD;\n</a:t>").replace(">Bottom right</a:t>", ">é changed</a:t>");
      assert.equal(readGeneratedZip(saved.bytes).get("deck/z.xml").toString(), expected);
      unchanged(input, saved.bytes);
      assert.deepEqual((await compareOpcPackages(input, saved.bytes)).differences.map(d => [d.code, d.part]), [["payload-changed", "/deck/z.xml"]]);
      assert.deepEqual(saved.report.changes.map(c => c.documentPath), ["/slides/0/elements/0/rows/0/cells/0/content/0/runs/0/text", "/slides/0/elements/0/rows/1/cells/0/content/0/runs/0/text"]);
      const reopened = await importPptx(saved.bytes, profile);
      assert.deepEqual(normalizePresentation(reopened.document), normalizePresentation(session.document)); reopened.dispose();
    } finally { session.dispose(); }
  }
});

test("default, v1 and v2 preserve their report shapes and table-edit rejection; v3 requires content mode", async () => {
  for (const editSaveProfile of [undefined, "text-runs-v1", "text-runs-v2"]) {
    const session = await open(tableFixture(), { editSaveProfile });
    assert.equal(session.report.tableTextRuns, undefined);
    edit(session, 0, 0, "outside earlier profile");
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked"); session.dispose();
  }
  await assert.rejects(importPptx(encodeStructure(tableFixture()), { projection: "structure", editSaveProfile: "text-runs-v3" }), TypeError);
  const empty = await open(tableFixture(shape({ txBox: true, text: body() })));
  assert.deepEqual(empty.report.tableTextRuns, []); assert.deepEqual(empty.report.notesTextRuns, []);
  assert.equal(empty.report.textRuns[0].owner, "text-box"); empty.dispose();
});

test("multiple tables and nested rotated/flipped tables keep distinct run identities and source geometry", async () => {
  const f = tableFixture(table() + group(table({ id: "3", transform: tableTransform('rot="2700000" flipV="1"') })));
  const session = await open(f);
  edit(session, 0, 0, "only nested", 3);
  const saved = await save(session);
  const xml = f.parts["deck/z.xml"], at = xml.lastIndexOf(">R1C1</a:t>");
  assert.equal(readGeneratedZip(saved.bytes).get("deck/z.xml").toString(), xml.slice(0, at) + xml.slice(at).replace(">R1C1</a:t>", ">only nested</a:t>"));
  assert.equal(saved.report.changes[0].elementId, session.document.slides[0].elements[1].children[0].id);
  assert.equal(session.report.tableTextRuns.length, 8);
  assert.equal(saved.report.changes[0].documentPath, "/slides/0/elements/1/children/0/rows/0/cells/0/content/0/runs/0/text");
  session.dispose();
});

test("repeated characters inside separate cells edit one raw atom rather than searching source text", async () => {
  const f = tableFixture(table({ rows: [tableRow([cell({ text: "same" }), cell({ text: "same" })], 50)] }));
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("</a:tbl>", '<a:extLst><a:ext uri="urn:future"><x:value xmlns:x="urn:future">same</x:value></a:ext></a:extLst></a:tbl>');
  const session = await open(f); edit(session, 0, 1, "selected");
  const saved = await save(session), xml = f.parts["deck/z.xml"], at = xml.lastIndexOf(">same</a:t>");
  assert.equal(readGeneratedZip(saved.bytes).get("deck/z.xml").toString(), xml.slice(0, at) + xml.slice(at).replace(">same</a:t>", ">selected</a:t>"));
  session.dispose();
});

test("empty source atoms are editable but absent bodies and synthetic empty paragraphs stay unbound", async () => {
  const f = tableFixture(table({ columns: [25, 25, 25, 25], rows: [tableRow([
    cell({ body: tableBody("<a:p><a:r><a:t/></a:r></a:p>") }), cell({ body: "" }),
    cell({ body: tableBody("<a:p/>") }), cell({ text: "control" }),
  ], 50)] }));
  const session = await open(f);
  assert.deepEqual(session.report.tableTextRuns.map(b => b.status), ["bound", "unmapped", "unmapped", "bound"]);
  edit(session, 0, 0, "inserted into atom");
  const saved = await save(session);
  assert.equal(readGeneratedZip(saved.bytes).get("deck/z.xml").toString(), f.parts["deck/z.xml"].replace("<a:t/>", "<a:t>inserted into atom</a:t>"));
  for (const index of [1, 2]) {
    edit(session, 0, index, "no source"); assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    edit(session, 0, index, "");
  }
  await save(session); session.dispose();
});

test("namespace aliases and inactive duplicate tables retain exact selected raw origins", async () => {
  const f = tableFixture(`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x">${table()}</mc:Choice><mc:Fallback>${table()}</mc:Fallback></mc:AlternateContent>`);
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replaceAll("<a:", "<d:").replaceAll("</a:", "</d:").replace('xmlns:a=', 'xmlns:d=');
  const session = await open(f); edit(session, 0, 0, "selected fallback");
  const xml = f.parts["deck/z.xml"], at = xml.lastIndexOf(">R1C1</d:t>");
  assert.equal(readGeneratedZip((await save(session)).bytes).get("deck/z.xml").toString(), xml.slice(0, at) + xml.slice(at).replace(">R1C1</d:t>", ">selected fallback</d:t>"));
  assert.ok(session.report.tableTextRuns.every(b => b.source.elementIndex === b.sourceElementIndex));
  session.dispose();
});

test("known table range/opaque references block the edited owner while whole-shape and other-table references survive", async () => {
  for (const [target, permitted] of [
    ['<p:spTgt spid="2"><p:txEl><p:charRg st="0" end="2"/></p:txEl></p:spTgt>', false],
    ['<p:spTgt spid=" +0002 "><p:graphicEl/></p:spTgt>', false],
    ['<p:spTgt spid="2" future="opaque"/>', false],
    ['<p:spTgt spid="2"/>', true], ['<p:spTgt spid="3"><p:txEl/></p:spTgt>', true],
  ]) {
    const f = tableFixture(table() + table({ id: "3" }));
    f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replace("</p:sld>", `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x"><p:timing>${target}</p:timing></mc:Choice><mc:Fallback/></mc:AlternateContent></p:sld>`);
    const session = await open(f); edit(session, 0, 0, "changed");
    const plan = await session.prepareSave({ mode: "edit" });
    assert.equal(plan.status, permitted ? "ready" : "blocked");
    if (permitted) assert.ok(readGeneratedZip((await session.save(plan.plan)).bytes).get("deck/z.xml").toString().includes(target));
    else assert.equal(plan.report.diagnostics[0].code, "save-text-range-reference");
    session.dispose();
  }
});

test("cell hyperlinks, list/run styles, fields and unknown table/cell extensions retain their source bytes", async () => {
  const content = '<a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="4"/></a:pPr><a:r><a:rPr b="1"><a:hlinkClick r:id="url"/></a:rPr><a:t>linked</a:t></a:r><a:fld id="field" type="slidenum"><a:t>1</a:t></a:fld><a:endParaRPr i="1"/></a:p>';
  const props = '<a:tcPr><a:extLst><a:ext uri="urn:future"><x:value xmlns:x="urn:future">linked</x:value></a:ext></a:extLst></a:tcPr>';
  const f = tableFixture(table({ properties: '<a:tblPr firstRow="1"><a:tableStyleId>{style-identity}</a:tableStyleId></a:tblPr>', rows: [tableRow([cell({ body: tableBody(content), properties: props }), cell()], 50)] }));
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), { ...f.rel("url", "hyperlink", "https://example.invalid/?x=1&y=2"), TargetMode: "External" }]);
  const session = await open(f); edit(session, 0, 0, "link edit");
  const saved = await save(session);
  assert.equal(readGeneratedZip(saved.bytes).get("deck/z.xml").toString(), f.parts["deck/z.xml"].replace(">linked</a:t>", ">link edit</a:t>"));
  unchanged(encodeStructure(f), saved.bytes);
  const reopened = await importPptx(saved.bytes, profile);
  assert.equal(reopened.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0].action.url, "https://example.invalid/?x=1&y=2");
  reopened.dispose(); session.dispose();
});

test("mixed source atoms and break runs remain uneditable and block otherwise supported table changes atomically", async () => {
  for (const content of [
    "<a:p><a:r><a:t>mixed<!--keep-->text</a:t></a:r></a:p>",
    "<a:p><a:br/></a:p>",
  ]) {
    const f = tableFixture(table({ rows: [tableRow([cell({ body: tableBody(content) }), cell({ text: "control" })], 50)] }));
    const session = await open(f); edit(session, 0, 1, "valid"); edit(session, 0, 0, "blocked");
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
    edit(session, 0, 0, content.includes("br/") ? "\n" : "mixedtext"); await save(session); session.dispose();
  }
});

test("styles, grid metrics, spans, row/cell topology and actions cannot hide in a supported table-character plan", async () => {
  const mutations = [
    ["style", e => { e.rows[0].cells[0].content[0].runs[0].style.bold = true; }],
    ["style", e => { e.rows[0].cells[0].style.margin.left = 9; }],
    ["geometry", e => { e.rows[0].height += 1; }], ["geometry", e => { e.columns[0] += 1; }],
    ["geometry", e => { e.box.width += 1; }], ["structure", e => { e.rows[0].cells[0].content[0].runs.push({ text: "new" }); }],
    ["other", e => { e.rows[0].cells[0].content[0].runs[0].action = { type: "url", url: "https://example.invalid/" }; }],
  ];
  for (const [kind, mutate] of mutations) {
    const session = await open(); edit(session, 1, 1, "valid"); modify(session, 2, mutate);
    const blocked = await session.prepareSave({ mode: "edit" });
    assert.equal(blocked.status, "blocked"); assert.ok(blocked.report.diagnostics.some(d => d.code === "save-edit-unsupported"));
    assert.ok(blocked.report.changes.some(c => c.kind === kind), JSON.stringify(blocked.report.changes));
    assert.ok(blocked.report.changes.some(c => c.kind === "text")); session.dispose();
  }
  const session = await open(), original = structuredClone(session.document.slides[0].elements[0]);
  modify(session, 2, e => { e.rows[0].cells = [{ ...e.rows[0].cells[0], colSpan: 2 }]; });
  assert.equal((await session.prepareSave({ mode: "edit" })).status, "blocked");
  modify(session, 2, e => Object.assign(e, original)); edit(session, 0, 0, "restored"); await save(session); session.dispose();
});

test("unsupported sparse/zero/continuation-text tables acquire no bindings while unrelated text remains editable", async () => {
  for (const markup of [
    table({ columns: [0, 100] }), table({ rows: [tableRow([cell()])] }),
    table({ rows: [tableRow([cell({ attributes: 'gridSpan="2"' }), cell({ attributes: 'hMerge="1"', text: "hidden" })], 50)] }),
  ]) {
    const session = await open(tableFixture(markup + shape({ id: "3", txBox: true, text: body() })));
    assert.deepEqual(session.report.tableTextRuns, []);
    modify(session, 3, e => { e.content[0].runs[0].text = "safe text"; }); await save(session); session.dispose();
  }
});

test("exact undo restores original-byte no-op and table changes invalidate prepared plans without advancing baseline", async () => {
  const f = tableFixture(), input = encodeStructure(f), session = await open(f), original = structuredClone(session.document.slides[0].elements[0]);
  const noop = await session.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready");
  edit(session, 0, 0, "edited"); const saved = await save(session);
  assert.equal((await session.prepareSave({ mode: "no-op" })).status, "blocked");
  edit(session, 1, 1, "stale"); await assert.rejects(session.save(saved.plan), code("save-snapshot-stale"));
  modify(session, 2, e => Object.assign(e, original));
  assert.equal(digest((await session.save(noop.plan)).bytes), digest(input)); session.dispose();
});

test("table provenance and complete change reports obey shared limits; failures permit valid retries", async () => {
  const many = tableFixture(table({ columns: [100], rows: [tableRow([cell({ body: tableBody(`<a:p>${"<a:r><a:t>x</a:t></a:r>".repeat(600)}</a:p>`) })], 50)] }));
  await assert.rejects(open(many, { sessionLimits: { maxReportFindings: 500 } }), code("import-limit-exceeded"));
  const repeated = `<a:p>${"<a:r><a:t>x</a:t></a:r>".repeat(260)}</a:p>`;
  const combined = notesFixture(notesShape({ text: body(repeated) }));
  combined.parts["deck/z.xml"] = combined.parts["deck/z.xml"].replace(shape({ txBox: true, text: body() }), table({ columns: [100], rows: [tableRow([cell({ body: tableBody(repeated) })], 50)] }));
  await assert.rejects(open(combined, { sessionLimits: { maxReportFindings: 500 } }), error => {
    assert.equal(error.code, "import-limit-exceeded"); assert.match(error.message, /Text run source bindings/); return true;
  });
  const session = await open(tableFixture(), { sessionLimits: { maxSaveChanges: 1, maxLivePlans: 1 } });
  edit(session, 0, 0, "one"); edit(session, 0, 1, "two");
  await assert.rejects(session.prepareSave({ mode: "edit" }), code("save-limit-exceeded"));
  edit(session, 0, 1, "R1C2"); edit(session, 0, 0, "invalid\u0001");
  const invalid = await session.prepareSave({ mode: "edit" }); assert.equal(invalid.status, "blocked");
  assert.equal(invalid.report.diagnostics[0].code, "save-text-candidate-unsupported");
  edit(session, 0, 0, "valid"); const saved = await save(session); saved.plan.dispose();
  const abort = new AbortController(); abort.abort();
  await assert.rejects(session.prepareSave({ mode: "edit", signal: abort.signal }), code("operation-aborted"));
  await save(session); session.dispose();
});

test("private alias reconciliation includes row/cell identity and blocks conflicting views of one raw atom", async () => {
  const session = await open(), baseline = normalizePresentation(session.document), candidate = structuredClone(baseline), first = session.report.tableTextRuns[0];
  baseline.slides[0].elements[0].rows[0].cells[1].content = structuredClone(baseline.slides[0].elements[0].rows[0].cells[0].content);
  candidate.slides[0].elements[0].rows[0].cells[1].content = structuredClone(baseline.slides[0].elements[0].rows[0].cells[0].content);
  const binding = { ...first, owner: "table-cell", baselineText: "R1C1", namespaceUri: tableFixture().A };
  const bindings = [binding, { ...binding, cellIndex: 1, columnIndex: 1 }];
  candidate.slides[0].elements[0].rows[0].cells[0].content[0].runs[0].text = "agreed";
  const conflict = await analyzeTextEdits(baseline, candidate, bindings, [], DEFAULT_SESSION_LIMITS, async () => {});
  assert.equal(conflict.diagnostics[0].code, "save-source-shared-conflict");
  candidate.slides[0].elements[0].rows[0].cells[1].content[0].runs[0].text = "agreed";
  const agreed = await analyzeTextEdits(baseline, candidate, bindings, [], DEFAULT_SESSION_LIMITS, async () => {});
  assert.deepEqual(agreed.diagnostics, []); assert.equal(agreed.parts.get("/deck/z.xml").length, 1);
  session.dispose();
});

test("UTF-16LE/BE table source edits preserve encoding and every byte outside approved character regions", async () => {
  for (const bigEndian of [false, true]) {
    const f = tableFixture(), parts = readGeneratedZip(encodeStructure(f)), xml = f.parts["deck/z.xml"].replace(/encoding="UTF-8"/i, 'encoding="UTF-16"'), bytes = Buffer.from(xml, "utf16le");
    parts.set("deck/z.xml", Buffer.concat([Buffer.from(bigEndian ? [254, 255] : [255, 254]), bigEndian ? bytes.swap16() : bytes]));
    const input = zip([...parts].map(([name, body]) => ({ name, body }))).bytes;
    const session = await importPptx(input, profile); edit(session, 0, 0, "编码🙂\r<&");
    const saved = await save(session), output = readGeneratedZip(saved.bytes).get("deck/z.xml");
    assert.deepEqual([...output.subarray(0, 2)], bigEndian ? [254, 255] : [255, 254]);
    const decoded = (bigEndian ? Buffer.from(output.subarray(2)).swap16() : output.subarray(2)).toString("utf16le");
    assert.equal(decoded, xml.replace(">R1C1</a:t>", ">编码🙂&#xD;&lt;&amp;</a:t>")); unchanged(input, saved.bytes); session.dispose();
  }
});

test("24 independently partitioned grids save every anchor at its physical column without continuation drift", async () => {
  let seed = 8367;
  const random = max => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % max; };
  for (let sample = 0; sample < 24; sample++) {
    const height = 1 + random(4), width = 1 + random(5), rectangles = [{ x: 0, y: 0, width, height }];
    for (let split = 0; split < 7; split++) {
      const at = random(rectangles.length), rect = rectangles[at];
      if (rect.width > 1 && (rect.height === 1 || random(2))) {
        const n = 1 + random(rect.width - 1); rectangles.splice(at, 1, { ...rect, width: n }, { ...rect, x: rect.x + n, width: rect.width - n });
      } else if (rect.height > 1) {
        const n = 1 + random(rect.height - 1); rectangles.splice(at, 1, { ...rect, height: n }, { ...rect, y: rect.y + n, height: rect.height - n });
      }
    }
    const rows = [], expected = [];
    for (let y = 0; y < height; y++) {
      const cells = []; let coreCell = 0;
      for (let x = 0; x < width; x++) {
        const rect = rectangles.find(r => x >= r.x && x < r.x + r.width && y >= r.y && y < r.y + r.height);
        if (rect.x === x && rect.y === y) {
          cells.push(cell({ text: `cell-${y}-${x}`, attributes: `gridSpan="${rect.width}" rowSpan="${rect.height}"` }));
          expected.push([y, coreCell++, x]);
        } else cells.push(cell({ body: "", attributes: `${x > rect.x ? 'hMerge="1"' : ""} ${y > rect.y ? 'vMerge="1"' : ""}` }));
      }
      rows.push(tableRow(cells, 10));
    }
    const f = tableFixture(table({ columns: Array(width).fill(20), rows, transform: tableTransform("", width * 20, height * 10) })), session = await open(f);
    assert.deepEqual(session.report.tableTextRuns.map(b => [b.rowIndex, b.cellIndex, b.columnIndex]), expected);
    for (const [row, cellIndex, column] of expected) edit(session, row, cellIndex, `edit-${row}-${column}`);
    const saved = await save(session);
    assert.equal(readGeneratedZip(saved.bytes).get("deck/z.xml").toString(), f.parts["deck/z.xml"].replaceAll(/>cell-(\d+)-(\d+)<\/a:t>/g, ">edit-$1-$2</a:t>"));
    session.dispose();
  }
});

test("three save/reimport rounds retain stable merge identities, unknown source content and output size", async () => {
  for (const strict of [false, true]) {
    const f = tableFixture(merged(), strict);
    f.parts["orphan.xml"] = '<future xmlns="urn:future">keep</future>';
    let input = encodeStructure(f), previousSize, identities;
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(input, profile), current = session.report.tableTextRuns.map(b => [b.elementId, b.rowIndex, b.cellIndex, b.columnIndex, b.sourceElementIndex]);
      if (identities) assert.deepEqual(current, identities); identities = current;
      edit(session, 1, 0, `Round ${round}`); const saved = await save(session);
      unchanged(input, saved.bytes);
      if (previousSize) assert.equal(saved.bytes.length, previousSize);
      previousSize = saved.bytes.length; input = saved.bytes; session.dispose();
    }
  }
});
