import assert from "node:assert/strict";
import test from "node:test";
import { normalizePresentation, validatePresentation } from "@pptkit/core";
import { importPptx, PptxImportError } from "../dist/index.js";
import { cell, table, tableBody, tableBytes, tableFixture, tableRow, tableTransform } from "./table-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { shape } from "./content-fixture.mjs";

const open = (input = tableBytes(), options = {}) => importPptx(input, { projection: "content", ...options });
const elements = s => normalizePresentation(s.document).slides[0].elements;
const plain = cell => cell.content.map(p => p.runs.map(r => r.text).join("")).join("\n");
const findings = s => s.report.findings.filter(f => f.operation === "import");
const unhandled = s => findings(s).filter(f => f.feature === "drawing.unmapped-property");
const errorCode = code => e => e instanceof PptxImportError && e.code === code;
const mergedRows = () => [
  tableRow([cell({ text: "Anchor", attributes: 'gridSpan="2" rowSpan="2"' }), cell({ attributes: 'hMerge="1"' }), cell({ text: "Top right" })], 20),
  tableRow([cell({ attributes: 'vMerge="1"' }), cell({ attributes: 'hMerge="1" vMerge="1"' }), cell({ text: "Bottom right" })], 30),
];

test("tables import exact physical grids and stable graphic-frame identities in both dialects and table URIs", async () => {
  for (const strict of [false, true]) for (const uri of ["http://schemas.openxmlformats.org/drawingml/2006/table", "http://purl.oclc.org/ooxml/drawingml/table"]) {
    const s = await open(tableBytes(table({ uri }), strict)), [t] = elements(s);
    assert.equal(t.type, "table"); assert.equal(t.id, "source-slide-300-object-2");
    assert.deepEqual(t.box, { x: 10, y: 20, width: 100, height: 50 });
    assert.deepEqual(t.columns, [40, 60]); assert.deepEqual(t.rows.map(r => r.height), [20, 30]);
    assert.deepEqual(t.rows.map(r => r.cells.map(c => plain(c))), [["R1C1", "R1C2"], ["R2C1", "R2C2"]]);
    assert.equal(t.name, "Table 2"); assert.equal(t.accessibility.description, "Table description");
    assert.deepEqual(validatePresentation(s.document), []);
    assert.match(s.report.identities.find(i => i.kind === "element").source.path, /p:graphicFrame\[1\]$/);
    const mapped = s.report.findings.filter(f => f.feature === "drawing.table.mapped-fields");
    assert.equal(mapped.length, 7);
    assert.equal(mapped.find(f => f.operation === "edit-save").tier, "unsupported");
    assert.equal(mapped.find(f => f.operation === "no-op-save").tier, "preserved");
    assert.equal(mapped.find(f => f.operation === "preview").tier, "editable");
    s.dispose();
  }
});

test("table rectangle merges become compact anchor rows and fully merged rows may be empty", async () => {
  const s = await open(tableBytes(table({ columns: [30, 30, 40], rows: mergedRows() }))), [t] = elements(s);
  assert.deepEqual(t.rows.map(r => r.cells.map(c => [plain(c), c.rowSpan, c.colSpan])), [
    [["Anchor", 2, 2], ["Top right", 1, 1]], [["Bottom right", 1, 1]],
  ]);
  assert.deepEqual(validatePresentation(s.document), []); s.dispose();
  const all = await open(tableBytes(table({ rows: [
    tableRow([cell({ text: "All", attributes: 'gridSpan="2" rowSpan="2"' }), cell({ attributes: 'hMerge="true" rowSpan="2"' })]),
    tableRow([cell({ attributes: 'vMerge="true" gridSpan="2"' }), cell({ attributes: 'vMerge="1" hMerge="1"' })], 30),
  ] })));
  assert.equal(elements(all)[0].rows[1].cells.length, 0);
  assert.equal(elements(all)[0].rows[0].cells[0].rowSpan, 2);
  assert.deepEqual(validatePresentation(all.document), []); all.dispose();
});

test("expired vertical occupancy yields to later anchors without shifting following cells", async () => {
  const rows = [
    tableRow([cell({ text: "A", attributes: 'rowSpan="2"' }), cell({ text: "B", attributes: 'rowSpan="3"' })], 10),
    tableRow([cell({ attributes: 'vMerge="1"' }), cell({ attributes: 'vMerge="1"' })], 10),
    tableRow([cell({ text: "C" }), cell({ attributes: 'vMerge="1"' })], 10),
  ];
  const s = await open(tableBytes(table({ rows, transform: tableTransform("", 100, 30) })));
  assert.deepEqual(elements(s)[0].rows.map(r => r.cells.map(c => plain(c))), [["A", "B"], [], ["C"]]);
  s.dispose();
});

test("independently partitioned rectangular grids import every anchor exactly once", async () => {
  let seed = 72381;
  const random = max => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % max; };
  for (let sample = 0; sample < 40; sample++) {
    const height = 1 + random(5), width = 1 + random(5);
    // Generate a disjoint partition by geometric bisection, independently of
    // the parser's streaming occupancy/continuation algorithm.
    const rectangles = [{ y: 0, x: 0, height, width }];
    for (let split = 0; split < 8; split++) {
      const at = random(rectangles.length), rect = rectangles[at];
      if (rect.width > 1 && (rect.height === 1 || random(2))) {
        const first = 1 + random(rect.width - 1);
        rectangles.splice(at, 1, { ...rect, width: first }, { ...rect, x: rect.x + first, width: rect.width - first });
      } else if (rect.height > 1) {
        const first = 1 + random(rect.height - 1);
        rectangles.splice(at, 1, { ...rect, height: first }, { ...rect, y: rect.y + first, height: rect.height - first });
      }
    }
    const physical = [], expected = [];
    for (let y = 0; y < height; y++) {
      const cells = [], anchors = [];
      for (let x = 0; x < width; x++) {
        const rect = rectangles.find(r => x >= r.x && x < r.x + r.width && y >= r.y && y < r.y + r.height);
        assert.ok(rect);
        if (rect.x === x && rect.y === y) {
          const text = `${y}:${x}`;
          cells.push(cell({ text, attributes: `rowSpan="${rect.height}" gridSpan="${rect.width}"` }));
          anchors.push([text, rect.height, rect.width]);
        } else cells.push(cell({ body: "", attributes: `${x > rect.x ? 'hMerge="1"' : ""} ${y > rect.y ? 'vMerge="1"' : ""}` }));
      }
      physical.push(tableRow(cells, 10)); expected.push(anchors);
    }
    const s = await open(tableBytes(table({ columns: Array(width).fill(20), rows: physical, transform: tableTransform("", width * 20, height * 10) })));
    assert.deepEqual(elements(s)[0].rows.map(r => r.cells.map(c => [plain(c), c.rowSpan, c.colSpan])), expected, `partition ${sample}`);
    assert.equal(elements(s)[0].rows.reduce((n, r) => n + r.cells.length, 0), rectangles.length);
    s.dispose();
  }
});

test("graphic-frame rotation/flips and nonzero parent origins are projected once", async () => {
  const child = table({ transform: tableTransform('rot="2700000" flipH="1" flipV="true"') });
  const group = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Parent"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000"><a:off x="1270000" y="1270000"/><a:ext cx="2540000" cy="1270000"/><a:chOff x="63500" y="88900"/><a:chExt cx="1270000" cy="635000"/></a:xfrm></p:grpSpPr>${child}</p:grpSp>`;
  const s = await open(tableBytes(group)), [g] = elements(s), t = g.children[0];
  assert.deepEqual(t.box, { x: 5, y: 13, width: 100, height: 50 });
  assert.deepEqual(t.transform, { rotation: 45, flipH: true, flipV: true });
  assert.deepEqual(g.coordinateSize, { width: 100, height: 50 });
  assert.equal(g.transform.rotation, 90); s.dispose();
});

test("orphan, inconsistent, overlapping and out-of-grid merge structures fail with a typed package error", async () => {
  const cases = [
    [cell({ attributes: 'hMerge="1"' }), cell()],
    [cell({ attributes: 'vMerge="1"' }), cell()],
    [cell({ attributes: 'gridSpan="3"' }), cell()],
    [cell({ attributes: 'rowSpan="3"' }), cell()],
    [cell({ attributes: 'gridSpan="2"' }), cell()], // Missing continuation flag.
    [cell({ attributes: 'gridSpan="2"' }), cell({ attributes: 'hMerge="1" gridSpan="2"' })],
    [cell({ attributes: 'gridSpan="2"' }), cell({ attributes: 'hMerge="1" vMerge="1"' })],
  ];
  for (const cells of cases) await assert.rejects(open(tableBytes(table({ rows: [tableRow(cells), tableRow([cell(), cell()], 30)] }))), errorCode("import-invalid-package"));
  // Row two tries to cross an active vertical merge that starts in column two.
  const rows = [tableRow([cell(), cell({ attributes: 'rowSpan="2"' })]), tableRow([cell({ attributes: 'gridSpan="2"' }), cell({ attributes: 'vMerge="1"' })], 30)];
  await assert.rejects(open(tableBytes(table({ rows }))), errorCode("import-invalid-package"));
});

test("table numeric and boolean lexical errors reject without manufacturing dimensions", async () => {
  const original = table();
  const cases = [
    original.replace('w="508000"', 'w="NaN"'),
    original.replace('w="508000"', 'w="1.5"'),
    original.replace('w="508000"', 'w="27273042316901"'),
    original.replace('w="508000"', ""),
    original.replace('h="254000"', 'h="1e5"'),
    original.replace('h="254000"', 'h="-27273042329601"'),
    original.replace("<a:tc ", '<a:tc gridSpan="0" '),
    original.replace("<a:tc ", '<a:tc rowSpan="-1" '),
    original.replace("<a:tc ", '<a:tc gridSpan="2147483648" '),
    original.replace("<a:tc ", '<a:tc hMerge="yes" '),
  ];
  for (const markup of cases) await assert.rejects(open(tableBytes(markup)), errorCode("import-invalid-package"));
});

test("nonpositive, sparse and excess grids remain source-only with exact original-byte no-op", async () => {
  const cases = [
    table({ columns: [0, 100] }), table({ columns: [-1, 101] }),
    table({ rows: [tableRow([cell(), cell()], 0)] }),
    table({ rows: [tableRow([cell(), cell()], -1)] }),
    table({ rows: [tableRow([cell()])] }),
    table({ rows: [tableRow([cell(), cell(), cell()])] }),
    table({ transform: tableTransform("", 0, 50) }),
  ];
  for (const markup of cases) {
    const input = tableBytes(markup), s = await open(input);
    assert.deepEqual(elements(s), []);
    assert.ok(findings(s).some(f => f.feature === "drawing.graphicFrame" && f.tier === "unsupported"));
    const prepared = await s.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
    assert.equal(Buffer.compare(Buffer.from((await s.save(prepared.plan)).bytes), input), 0); s.dispose();
  }
});

test("missing required grids/rows and duplicate singleton structures reject", async () => {
  for (const markup of [
    table({ columns: [] }), table({ rows: [] }),
    table().replace("</a:tbl>", "<a:tblGrid/></a:tbl>"),
    table().replace("<a:tcPr/>", "<a:tcPr/><a:tcPr/>"),
    table().replace("<a:graphic>", "<a:graphic/><a:graphic>"),
  ]) await assert.rejects(open(tableBytes(markup)), errorCode("import-invalid-package"));
});

test("table continuation text, whitespace, fields and breaks prevent lossy anchor-only import", async () => {
  for (const body of [
    tableBody("<a:p><a:r><a:t> </a:t></a:r></a:p>"),
    tableBody("<a:p><a:r><a:t>Hidden content</a:t></a:r></a:p>"),
    tableBody('<a:p><a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>1</a:t></a:fld></a:p>'),
    tableBody("<a:p><a:br/></a:p>"),
  ]) {
    const markup = table({ rows: [tableRow([cell({ attributes: 'gridSpan="2"', text: "Keep" }), cell({ attributes: 'hMerge="1"', body })], 50)] });
    const s = await open(tableBytes(markup));
    assert.deepEqual(elements(s), []);
    assert.ok(findings(s).some(f => f.feature === "drawing.graphicFrame" && /continuation contains/.test(f.reason)));
    assert.ok(!s.report.findings.some(f => f.feature === "drawing.click-action"));
    s.dispose();
  }
});

test("formatted or actionable empty continuation text remains source-only without changing no-op bytes", async () => {
  const bodies = [
    tableBody('<a:p><a:r><a:rPr b="1"/><a:t/></a:r></a:p>'),
    tableBody('<a:p><a:r><a:rPr><a:hlinkClick r:id="missing"/></a:rPr><a:t/></a:r></a:p>'),
    tableBody('<a:p><a:pPr marL="12700"/></a:p>'),
    tableBody('<a:p><a:endParaRPr sz="1200"/></a:p>'),
    tableBody('<a:p><a:endParaRPr lang="fr-FR"/></a:p>'),
    tableBody('<a:p/>', 'anchor="ctr"'),
    tableBody('<a:p/><a:p/>'),
    '<a:txBody><a:bodyPr/><a:lstStyle><a:lvl1pPr marL="100"/></a:lstStyle><a:p/></a:txBody>',
    tableBody('<a:p><a:extLst><a:ext uri="urn:unknown"/></a:extLst></a:p>'),
    tableBody('<a:p><x:future xmlns:x="urn:unknown" value="1"/></a:p>'),
    tableBody('<a:p unknown="value"/>'),
    tableBody('<a:p><a:r><a:t><a:rPr/></a:t></a:r></a:p>'),
    tableBody('<a:p>unexpected text</a:p>'),
  ];
  for (const strict of [false, true]) for (const body of bodies) {
    const markup = table({ rows: [tableRow([cell({ attributes: 'gridSpan="2"', text: "Keep" }), cell({ attributes: 'hMerge="1"', body })], 50)] });
    const bytes = tableBytes(markup, strict), session = await open(bytes);
    try {
      assert.deepEqual(elements(session), []);
      assert.ok(findings(session).some(f => f.feature === "drawing.graphicFrame" && /continuation contains.*formatting\/actions/.test(f.reason)));
      const plan = await session.prepareSave({ mode: "no-op" });
      assert.equal(plan.status, "ready");
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
  for (const body of [tableBody('<a:p><a:endParaRPr lang="en-US"/></a:p>'), tableBody('<a:p><a:r><a:t xml:space="preserve"/></a:r></a:p>')]) {
    const safe = table({ rows: [tableRow([cell({ attributes: 'gridSpan="2"', text: "Keep" }), cell({ attributes: 'hMerge="1"', body })], 50)] });
    const session = await open(tableBytes(safe));
    try { assert.equal(elements(session)[0].rows[0].cells[0].colSpan, 2); } finally { session.dispose(); }
  }
});

test("empty optional cell bodies and styled empty paragraphs import without synthesized visible text", async () => {
  const s = await open(tableBytes(table({ rows: [tableRow([
    cell({ body: "", properties: '<a:tcPr marL="0"><a:solidFill><a:srgbClr val="FFAA00"/></a:solidFill></a:tcPr>' }),
    cell({ body: tableBody('<a:p><a:endParaRPr sz="1200" b="1"/></a:p>') }),
  ], 50)] }))), [t] = elements(s);
  assert.equal(plain(t.rows[0].cells[0]), ""); assert.equal(t.rows[0].cells[0].style.fill.color, "FFAA00");
  assert.equal(plain(t.rows[0].cells[1]), "");
  assert.equal(t.rows[0].cells[1].content[0].runs[0].style.fontSize, 12);
  assert.equal(t.rows[0].cells[1].content[0].runs[0].style.bold, true); s.dispose();
});

test("cell rich text retains lists, breaks, Unicode, run actions and source locations", async () => {
  const paragraphs = `<a:p><a:pPr algn="r"><a:buAutoNum type="arabicPeriod" startAt="4"/></a:pPr><a:r><a:rPr sz="1800" b="1"><a:hlinkClick r:id="link"/></a:rPr><a:t>α😀e\u0301 &amp; &lt;literal&gt;</a:t></a:r><a:br/><a:r><a:rPr i="1"/><a:t>Next</a:t></a:r></a:p><a:p/>`;
  const f = tableFixture(table({ rows: [tableRow([cell({ body: tableBody(paragraphs) }), cell()], 50)] }));
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), { ...f.rel("link", "hyperlink", "https://example.invalid/?a=1&b=2"), TargetMode: "External" }]);
  const s = await open(encodeStructure(f)), first = elements(s)[0].rows[0].cells[0];
  assert.equal(plain(first), "α😀e\u0301 & <literal>\nNext\n");
  assert.equal(first.content[0].style.align, "right");
  assert.equal(first.content[0].style.bullet.startAt, 4);
  assert.equal(first.content[0].runs[0].style.fontSize, 18);
  assert.equal(first.content[0].runs[0].style.bold, true);
  assert.equal(first.content[0].runs[2].style.italic, true);
  assert.equal(first.content[0].runs[0].action.url, "https://example.invalid/?a=1&b=2");
  assert.ok(findings(s).some(f => f.feature === "drawing.click-action" && /a:tr\[1\].*a:tc\[1\].*a:hlinkClick/.test(f.source.path) && f.elementId === "source-slide-300-object-2"));
  s.dispose();
});

test("direct fill, opacity and uniform four-edge borders map including sysDot", async () => {
  const borders = ["lnL", "lnR", "lnT", "lnB"].map(n => `<a:${n} w="25400"><a:solidFill><a:srgbClr val="123456"><a:alpha val="75000"/></a:srgbClr></a:solidFill><a:prstDash val="sysDot"/></a:${n}>`).join("");
  const props = `<a:tcPr>${borders}<a:solidFill><a:srgbClr val="ABCDEF"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr>`;
  const s = await open(tableBytes(table({ rows: [tableRow([cell({ properties: props }), cell()], 50)] })));
  const style = elements(s)[0].rows[0].cells[0].style;
  assert.deepEqual(style.fill, { type: "solid", color: "ABCDEF", opacity: .5 });
  assert.equal(style.stroke.width, 2); assert.equal(style.stroke.dash, "dot");
  assert.deepEqual(style.stroke.paint, { type: "solid", color: "123456", opacity: .75 });
  assert.ok(!unhandled(s).some(f => /Per-edge/.test(f.reason))); s.dispose();
});

test("incomplete and unequal direct borders map individual edges without fabricating a uniform stroke", async () => {
  for (const lines of [
    '<a:lnL w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL>',
    ["lnL", "lnR", "lnT", "lnB"].map((n, i) => `<a:${n} w="${(i + 1) * 12700}"><a:noFill/></a:${n}>`).join(""),
  ]) {
    const s = await open(tableBytes(table({ rows: [tableRow([cell({ properties: `<a:tcPr>${lines}</a:tcPr>` }), cell()], 50)] })));
    assert.equal(elements(s)[0].rows[0].cells[0].style.stroke.paint.type, "none");
    assert.ok(elements(s)[0].rows[0].cells[0].style.borders.left);
    assert.ok(!unhandled(s).some(f => /Per-edge\/unequal/.test(f.reason)));
    s.dispose();
  }
});

test("standard cell margins/anchor and supported direction override legacy body insets while unsupported extras remain source-only", async () => {
  const first = cell({ body: tableBody("<a:p/>", 'lIns="12700" rIns="25400" tIns="38100" bIns="50800" anchor="ctr" wrap="none"'), properties: '<a:tcPr marL="63500" marR="76200" marT="0" marB="101600" anchor="b" vert="vert"/>' });
  const second = cell({ body: tableBody("<a:p/>", 'lIns="12700" anchor="ctr"'), properties: '<a:tcPr marL="-1"/>' });
  const s = await open(tableBytes(table({ rows: [tableRow([first, second], 50)] }))), [t] = elements(s);
  assert.deepEqual(t.rows[0].cells[0].style.margin, { left: 5, right: 6, top: 0, bottom: 8 });
  assert.equal(t.rows[0].cells[0].style.verticalAlign, "bottom");
  assert.equal(t.rows[0].cells[0].style.textDirection, "rotate90");
  assert.equal(t.rows[0].cells[1].style.verticalAlign, "middle");
  assert.ok(unhandled(s).some(f => f.source.path.endsWith("/@wrap")));
  assert.ok(!unhandled(s).some(f => f.source.path.endsWith("/@vert")));
  assert.ok(unhandled(s).some(f => /Negative table/.test(f.reason))); s.dispose();
});

test("table style IDs, theme fill, diagonal borders and opaque extensions remain located source properties", async () => {
  const props = '<a:tcPr><a:lnTlToBr/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:extLst><a:ext uri="urn:future"><x:value xmlns:x="urn:future">Secret</x:value></a:ext></a:extLst></a:tcPr>';
  const markup = table({ properties: '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{00000000-0000-0000-0000-000000000001}</a:tableStyleId></a:tblPr>', rows: [tableRow([cell({ properties: props }), cell()], 50)] });
  const s = await open(tableBytes(markup)), paths = unhandled(s).map(f => f.source.path);
  assert.equal(s.report.tableStyleInheritance.tables[0].flags.firstRow, true);
  assert.equal(s.report.tableStyleInheritance.tables[0].flags.bandRow, true);
  assert.equal(s.report.tableStyleInheritance.tables[0].status, "missing");
  for (const name of ["a:tableStyleId", "a:lnTlToBr", "a:solidFill", "a:extLst"]) assert.ok(paths.some(p => p.includes(name)), name);
  assert.ok(unhandled(s).every(f => f.elementId === "source-slide-300-object-2"));
  assert.equal(elements(s)[0].rows[0].cells[0].style.fill.type, "none"); s.dispose();
});

test("authored grid/frame disagreement retains both dimensions and marks mapped preview fallback", async () => {
  const s = await open(tableBytes(table({ transform: tableTransform("", 120, 80) }))), [t] = elements(s);
  assert.deepEqual(t.columns, [40, 60]); assert.deepEqual(t.rows.map(r => r.height), [20, 30]);
  assert.equal(t.box.width, 120); assert.equal(t.box.height, 80);
  assert.ok(s.report.findings.some(f => f.feature === "drawing.table.mapped-fields" && f.operation === "preview" && f.tier === "fallback"));
  assert.ok(unhandled(s).some(f => /Frame extent differs/.test(f.reason))); s.dispose();
});

test("unknown graphic frames stay source-only while following supported drawings still import", async () => {
  const s = await open(tableBytes(table({ uri: "urn:future:chart" }) + shape({ id: "3" })));
  assert.equal(elements(s).length, 1); assert.equal(elements(s)[0].type, "shape");
  assert.ok(findings(s).some(f => f.feature === "drawing.graphicFrame" && /not the supported/.test(f.reason))); s.dispose();
  await assert.rejects(open(tableBytes(table() + shape())), errorCode("import-invalid-package"));
});

test("table namespace aliases and AlternateContent use expanded names and selected branches", async () => {
  const f = tableFixture(`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:q="urn:future"><mc:Choice Requires="q">${shape({ id: "9" })}</mc:Choice><mc:Fallback>${table()}</mc:Fallback></mc:AlternateContent>`);
  f.parts["deck/z.xml"] = f.parts["deck/z.xml"].replaceAll("<a:", "<draw:").replaceAll("</a:", "</draw:").replace('xmlns:a=', 'xmlns:draw=').replaceAll("<p:", "<pres:").replaceAll("</p:", "</pres:").replace('xmlns:p=', 'xmlns:pres=');
  const s = await open(encodeStructure(f));
  assert.equal(elements(s)[0].type, "table"); assert.deepEqual(elements(s)[0].columns, [40, 60]); s.dispose();
});

test("table content participates in budget/cancellation and isolated structural sessions", async () => {
  const input = tableBytes();
  await assert.rejects(open(input, { sessionLimits: { maxReportFindings: 1 } }), errorCode("import-limit-exceeded"));
  const controller = new AbortController(); controller.abort();
  await assert.rejects(open(input, { signal: controller.signal }), errorCode("operation-aborted"));
  const inFlight = new AbortController(), pending = open(input, { signal: inFlight.signal }); inFlight.abort();
  await assert.rejects(pending, errorCode("operation-aborted"));
  const [a, b, structure] = await Promise.all([open(input), open(input), importPptx(input, { projection: "structure" })]);
  assert.deepEqual(elements(structure), []); assert.deepEqual(elements(a), elements(b));
  a.dispose(); assert.equal(plain(elements(b)[0].rows[0].cells[0]), "R1C1");
  b.dispose(); structure.dispose();
});

test("editing a table invalidates no-op plans, undo restores baseline, and edit/rebuild remain blocked", async () => {
  const input = tableBytes(), s = await open(input), slide = s.document.slides[0], original = slide.elements[0];
  const prepared = await s.prepareSave({ mode: "no-op" }); assert.equal(prepared.status, "ready");
  slide.removeElement(original.id); slide.addElement({ ...original, columns: [50, 50] });
  await assert.rejects(s.save(prepared.plan), errorCode("save-snapshot-stale"));
  assert.equal((await s.prepareSave({ mode: "no-op" })).status, "blocked");
  assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
  assert.equal((await s.prepareSave({ mode: "rebuild" })).status, "blocked");
  slide.removeElement(original.id); slide.addElement(original);
  assert.equal(Buffer.compare(Buffer.from((await s.save(prepared.plan)).bytes), input), 0);
  s.dispose(); await assert.rejects(s.save(prepared.plan), errorCode("session-disposed"));
});
