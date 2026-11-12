import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, editTable } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg } from "../packages/svg-renderer/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { tableEditDeck, tableEditBatches, tableEditSource } from "../tests/fixtures/table-edits.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
const profile = { projection: "content", editSaveProfile: "text-runs-v3" };

test("guarded deletion has matching Node, browser ESM and global outcomes",async t=>{
  const source={type:'table',box:{x:0,y:0,width:20,height:20},columns:[10,10],rows:[{height:10,cells:[{content:'A'},{content:''}]},{height:10,cells:[{content:''},{content:''}]}]};
  const run=(api,source)=>{let blocked,blockedData;try{api.editTable(source,[{type:'delete-row',index:0,content:'reject'}],{frame:'fit-grid'});}catch(e){blocked={code:e.code,path:e.path};}const styled=structuredClone(source);styled.rows[1].cells[0].style={margin:3};try{api.editTable(styled,[{type:'delete-row',index:1,content:'reject-data'}],{frame:'fit-grid'});}catch(e){blockedData={code:e.code,path:e.path};}return {blocked,blockedData,safe:api.editTable(source,[{type:'delete-row',index:1,content:'reject'}],{frame:'fit-grid'}),discard:api.editTable(source,[{type:'delete-row',index:0,content:'discard'}],{frame:'fit-grid'})}};
  const expected=run({editTable},source),browser=await chromium.launch();assert.equal(expected.blockedData.code,'table-edit-content-conflict');t.after(()=>browser.close());
  const esm=`data:text/javascript;base64,${Buffer.from(await readFile('packages/core/dist/browser/index.js')).toString('base64')}`,global=await readFile('packages/core/dist/browser/global.js','utf8');
  for(const format of ['esm','global']){const page=await browser.newPage();if(format==='global')await page.addScriptTag({content:global});const actual=await page.evaluate(async({esm,source,fn})=>{const api=esm?await import(esm):PPTKit;return (0,eval)(`(${fn})`)(api,source)}, {esm:format==='esm'?esm:null,source,fn:run.toString()});assert.deepEqual(actual,expected);await page.close();}
});
const plain = cell => cell.content.map(p => p.runs.map(r => r.text).join("")).join("\n");
function model(table) {
  return { columns: table.columns, rows: table.rows.map(row => ({ height: row.height, cells: row.cells.map(cell => ({ rowSpan: cell.rowSpan, colSpan: cell.colSpan, text: plain(cell) })) })) };
}
async function rounds() {
  const { document, slide } = tableEditDeck(createPresentation), snapshots = [];
  for (const operations of tableEditBatches) {
    const before = slide.elements[0], result = editTable(before, operations, { frame: "fit-grid" });
    assert.equal(slide.elements[0], before); slide.replaceElement(before.id, result.table);
    const expected = model(normalizePresentation(document).slides[0].elements[0]), generated = await generatePptx(document), svg = await renderPresentationToSvg(document);
    snapshots.push({ expected, result, generated, svg });
  }
  return { document, slide, snapshots };
}

test("three structural edit batches retain exact grid/content through fresh PPTX import and static SVG without duplicating untouched controls", async () => {
  const { snapshots } = await rounds();
  for (const { expected, generated, svg } of snapshots) {
    assert.deepEqual(generated.warnings, []); assert.equal(generated.slideCount, 2); assert.equal(svg.slides.length, 2);
    const session = await importPptx(generated.bytes, profile);
    try {
      assert.deepEqual(model(normalizePresentation(session.document).slides[0].elements[0]), expected);
      const noop = await session.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready"); assert.deepEqual((await session.save(noop.plan)).bytes, generated.bytes);
    } finally { session.dispose(); }
    assert.ok(svg.slides[1].svg.includes("Unchanged table control"));
  }
  assert.ok(snapshots[1].result.cellMapping.some(m => m.status === "deleted"));
  assert.ok(snapshots[2].result.cellMapping.every(m => m.status === "retained"));
});

test("Node/ESM/global transactions agree and independent DOM verifies every physical PPTX merge slot and HTML anchor after all three edits", async t => {
  const { snapshots } = await rounds(), browser = await chromium.launch(); t.after(() => browser.close());
  let lower = 1, upper = 10000000;
  while (lower < upper) {
    const middle = Math.floor((lower + upper) / 2);
    try { editTable(tableEditSource(), tableEditBatches[0], { frame: "fit-grid", limits: { maxWorkUnits: middle } }); upper = middle; }
    catch (error) { assert.equal(error.code, "table-edit-limit-exceeded"); lower = middle + 1; }
  }
  const esm = `data:text/javascript;base64,${Buffer.from(await readFile("packages/core/dist/browser/index.js", "utf8")).toString("base64")}`, global = await readFile("packages/core/dist/browser/global.js", "utf8");
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    if (format === "global") await page.addScriptTag({ content: global });
    const result = await page.evaluate(async ({ esm, source, batches, exactWork }) => {
      const api = esm ? await import(esm) : PPTKit, outputs = []; let table = source;
      const exact = api.editTable(source, batches[0], { frame: "fit-grid", limits: { maxWorkUnits: exactWork } }); let quota;
      try { api.editTable(source, batches[0], { frame: "fit-grid", limits: { maxWorkUnits: exactWork - 1 } }); } catch (e) { quota = e.code; }
      for (const operations of batches) {
        const result = api.editTable(table, operations, { frame: "fit-grid" }); outputs.push(result); table = result.table;
        if (!Object.isFrozen(result) || !Object.isFrozen(table.rows[0].cells[0].content)) throw new Error("Mutable result");
      }
      const unchanged = JSON.stringify(table); let blocked;
      try { api.editTable(table, [{ type: "delete-column", index: 999 }], { frame: "fit-grid" }); } catch (e) { blocked = e.code; }
      return { outputs, blocked, unchanged: unchanged === JSON.stringify(table), exact, quota };
    }, { esm: format === "esm" ? esm : null, source: tableEditSource(), batches: tableEditBatches, exactWork: lower });
    assert.deepEqual(result.outputs, snapshots.map(s => s.result)); assert.equal(result.blocked, "invalid-table-operation"); assert.equal(result.unchanged, true);
    assert.deepEqual(result.exact, snapshots[0].result); assert.equal(result.quota, "table-edit-limit-exceeded");
    for (const { generated, svg, expected } of snapshots) {
      const xml = readGeneratedZip(generated.bytes).get("ppt/slides/slide1.xml").toString();
      const verified = await page.evaluate(({ xml, svg, expected }) => {
        const parser = new DOMParser(), dom = parser.parseFromString(xml, "application/xml"), a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const table = dom.getElementsByTagNameNS(a, "tbl")[0], rows = [...table.getElementsByTagNameNS(a, "tr")], owner = Array.from({ length: rows.length }, () => Array(expected.columns.length).fill(null));
        const actual = { columns: [...table.getElementsByTagNameNS(a, "gridCol")].map(n => Number(n.getAttribute("w")) / 12700), rows: [] };
        for (let r = 0; r < rows.length; r++) {
          const physical = [...rows[r].children].filter(n => n.localName === "tc"); if (physical.length !== expected.columns.length) return false;
          const cells = [];
          for (let c = 0; c < physical.length; c++) {
            const n = physical[c], existing = owner[r][c];
            if (existing) {
              if ((n.getAttribute("vMerge") === "1") !== (r > existing.row) || (n.getAttribute("hMerge") === "1") !== (c > existing.col)) return false;
              if ([...n.getElementsByTagNameNS(a, "t")].some(n => n.textContent)) return false;
              continue;
            }
            if (n.hasAttribute("hMerge") || n.hasAttribute("vMerge")) return false;
            const rs = Number(n.getAttribute("rowSpan") ?? 1), cs = Number(n.getAttribute("gridSpan") ?? 1), anchor = { row: r, col: c };
            for (let y = r; y < r + rs; y++) for (let x = c; x < c + cs; x++) { if (!owner[y] || x >= expected.columns.length || owner[y][x]) return false; owner[y][x] = anchor; }
            const body = n.getElementsByTagNameNS(a, "txBody")[0], text = [...body.children].filter(n => n.localName === "p").map(p => [...p.getElementsByTagNameNS(a, "t")].map(n => n.textContent).join("")).join("\n");
            cells.push({ rowSpan: rs, colSpan: cs, text });
          }
          actual.rows.push({ height: Number(rows[r].getAttribute("h")) / 12700, cells });
        }
        if (JSON.stringify(actual) !== JSON.stringify(expected)) return false;
        const html = parser.parseFromString(svg, "image/svg+xml"), htmlRows = [...html.getElementsByTagNameNS("http://www.w3.org/1999/xhtml", "tr")];
        if (htmlRows.length !== expected.rows.length) return false;
        return htmlRows.every((row, r) => [...row.children].length === expected.rows[r].cells.length && [...row.children].every((n, c) => Number(n.getAttribute("rowspan") ?? 1) === expected.rows[r].cells[c].rowSpan && Number(n.getAttribute("colspan") ?? 1) === expected.rows[r].cells[c].colSpan));
      }, { xml, svg: svg.slides[0].svg, expected });
      assert.equal(verified, true);
    }
    assert.deepEqual(requests, []); await page.close();
  }
});

test("edited tables remeasure and paginate with indivisible merge blocks while authored state stays unchanged", async () => {
  const { document, slide } = await rounds(), before = normalizePresentation(document), table = slide.elements[0];
  const options = {
    textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 10, height: 60 } }) },
    tableSizing: [{ elementId: table.id, autoRows: table.rows.map((_, row) => ({ row })) }],
    tablePagination: [{ elementId: table.id, area: { x: 10, y: 10, width: 230, height: 130 } }],
  };
  const layout = resolveLayout(document, options);
  assert.equal(layout.slideCount, 3); assert.deepEqual(layout.pagination.fragments.map(f => f.sourceRowIndices), [[0, 1], [2, 3]]);
  assert.deepEqual(layout.slides.slice(0, 2).map(s => s.elements[0].rows.map(r => r.height)), [[60, 60], [60, 60]]);
  assert.deepEqual(normalizePresentation(document), before);
  const generated = await generatePptx(document, { layout: options }), parsed = await importPptx(generated.bytes, { projection: "content" });
  try { assert.equal(parsed.document.slides.length, 3); assert.equal(parsed.document.slides[0].elements[0].rows[0].cells[0].rowSpan, 2); }
  finally { parsed.dispose(); }
});

test("Core structural mappings never grant Parser source-write authority and a blocked source edit can restore no-op and retry an allowed scalar edit", async () => {
  const { document } = tableEditDeck(createPresentation), source = (await generatePptx(document)).bytes, session = await importPptx(source, profile);
  try {
    const slide = session.document.slides[0], original = slide.elements[0], result = editTable(original, tableEditBatches[0], { frame: "fit-grid" });
    slide.replaceElement(original.id, result.table);
    const blocked = await session.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.ok(blocked.report.diagnostics.some(d => d.code === "save-edit-unsupported")); assert.equal(blocked.plan, undefined);
    slide.replaceElement(original.id, original); const noop = await session.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready"); assert.deepEqual((await session.save(noop.plan)).bytes, source);
    const edited = structuredClone(original); edited.rows[0].cells[0].content[0].runs[0].text = "Allowed scalar"; slide.replaceElement(original.id, edited);
    assert.equal((await session.prepareSave({ mode: "edit" })).status, "ready");
  } finally { session.dispose(); }
});

test("edited header cells retain accessible roles after merge, split and insertion", async t => {
  const initial = { type: "table", id: "table", name: "Report", box: { x: 0, y: 0, width: 200, height: 60 }, columns: [100, 100], rows: [
    { height: 30, cells: [{ content: "North", headerScope: "column" }, { content: "South", headerScope: "column" }] },
    { height: 30, cells: [{ content: "10" }, { content: "20" }] },
  ] };
  const edits = [
    [{ type: "merge", region: { rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 2 }, content: "append-paragraphs" }],
    [{ type: "split", rowIndex: 0, columnIndex: 1 }],
    [{ type: "insert-row", index: 1, height: 30, cell: { content: "New", headerScope: "row" } }],
  ];
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), cdp = await page.context().newCDPSession(page);
  const global = await readFile("packages/core/dist/browser/global.js", "utf8");
  await page.addScriptTag({ content: global });
  let node = initial, browserTable = initial;
  const states = [];
  for (const operations of edits) {
    node = editTable(node, operations, { frame: "fit-grid" }).table;
    browserTable = await page.evaluate(({ table, operations }) => PPTKit.editTable(table, operations, { frame: "fit-grid" }).table, { table: browserTable, operations });
    assert.deepEqual(browserTable, node);
    states.push(node);
  }
  for (const [index, table] of states.entries()) {
    const deck = createPresentation();
    deck.addSlide({ elements: [table] });
    await page.setContent((await renderPresentationToSvg(deck)).slides[0].svg);
    const { nodes } = await cdp.send("Accessibility.getFullAXTree");
    const cells = nodes.filter(node => !node.ignored && ["columnheader", "rowheader", "cell"].includes(node.role?.value));
    assert.deepEqual(cells.map(node => node.role.value),
      index === 0 ? ["columnheader", "cell", "cell"] :
      index === 1 ? ["columnheader", "columnheader", "cell", "cell"] :
        ["columnheader", "columnheader", "rowheader", "rowheader", "cell", "cell"]);
    const native = await generatePptx(deck);
    assert.equal(native.warnings.filter(w => w.code === "table-header-semantics-not-exported").length, 1);
  }
});
