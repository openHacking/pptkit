import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, editTable } from "../packages/core/dist/index.js";
import { resolveLayout } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg, renderLayoutToSvg } from "../packages/svg-renderer/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { tableStyleDeck } from "../tests/fixtures/table-style.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
const profile = { projection: "content", editSaveProfile: "text-runs-v3" };
const semantics = document => {
  const table = resolveLayout(document).slides[0].elements[0];
  return { columns: table.columns, grid: table.borderGrid, rows: table.rows.map(row => ({ height: row.height, cells: row.cells.map(cell => ({ rowSpan: cell.rowSpan, colSpan: cell.colSpan, fill: cell.style.fill, margin: cell.style.margin, verticalAlign: cell.style.verticalAlign, content: cell.content })) })) };
};

test("region styles materialize once and three RGB fresh rounds retain effective cells, shared edges and exact no-op bytes", async () => {
  let document = tableStyleDeck(createPresentation); const expected = semantics(document), original = normalizePresentation(document);
  for (let round = 0; round < 3; round++) {
    const generated = await generatePptx(document);
    if (!round) assert.ok(generated.warnings.some(w => w.code === "table-style-materialized" && w.elementId === "styled-table"));
    else assert.equal(generated.warnings.some(w => w.code === "table-style-materialized"), false);
    const session = await importPptx(generated.bytes, profile);
    try {
      assert.deepEqual(semantics(session.document), expected);
      assert.equal(session.document.slides[0].elements[0].style, undefined, "fresh output carries effective native cells, not invented reusable rules");
      const plan = await session.prepareSave({ mode: "no-op" }); assert.equal(plan.status, "ready"); assert.deepEqual((await session.save(plan.plan)).bytes, generated.bytes);
      document = session.document;
    } finally { session.dispose(); }
  }
  assert.equal(original.slides[0].elements[0].styleSource, "materialized-regions");
});

test("Node/ESM/global normalization agrees and independent browser DOM verifies native and SVG effective cells", async t => {
  const document = tableStyleDeck(createPresentation), input = document.slides[0].elements[0], expected = normalizePresentation(document), layout = resolveLayout(document), table = layout.slides[0].elements[0], generated = await generatePptx(document), svg = await renderLayoutToSvg(layout);
  const browser = await chromium.launch(); t.after(() => browser.close());
  const esm = `data:text/javascript;base64,${Buffer.from(await readFile("packages/core/dist/browser/index.js", "utf8")).toString("base64")}`, global = await readFile("packages/core/dist/browser/global.js", "utf8");
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    if (format === "global") await page.addScriptTag({ content: global });
    const actual = await page.evaluate(async ({ esm, input, theme, id }) => { const core = esm ? await import(esm) : PPTKit; const d = core.createPresentation({ id, size: { width: 400, height: 300 }, theme }); d.addSlide({ id: "styled" }).addElement(input); return core.normalizePresentation(d); }, { esm: format === "esm" ? esm : null, input, theme: expected.theme, id: expected.id });
    assert.deepEqual(actual, expected);
    const inspected = await page.evaluate(({ xml, svg, table }) => {
      const ns = "http://schemas.openxmlformats.org/drawingml/2006/main", parse = (s, mime) => new DOMParser().parseFromString(s, mime), native = parse(xml,"application/xml"), preview = parse(svg,"image/svg+xml");
      const rows = [...native.getElementsByTagNameNS(ns,"tr")]; let count = 0;
      for (const at of table.cellLayout) {
        const cell = table.rows[at.rowIndex].cells[at.cellIndex], tc = [...rows[at.rowIndex].children].filter(n=>n.localName === "tc")[at.columnIndex], fill = [...tc.getElementsByTagNameNS(ns,"tcPr")[0].children].find(n=>n.namespaceURI===ns&&n.localName==="solidFill")?.firstElementChild;
        if (fill?.getAttribute("val") !== cell.style.fill.color) return { error: "native-fill" };
        const html = preview.querySelector(`[data-pptkit-cell="${at.rowIndex}:${at.cellIndex}"]`);
        if (!html) return { error: "preview-content" };
        const expectedText = cell.content.map(p=>p.runs.map(r=>r.text).join("")), authored = [...html.querySelectorAll("span[lang]")].map(n=>n.textContent).join("");
        if (authored !== expectedText.join("")) return { error: "preview-authored-content" };
        const copy = html.cloneNode(true), emptyLines = [...copy.querySelectorAll('span[aria-hidden="true"]')];
        if (emptyLines.length !== expectedText.filter(text=>text === "" || /[\r\n]$/.test(text)).length) return { error: "preview-empty-line-count" };
        for (const placeholder of emptyLines) { if (placeholder.textContent !== "\u200b") return { error: "preview-empty-line-content" }; placeholder.remove(); }
        if (copy.textContent !== expectedText.join("")) return { error: "preview-content" };
        if (!html.getAttribute("style").includes(`background:#${cell.style.fill.color}`)) return { error: "preview-fill" };
        const runs = [...tc.getElementsByTagNameNS(ns,"r")];
        if (runs.length !== cell.content.reduce((n,p)=>n+p.runs.length,0)) return { error: "native-runs" };
        const styles = cell.content.flatMap(p=>p.runs.map(r=>r.style));
        for (let i=0;i<runs.length;i++) {
          const props = runs[i].getElementsByTagNameNS(ns,"rPr")[0], style = styles[i];
          if (Number(props.getAttribute("sz"))/100 !== style.fontSize || (props.getAttribute("b") === "1") !== style.bold || (props.getAttribute("i") === "1") !== style.italic) return { error: "native-text-style" };
        }
        count++;
      }
      return { count };
    }, { xml: readGeneratedZip(generated.bytes).get("ppt/slides/slide1.xml").toString(), svg: svg.slides[0].svg, table });
    assert.deepEqual(inspected, { count: table.cellLayout.length }); assert.deepEqual(requests, []); await page.close();
  }
});

test("theme roles remain semantic across documents while direct overrides stay fixed and export remains explicit about native rules", async () => {
  const original = tableStyleDeck(createPresentation, { theme: true }), source = original.slides[0].elements[0], documents = [original];
  const changed = createPresentation({ size: { width: 400, height: 300 }, theme: { colors: { accent1: "9135A7", accent2: "196B24" }, fonts: { body: "Arial", heading: "Arial" } } }); changed.addSlide().addElement(source); documents.push(changed);
  for (const document of documents) {
    const table = normalizePresentation(document).slides[0].elements[0]; assert.deepEqual(table.rows[1].cells[0].style.fill.color, { theme: "accent1" }); assert.equal(table.rows[1].cells[1].style.fill.color, "FFCC66");
    const svg = (await renderPresentationToSvg(document)).slides[0].svg; assert.ok(svg.includes(document.theme.colors.accent1)); assert.ok(svg.includes("FFCC66"));
    const generated = await generatePptx(document), xml = readGeneratedZip(generated.bytes).get("ppt/slides/slide1.xml").toString(); assert.match(xml, /<a:schemeClr val="accent1">/); assert.match(xml, /<a:latin typeface="\+mn-lt"/); assert.ok(generated.warnings.some(w=>w.code === "table-style-materialized"));
  }
  assert.equal(original.slides[0].elements[0].style.wholeTable.text.run.fontFamily.theme, "body");
});

test("structure edits recompute regions before sizing/pagination and source rule mutations block atomically", async () => {
  const document = tableStyleDeck(createPresentation), slide = document.slides[0], before = slide.elements[0];
  const edited = editTable(before, [{ type: "insert-row", index: 3, height: 30 }], { frame: "fit-grid" }); slide.replaceElement(before.id, edited.table);
  const normalized = normalizePresentation(document), styled = normalized.slides[0].elements[0]; assert.equal(styled.rows[3].cells[0].style.fill.color,"DDEEFF"); assert.equal(styled.rows[4].cells[0].style.fill.color,"EEEEEE");
  const measured = [];
  const sized = resolveLayout(document, { tableSizing: [{ elementId: before.id, autoRows: [{ row: 4 }] }], textMeasurement: { measure: request => {
    const runs = request.paragraphs.flatMap(p => p.runs); measured.push(...runs.map(r => ({ text: r.text, fontSize: r.style.fontSize })));
    return { bounds: { x: 0, y: 0, width: 10, height: Math.max(...runs.map(r => r.style.fontSize), 1) * 2 } };
  } } });
  assert.ok(measured.some(r => r.text === "Local" && r.fontSize === 12));
  assert.equal(sized.slides[0].elements[0].rows[4].height, 28, "12pt local text height plus resolved 2pt top/bottom margins");
  const paged = resolveLayout(document, { tablePagination: [{ elementId: before.id, area: { x:30,y:30,width:320,height:120 }, headerRows:1 }] });
  assert.equal(paged.slides.length,2); assert.equal(paged.slides[1].elements[0].rows[0].cells[0].style.fill.color,"335577"); assert.equal(paged.slides[1].elements[0].rows.at(-1).cells[0].style.fill.color,"CCDDBB"); assert.deepEqual(normalizePresentation(document),normalized);
  const session = await importPptx((await generatePptx(document)).bytes,profile);
  try {
    const owner = session.document.slides[0], baseline = owner.elements[0]; owner.replaceElement(baseline.id,{ ...baseline, style: before.style });
    assert.equal((await session.prepareSave({mode:"edit"})).status,"blocked"); owner.replaceElement(baseline.id,baseline);
    const ready = await session.prepareSave({mode:"no-op"}); assert.equal(ready.status,"ready");
  } finally { session.dispose(); }
});

test("SVG reports native table style switches without altering effective cells", async () => {
  const deck = createPresentation();
  const base = { id: "flagged", type: "table", box: { x: 0, y: 0, width: 120, height: 60 },
    columns: [120], rows: [{ cells: [{ content: "Cell" }] }] };
  const slide = deck.addSlide({ id: "page", elements: [base] });
  const plain = await renderPresentationToSvg(deck);
  slide.replaceElement("flagged", { ...base, nativeFirstRowStyle: false, nativeTableStyleFlags: { bandedColumns: true } });
  const flagged = await renderPresentationToSvg(deck);
  assert.equal(flagged.slides[0].svg, plain.slides[0].svg);
  assert.deepEqual(flagged.warnings.filter(w => w.code === "native-table-style-flags-not-rendered").map(w => [w.slideId, w.elementId]), [["page", "flagged"]]);
});
