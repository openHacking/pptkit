import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { chromium } from "@playwright/test";
import { createPresentation, normalizePresentation, editTable } from "../packages/core/dist/index.js";
import { resolveLayout, resolveNormalizedLayout, resolveTableBorders } from "../packages/layout/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { renderPresentationToSvg, renderLayoutToSvg } from "../packages/svg-renderer/dist/index.js";
import { importPptx } from "../packages/pptx-parser/dist/index.js";
import { tableBorderDeck } from "../tests/fixtures/table-borders.mjs";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";
import { table as sourceTable, tableRow, cell, tableBody, tableBytes } from "../packages/pptx-parser/test/table-fixture.mjs";
const profile = { projection: "content", editSaveProfile: "text-runs-v3" };
const moduleUrl = source => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;

test("formatted and unknown empty merge continuations stay source-only in Node and browser", async t => {
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  await page.addScriptTag({ content: await readFile("packages/pptx-parser/dist/browser/global.js", "utf8") });
  for (const body of ['<a:p><a:r><a:rPr b="1"/><a:t/></a:r></a:p>', '<a:p><x:future xmlns:x="urn:unknown"/></a:p>', '<a:p><a:endParaRPr lang="en-US"/></a:p>']) {
    const markup = sourceTable({ rows: [tableRow([cell({ attributes: 'gridSpan="2"', text: "Anchor" }), cell({ attributes: 'hMerge="1"', body: tableBody(body) })], 50)] });
    const bytes = tableBytes(markup), session = await importPptx(bytes, { projection: "content" });
    try {
      const browserResult = await page.evaluate(async input => {
        const imported = await PPTKit.importPptx(Uint8Array.from(input), { projection: "content" });
        const plan = await imported.prepareSave({ mode: "no-op" }), saved = await imported.save(plan.plan);
        const result = { report: imported.report, count: imported.document.slides[0].elements.length, saved: [...saved.bytes] };
        imported.dispose(); return result;
      }, [...bytes]);
      assert.equal(browserResult.count, body.includes("en-US") ? 1 : 0);
      assert.equal(session.document.slides[0].elements.length, browserResult.count);
      assert.deepEqual(browserResult.report, session.report);
      assert.deepEqual(browserResult.saved, [...bytes]);
    } finally { session.dispose(); }
  }
});

test("table fill and stroke alpha survive current and legacy layout rendering paths", async t => {
  const document = createPresentation({ size: { width: 300, height: 200 }, theme: { colors: { accent2: "445566" } } });
  document.addSlide({ elements: [{ type: "table", box: { x: 0, y: 0, width: 150, height: 50 }, columns: [50, 50, 50], rows: [{ height: 50, cells: [
    { content: "Translucent", style: { borders: {}, fill: { type: "solid", color: "112233", opacity: 0.25 }, stroke: { width: 2, paint: { type: "solid", color: "AABBCC", opacity: 0.5 } } } },
    { content: "Clear", style: { fill: { type: "solid", color: { theme: "accent2" }, opacity: 0 }, stroke: { width: 2, paint: { type: "solid", color: { theme: "accent2" }, opacity: 0 } } } },
    { content: "Opaque", style: { fill: { type: "solid", color: "778899", opacity: 1 }, stroke: { width: 2, paint: { type: "solid", color: "123456", opacity: 1 } } } },
  ] }] }] });
  const current = resolveLayout(document), browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const mode of ["current", "without-border-grid", "legacy-html-table"]) {
    const layout = structuredClone(current), table = layout.slides[0].elements[0];
    if (mode !== "current") delete table.borderGrid;
    if (mode === "legacy-html-table") delete table.cellLayout;
    await page.setContent((await renderLayoutToSvg(layout)).slides[0].svg);
    const cells = page.locator(mode === "legacy-html-table" ? "td" : "[data-pptkit-cell]");
    assert.deepEqual(await cells.evaluateAll(nodes => nodes.map(node => getComputedStyle(node).backgroundColor)), ["rgba(17, 34, 51, 0.25)", "rgba(68, 85, 102, 0)", "rgb(119, 136, 153)"], mode);
    if (mode !== "current") {
      const borders = mode === "legacy-html-table" ? cells : page.locator('[data-pptkit-cell] > [aria-hidden="true"]');
      assert.deepEqual(await borders.evaluateAll(nodes => nodes.map(node => getComputedStyle(node).borderTopColor)), ["rgba(170, 187, 204, 0.5)", "rgba(68, 85, 102, 0)", "rgb(18, 52, 86)"], mode);
    }
  }
});

test("uniform dashed table strokes reach resolved SVG lines and native cell borders", async t => {
  const document = createPresentation({ size: { width: 300, height: 200 } });
  for (const dash of ["dash", "dot", "dashDot"]) {
    document.addSlide({ id: dash }).addElement({ id: `table-${dash}`, type: "table", box: { x: 10, y: 10, width: 100, height: 40 }, columns: [100], rows: [{ height: 40, cells: [{ content: dash, style: { stroke: { width: 2, dash, paint: { type: "solid", color: "123456", opacity: 0.5 } } } }] }] });
  }
  const [preview, native] = await Promise.all([renderPresentationToSvg(document), generatePptx(document)]);
  const entries = readGeneratedZip(native.bytes), browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const [index, [svgDash, nativeDash]] of [["8 6", "dash"], ["2 4", "sysDot"], ["8 4 2 4", "dashDot"]].entries()) {
    await page.setContent(preview.slides[index].svg);
    const lines = page.locator("[data-pptkit-table-border]");
    assert.equal(await lines.count(), 4);
    assert.deepEqual(await lines.evaluateAll(nodes => nodes.map(node => [node.getAttribute("stroke-dasharray"), node.getAttribute("stroke-opacity")])), Array(4).fill([svgDash, "0.5"]));
    const xml = entries.get(`ppt/slides/slide${index + 1}.xml`).toString();
    assert.equal([...xml.matchAll(new RegExp(`<a:prstDash val="${nativeDash}"`, "g"))].length, 4);
  }
});

test("uniform dashed borders survive three fresh export and import rounds", async () => {
  let document = createPresentation({ size: { width: 300, height: 200 } });
  document.addSlide().addElement({ id: "table", type: "table", box: { x: 10, y: 10, width: 100, height: 40 }, columns: [50, 50], rows: [{ height: 40, cells: [
    { content: "A", style: { stroke: { width: 2, dash: "dashDot", paint: { type: "solid", color: "123456", opacity: 0.5 } } } },
    { content: "B", style: { stroke: { width: 1, dash: "dot", paint: { type: "solid", color: "ABCDEF", opacity: 1 } } } },
  ] }] });
  const original = resolveLayout(document).slides[0].elements[0].borderGrid;
  for (let round = 0; round < 3; round++) {
    const result = await generatePptx(document), session = await importPptx(result.bytes, profile);
    try {
      const table = resolveLayout(session.document).slides[0].elements[0];
      assert.deepEqual(table.borderGrid, original, `round ${round + 1}`);
      assert.match((await renderLayoutToSvg(resolveLayout(session.document))).slides[0].svg, /stroke-dasharray="8 4 2 4"/);
      document = session.document;
    } finally { session.dispose(); }
  }
});

test("three fresh RGB table-border rounds retain the resolved physical graph, including segmented merge boundaries", async () => {
  let document = tableBorderDeck(createPresentation, { theme: false });
  const expected = resolveLayout(document).slides.map(s => s.elements[0].borderGrid);
  for (let round = 0; round < 3; round++) {
    const result = await generatePptx(document), session = await importPptx(result.bytes, profile);
    try {
      const merged = document.slides[1];
      assert.deepEqual(result.warnings.filter(w => w.code === "table-border-segmented-native").map(w => [w.slideId, w.elementId]), [[merged.id, merged.elements[0].id]]);
      const presets = result.warnings.filter(w => w.code === "native-table-preset-direct-formatting");
      assert.deepEqual(presets.map(w => [w.slideId, w.elementId]), round === 0 ? [] : document.slides.map(slide => [slide.id, slide.elements[0].id]));
      assert.equal(result.warnings.length, 1 + presets.length);
      assert.deepEqual(resolveLayout(session.document).slides.map(s => s.elements[0].borderGrid), expected);
      assert.ok(session.report.findings.some(f => /physical right border segments/.test(f.reason)));
      const noop = await session.prepareSave({ mode: "no-op" }); assert.equal(noop.status, "ready"); assert.deepEqual((await session.save(noop.plan)).bytes, result.bytes);
      document = session.document;
    } finally { session.dispose(); }
  }
});

test("Node/ESM/global share exact grid/work quotas and independent DOM verifies every native edge, SVG line and Office merge continuation", async t => {
  const document = tableBorderDeck(createPresentation), normalized = normalizePresentation(document), expected = resolveLayout(document), pptx = await generatePptx(document), entries = readGeneratedZip(pptx.bytes), svg = await renderPresentationToSvg(document);
  const browser = await chromium.launch(); t.after(() => browser.close());
  const coreModule = moduleUrl(await readFile("packages/core/dist/browser/index.js", "utf8"));
  const esm = moduleUrl((await readFile("packages/layout/dist/browser/index.js", "utf8")).replaceAll('"@pptkit/core"', JSON.stringify(coreModule)));
  const global = await readFile("packages/layout/dist/browser/global.js", "utf8"), coreGlobal = await readFile("packages/core/dist/browser/global.js", "utf8");
  let lo = 1, hi = 2000000;
  while (lo < hi) { const mid = Math.floor((lo + hi) / 2); try { resolveTableBorders(normalized.slides[0].elements[0], { maxWorkUnits: mid }); hi = mid; } catch (e) { assert.equal(e.code, "table-border-limit-exceeded"); lo = mid + 1; } }
  for (const format of ["esm", "global"]) {
    const page = await browser.newPage(), requests = []; await page.route("**/*", route => { requests.push(route.request().url()); return route.abort(); });
    if (format === "global") { await page.addScriptTag({ content: coreGlobal }); await page.addScriptTag({ content: global }); }
    const computed = await page.evaluate(async ({ normalized, esm, work }) => {
      const api = esm ? await import(esm) : PPTKit, layout = api.resolveNormalizedLayout(normalized);
      const direct = api.resolveTableBorders(normalized.slides[0].elements[0], { maxWorkUnits: work }); let quota;
      try { api.resolveTableBorders(normalized.slides[0].elements[0], { maxWorkUnits: work - 1 }); } catch (e) { quota = e.code; }
      return { layout, direct, quota };
    }, { normalized, esm: format === "esm" ? esm : null, work: lo });
    assert.deepEqual(computed.layout, expected); assert.deepEqual(computed.direct, resolveTableBorders(normalized.slides[0].elements[0])); assert.equal(computed.quota, "table-border-limit-exceeded");
    for (let index = 0; index < normalized.slides.length; index++) {
      const xml = entries.get(`ppt/slides/slide${index + 1}.xml`).toString(), table = expected.slides[index].elements[0];
      const result = await page.evaluate(({ xml, svg, table, theme }) => {
        const parse = (text, type) => new DOMParser().parseFromString(text, type), doc = parse(xml, "application/xml"), a = "http://schemas.openxmlformats.org/drawingml/2006/main", native = doc.getElementsByTagNameNS(a, "tbl")[0], rows = [...native.getElementsByTagNameNS(a, "tr")], grid = table.borderGrid;
        const role = { background1: "bg1", background2: "bg2", text1: "tx1", text2: "tx2", accent1: "accent1", accent2: "accent2", accent3: "accent3", accent4: "accent4", accent5: "accent5", accent6: "accent6", hyperlink: "hlink", followedHyperlink: "folHlink" };
        const sameLine = (line, style) => {
          if (!line || Number(line.getAttribute("w")) !== Math.round(style.width * 12700)) return false;
          const fill = [...line.children].find(n => ["noFill", "solidFill"].includes(n.localName));
          if (style.paint.type === "none") return fill?.localName === "noFill";
          const color = fill?.firstElementChild, expected = style.paint.color;
          if (typeof expected === "string" ? color?.localName !== "srgbClr" || color.getAttribute("val") !== expected.replace(/^#/, "").toUpperCase() : color?.localName !== "schemeClr" || color.getAttribute("val") !== role[expected.theme]) return false;
          const alpha = color.getElementsByTagNameNS(a, "alpha")[0];
          return (alpha ? Number(alpha.getAttribute("val")) : 100000) === Math.round(style.paint.opacity * table.opacity * 100000)
            && line.getElementsByTagNameNS(a, "prstDash")[0].getAttribute("val") === (style.dash === "dot" ? "sysDot" : style.dash);
        };
        const owners = Array.from({ length: grid.rowCount }, () => Array(grid.columnCount).fill(null));
        for (const at of table.cellLayout) {
          const cell = table.rows[at.rowIndex].cells[at.cellIndex], owner = { row: at.rowIndex, col: at.columnIndex, rs: cell.rowSpan, cs: cell.colSpan };
          for (let r = owner.row; r < owner.row + owner.rs; r++) for (let c = owner.col; c < owner.col + owner.cs; c++) owners[r][c] = owner;
        }
        let nativeEdges = 0;
        for (let r = 0; r < grid.rowCount; r++) {
          const cells = [...rows[r].children].filter(n => n.localName === "tc"); if (cells.length !== grid.columnCount) return { error: "physical-width" };
          for (let c = 0; c < grid.columnCount; c++) {
            const cell = cells[c], owner = owners[r][c], props = cell.getElementsByTagNameNS(a, "tcPr")[0];
            let edges = [["lnL", grid.vertical[r * (grid.columnCount + 1) + c]], ["lnR", grid.vertical[r * (grid.columnCount + 1) + c + 1]], ["lnT", grid.horizontal[r * grid.columnCount + c]], ["lnB", grid.horizontal[(r + 1) * grid.columnCount + c]]];
            if (owner && r === owner.row && c === owner.col && (owner.rs > 1 || owner.cs > 1)) {
              const outer = [
                Array.from({ length: owner.rs }, (_, i) => grid.vertical[(r + i) * (grid.columnCount + 1) + c]),
                Array.from({ length: owner.rs }, (_, i) => grid.vertical[(r + i) * (grid.columnCount + 1) + c + owner.cs]),
                Array.from({ length: owner.cs }, (_, i) => grid.horizontal[r * grid.columnCount + c + i]),
                Array.from({ length: owner.cs }, (_, i) => grid.horizontal[(r + owner.rs) * grid.columnCount + c + i]),
              ];
              edges = outer.map((side, i) => [edges[i][0], side.every(s => JSON.stringify(s) === JSON.stringify(side[0])) ? side[0] : { ...side[0], width: 0, paint: { type: "none" } }]);
            }
            for (const [tag, style] of edges) { if (!sameLine(props.getElementsByTagNameNS(a, tag)[0], style)) return { error: `edge:${r}:${c}:${tag}` }; nativeEdges++; }
            if (owner) {
              if ((cell.getAttribute("vMerge") === "1") !== (r > owner.row) || (cell.getAttribute("hMerge") === "1") !== (c > owner.col)) return { error: "merge-flags" };
              const expectedRows = r === owner.row ? owner.rs : 1, expectedColumns = c === owner.col ? owner.cs : 1;
              if (Number(cell.getAttribute("rowSpan") ?? 1) !== expectedRows || Number(cell.getAttribute("gridSpan") ?? 1) !== expectedColumns) return { error: "Office-merge-spans" };
            }
          }
        }
        const preview = parse(svg, "image/svg+xml"), lines = [...preview.querySelectorAll("[data-pptkit-table-border]")];
        const visible = [...grid.horizontal, ...grid.vertical].filter(s => s.paint.type === "solid" && s.width > 0 && s.paint.opacity > 0).length;
        if (lines.length !== visible) return { error: "preview-count" };
        for (const line of lines) {
          const [axis, r, c] = line.getAttribute("data-pptkit-table-border").split(":");
          const style = axis === "h" ? grid.horizontal[Number(r) * grid.columnCount + Number(c)] : grid.vertical[Number(r) * (grid.columnCount + 1) + Number(c)];
          const color = typeof style.paint.color === "string" ? style.paint.color : theme.colors[style.paint.color.theme];
          if (line.getAttribute("stroke").replace(/^#/, "").toUpperCase() !== color.replace(/^#/, "").toUpperCase() || Number(line.getAttribute("stroke-width")) !== style.width || Number(line.getAttribute("stroke-opacity")) !== style.paint.opacity) return { error: "preview-style" };
          if (Number(line.parentElement.parentElement.getAttribute("opacity")) !== table.opacity) return { error: "double-opacity" };
        }
        const htmlCells = [...preview.querySelectorAll("[data-pptkit-cell]")];
        if (htmlCells.length !== table.cellLayout.length || htmlCells.some(n => n.querySelector('[style*="border:"]'))) return { error: "double-border" };
        return { nativeEdges, previewLines: lines.length };
      }, { xml, svg: svg.slides[index].svg, table, theme: expected.theme });
      assert.equal(result.error, undefined, `${format}/${index}/${result.error}`); assert.ok(result.nativeEdges > 0);
    }
    await page.setContent(svg.slides[3].svg);
    const rgba = await page.locator('[data-pptkit-cell="0:0"]').evaluate(n => getComputedStyle(n).backgroundColor); assert.equal(rgba, "rgba(17, 34, 51, 0.25)");
    assert.deepEqual(requests, []); await page.close();
  }
});

test("edits, sizing and pagination recompute border geometry, and unsupported source border writes remain blocked", async () => {
  const document = tableBorderDeck(createPresentation, { theme: false }), slide = document.slides[2], original = slide.elements[0];
  const result = editTable(original, [{ type: "split", rowIndex: 1, columnIndex: 1 }, { type: "insert-row", index: 1, height: 20, cell: { content: "", style: { margin: 0 } } }], { frame: "fit-grid" }); slide.replaceElement(original.id, result.table);
  const before = normalizePresentation(document), options = { tableSizing: [{ elementId: original.id, autoRows: result.table.rows.map((_, row) => ({ row })) }], textMeasurement: { measure: () => ({ bounds: { x: 0, y: 0, width: 10, height: 40 } }) }, tablePagination: [{ elementId: original.id, area: { x: 10, y: 10, width: 180, height: 85 } }] };
  const layout = resolveLayout(document, options), fragments = layout.slides.filter(s => s.elements[0]?.id === original.id || s.elements[0]?.id.startsWith("pptkit:table:"));
  assert.equal(fragments.length, 2); assert.ok(fragments.every(s => s.elements[0].borderGrid.rowCount === 2)); assert.deepEqual(normalizePresentation(document), before);
  const source = (await generatePptx(document)).bytes, session = await importPptx(source, profile);
  try {
    const owner = session.document.slides[0], previous = owner.elements[0], edited = structuredClone(previous); edited.rows[0].cells[0].style.borders.left.width = 30; owner.replaceElement(previous.id, edited);
    const blocked = await session.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.ok(blocked.report.diagnostics.some(d => d.code === "save-edit-unsupported"));
  } finally { session.dispose(); }
});

test("JSON Layout preserves border rendering while unknown versions, truncated edges and missing geometry reject", async () => {
  const document = tableBorderDeck(createPresentation), layout = resolveLayout(document);
  assert.deepEqual(await renderLayoutToSvg(JSON.parse(JSON.stringify(layout))), await renderPresentationToSvg(document));
  for (const mutate of [t => t.borderGrid.schemaVersion = "future", t => t.borderGrid.rowCount++, t => t.borderGrid.horizontal.pop(), t => delete t.cellLayout]) {
    const bad = structuredClone(layout); mutate(bad.slides[0].elements[0]); await assert.rejects(renderLayoutToSvg(bad), /table border Layout metadata/);
  }
});
