import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { test, expect } from "@playwright/test";
import { readFileSync } from "node:fs";

const read = (path) => readFileSync(new URL(path, import.meta.url));
const moduleUrl = (path) => `data:text/javascript;base64,${read(path).toString("base64")}`;
const latin = read("./fixtures/roboto-mono-latin-400-normal.woff2").toString("base64");
const cjk = read("../../layout/test/fixtures/noto-sc-measurement.otf").toString("base64");
const international = read("../../layout/test/fixtures/noto-kr-tab-measurement.otf").toString("base64");
const coverage = JSON.parse(read("../../layout/test/fixtures/cjk-coverage.json"));

test.beforeEach(async ({ page }) => {
  await page.route("**/*", (route) => route.abort());
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports: {
    "@pptkit/core": moduleUrl("../../core/dist/browser/index.js"),
    "@pptkit/layout": moduleUrl("../../layout/dist/browser/index.js"),
    "@pptkit/layout/dom": moduleUrl("../../layout/dist/browser/dom.js"),
    "@pptkit/svg-renderer": moduleUrl("../dist/browser/index.js"),
    "@pptkit/pptx-exporter": moduleUrl("../../pptx-exporter/dist/browser/index.js"),
  } })}</script>`);
  await page.evaluate(async ({ latin, cjk, international, coverage }) => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { createBrowserTextMeasurement } = await import("@pptkit/layout/dom");
    const latinFace = new FontFace("Pinned Mono", `url(data:font/woff2;base64,${latin})`);
    const cjkFace = new FontFace("Pinned CJK", `url(data:font/otf;base64,${cjk})`);
    const internationalFace = new FontFace("Pinned International", `url(data:font/otf;base64,${international})`);
    await Promise.all([latinFace.load(), cjkFace.load(), internationalFace.load()]);
    document.fonts.add(latinFace);
    document.fonts.add(cjkFace);
    document.fonts.add(internationalFace);
    const cmap = new Set(coverage);
    const fonts = [
      { family: "Pinned Mono", faces: [latinFace], covers: ({ text, bold, italic }) => !bold && !italic && /^[\x00-\x7f]*$/.test(text) },
      { family: "Pinned CJK", faces: [cjkFace], covers: ({ text, bold, italic }) => !bold && !italic && [...text].every((c) => /[\n\r\t]/.test(c) || cmap.has(c.codePointAt(0))) },
      { family: "Pinned International", faces: [internationalFace], covers: ({ text, bold, italic }) => !bold && !italic && [...text].every(c => /[\x20-\x7e\n\r\t]/u.test(c) || "あかアカ가한글中文".includes(c)) },
    ];
    const adapter = createBrowserTextMeasurement({ document, fonts });
    window.harness = {
      fonts, latinFace, cjkFace, adapter, createBrowserTextMeasurement,
      analyze(content, { width = 200, height = 500, margin = 0, wrap = true, provider = adapter, verticalAlign = "top", fontScale = 1 } = {}) {
        const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
        const paragraphs = typeof content === "string" ? [{ style: { lineSpacing: 1.5 }, runs: [{ text: content, style: { fontSize: 20 } }] }] : content;
        deck.addSlide({ id: "slide" }).addElement({ type: "text", id: "text", box: { x: 0, y: 0, width, height }, frame: { margin, wrap, verticalAlign, autoFit: { mode: "shrink", fontScale } }, content: paragraphs });
        return resolveLayout(deck, { textMeasurement: provider, diagnostics: { textOverflow: true } }).analysis;
      },
      width(text, size = 20, family = "Pinned Mono") {
        const context = document.createElement("canvas").getContext("2d");
        context.font = `${size}px "${family}"`;
        return context.measureText(text).width;
      },
    };
  }, { latin, cjk, international, coverage });
});

test("loaded real font advances agree with independent Canvas metrics and preserve fractional precision", async ({ page }) => {
  const result = await page.evaluate(() => {
    const { analyze, width } = harness;
    return { analysis: analyze("MMMMiiii"), expected: width("MMMMiiii") };
  });
  expect(result.analysis.diagnostics).toEqual([]);
  expect(result.analysis.measurements[0].bounds.width).toBeCloseTo(result.expected, 1);
  expect(result.analysis.measurements[0].bounds.height).toBeCloseTo(30, 1);
});

test("rotated table bounds track browser glyph positions using loaded font metrics", async ({ page }) => {
  const result=await page.evaluate(async()=>{
    const {createPresentation}=await import('@pptkit/core'),{resolveLayout}=await import('@pptkit/layout'),{renderPresentationToSvg}=await import('@pptkit/svg-renderer');
    const d=createPresentation({size:{width:240,height:160},theme:{fonts:{body:'Pinned Mono',heading:'Pinned Mono'}}});
    d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:20,y:20,width:200,height:120},columns:[100,100],rows:[{height:120,cells:['rotate90','rotate270'].map(textDirection=>({content:[{style:{lineSpacing:1.5},runs:[{text:'ABCD',style:{fontSize:16}}]}],style:{textDirection,margin:10,verticalAlign:'top'}}))}]});
    const layout={textMeasurement:harness.adapter,diagnostics:{textOverflow:true}},resolved=resolveLayout(d,layout),svg=await renderPresentationToSvg(d,{layout});const host=document.createElement('div');host.innerHTML=svg.slides[0].svg;document.body.appendChild(host);
    const root=host.querySelector('svg').getBoundingClientRect(),positions=[...host.querySelectorAll('[role="cell"] span')].map(span=>{const box=span.getBoundingClientRect();return {x:box.x-root.x,y:box.y-root.y,width:box.width,height:box.height};});
    return {positions,analysis:resolved.analysis};
  });
  expect(result.analysis.diagnostics.filter(d=>d.code==='table-text-direction-measurement-approximate')).toHaveLength(2);
  expect(result.analysis.diagnostics.some(d=>d.code==='text-overflow')).toBe(false);
  for(let i=0;i<2;i++){
    const measured=result.analysis.measurements[i].bounds,actual=result.positions[i];
    expect(Math.abs(actual.x-(20+measured.x))).toBeLessThan(6);
    expect(Math.abs(actual.y-(20+measured.y))).toBeLessThan(6);
    expect(Math.abs(actual.width-measured.width)).toBeLessThan(6);
    expect(Math.abs(actual.height-measured.height)).toBeLessThan(6);
  }
});

test("rotated table auto row sizing fits loaded-font glyphs in preview and exported row geometry", async ({ page }) => {
  const result=await page.evaluate(async()=>{
    const {createPresentation}=await import('@pptkit/core'),{resolveLayout}=await import('@pptkit/layout'),{renderPresentationToSvg}=await import('@pptkit/svg-renderer'),{generatePptx}=await import('@pptkit/pptx-exporter');
    const d=createPresentation({size:{width:240,height:160},theme:{fonts:{body:'Pinned Mono',heading:'Pinned Mono'}}});
    d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:20,y:20,width:160,height:20},columns:[80,80],rows:[{height:20,cells:['rotate90','rotate270'].map(textDirection=>({content:[{style:{lineSpacing:1.5},runs:[{text:'ABCD',style:{fontSize:16}}]}],style:{textDirection,margin:5}}))}]});
    const layout={textMeasurement:harness.adapter,diagnostics:{textOverflow:true},tableSizing:[{elementId:'t',autoRows:[{row:0}]}]},resolved=resolveLayout(d,layout),preview=await renderPresentationToSvg(d,{layout}),pptx=await generatePptx(d,{layout});
    const host=document.createElement('div');host.innerHTML=preview.slides[0].svg;document.body.appendChild(host);const cells=[...host.querySelectorAll('[role="cell"]')].map(cell=>{const outer=cell.getBoundingClientRect(),inner=cell.querySelector('span').getBoundingClientRect();return {outer:{x:outer.x,y:outer.y,width:outer.width,height:outer.height},inner:{x:inner.x,y:inner.y,width:inner.width,height:inner.height}}});
    return {height:resolved.slides[0].elements[0].rows[0].height,diagnostics:resolved.analysis.diagnostics,cells,bytes:[...pptx.bytes]};
  });
  expect(result.diagnostics.filter(x=>x.code==='table-text-direction-sizing-approximate')).toHaveLength(2);
  expect(result.diagnostics.some(x=>x.code==='table-cell-overflow'||x.code==='text-overflow')).toBe(false);
  expect(result.height).toBeGreaterThan(40);
  for(const {outer,inner} of result.cells){expect(inner.x).toBeGreaterThanOrEqual(outer.x-2);expect(inner.y).toBeGreaterThanOrEqual(outer.y-2);expect(inner.x+inner.width).toBeLessThanOrEqual(outer.x+outer.width+2);expect(inner.y+inner.height).toBeLessThanOrEqual(outer.y+outer.height+2);}
  const xml=readGeneratedZip(Uint8Array.from(result.bytes)).get('ppt/slides/slide1.xml').toString();expect(xml).toContain(`<a:tr h="${Math.round(result.height*12700)}">`);
});

test("rotated table row can use wrapping to satisfy a finite height maximum", async ({ page }) => {
  const result=await page.evaluate(async()=>{
    const {createPresentation}=await import('@pptkit/core'),{resolveLayout}=await import('@pptkit/layout');
    const d=createPresentation({size:{width:200,height:160},theme:{fonts:{body:'Pinned Mono',heading:'Pinned Mono'}}});
    d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:0,y:0,width:80,height:20},columns:[80],rows:[{height:20,cells:[{content:[{style:{lineSpacing:1.5},runs:[{text:'MMMM MMMM',style:{fontSize:16}}]}],style:{textDirection:'rotate90',margin:5}}]}]});
    const layout=resolveLayout(d,{textMeasurement:harness.adapter,diagnostics:{textOverflow:true},tableSizing:[{elementId:'t',autoRows:[{row:0,maxHeight:70}]}]});return {height:layout.slides[0].elements[0].rows[0].height,diagnostics:layout.analysis.diagnostics,measurement:layout.analysis.measurements[0]};
  });
  expect(result.height).toBeGreaterThan(20);
  expect(result.height).toBeLessThan(70);
  expect(result.diagnostics.some(x=>x.code==='table-text-direction-wrapped-fallback')).toBe(true);
  expect(result.diagnostics.some(x=>x.code==='table-row-overflow'||x.code==='table-cell-overflow'||x.code==='text-overflow')).toBe(false);
  expect(result.measurement.bounds.width).toBeGreaterThan(24);
});

test("compact rotated sizing reduces unbounded row height with real font metrics", async ({ page }) => {
  const result=await page.evaluate(async()=>{
    const {createPresentation}=await import('@pptkit/core'),{resolveLayout}=await import('@pptkit/layout');
    const d=createPresentation({size:{width:200,height:160},theme:{fonts:{body:'Pinned Mono',heading:'Pinned Mono'}}});
    d.addSlide({id:'s'}).addElement({id:'t',type:'table',box:{x:0,y:0,width:80,height:20},columns:[80],rows:[{height:20,cells:[{content:[{style:{lineSpacing:1.5},runs:[{text:'MMMM MMMM',style:{fontSize:16}}]}],style:{textDirection:'rotate270',margin:5}}]}]});
    const options=rotatedTextSizing=>({textMeasurement:harness.adapter,diagnostics:{textOverflow:true},tableSizing:[{elementId:'t',autoRows:[{row:0}],...(rotatedTextSizing?{rotatedTextSizing}:{})}]}),normal=resolveLayout(d,options()),compact=resolveLayout(d,options('compact'));
    return {normal:normal.slides[0].elements[0].rows[0].height,compact:compact.slides[0].elements[0].rows[0].height,diagnostics:compact.analysis.diagnostics};
  });
  expect(result.compact).toBeLessThan(result.normal-10);
  expect(result.diagnostics.some(x=>x.code==='table-text-direction-wrapped-fallback')).toBe(true);
  expect(result.diagnostics.some(x=>['table-row-overflow','table-cell-overflow','text-overflow'].includes(x.code))).toBe(false);
});

test("browser measurement refuses tab cases outside the explicit-stop profile", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const runs = text => [{ text, style: { fontSize: 20 } }];
    const inputs = [
      [{ runs: runs("A\tB") }],
      [{ style: { tabStops: [{ position: 120, align: "decimal" }] }, runs: runs("A\tB") }],
      [{ style: { tabStops: [{ position: 5, align: "left" }] }, runs: runs("A\tB") }],
      [{ style: { tabStops: [] }, runs: runs("\t") }],
      [{ runs: runs("A\nB\tC") }],
      [{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: runs("A\t") }],
      [{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: runs("A\t\nB") }],
      [{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: runs(`A\tB${"\n".repeat(4097)}`) }],
      [{ style: { indent: 2, hanging: 5, tabStops: [{ position: 120, align: "left" }] }, runs: runs("A\tB") }],
      [{ style: { align: "right", tabStops: [{ position: 120, align: "left" }] }, runs: runs("A\tB") }],
      [{ style: { tabStops: [{ position: 40, align: "right" }] }, runs: runs("AAA\tBBBB") }],
      [{ style: { tabStops: [{ position: 80, align: "left" }] }, runs: runs("A\tB\tC") }],
      [{ style: { tabStops: Array.from({ length: 257 }, (_, i) => ({ position: i + 1, align: "left" })) }, runs: runs("A\tB") }],
    ];
    const failures = inputs.map(content => harness.analyze(content, { width: 80, wrap: false }));
    const plain = harness.analyze([{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: runs("AB") }], { wrap: false });
    const custom = { resolveFont: harness.adapter.resolveFont, measure: () => ({ bounds: { x: 0, y: 0, width: 10, height: 10 } }) };
    const external = harness.analyze(inputs[0], { provider: custom });
    const { createPresentation } = await import("@pptkit/core"), { resolveLayout } = await import("@pptkit/layout");
    const { normalizePresentation } = await import("@pptkit/core");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    const slide = deck.addSlide({ id: "containers" }), box = { x: 0, y: 0, width: 200, height: 80 };
    slide.addElement({ id: "shape", type: "shape", shape: "rect", box, text: { content: "A\tB" } });
    slide.addElement({ id: "table", type: "table", box, columns: [200], rows: [{ height: 80, cells: [{ content: "C\tD" }] }] });
    const containers = resolveLayout(deck, { textMeasurement: harness.adapter, diagnostics: { textOverflow: true } }).analysis;
    const rejectedScripts = ["A\tא", "A\tا", "A\tᄀ", "A\te\u0301", "A\t😀", "A\t中\uFE0F", "A\tか\u3099", "A\t中\u{16FF0}"].map(text => {
      const p = createPresentation(); p.addSlide().addElement({ type: "text", box, content: [{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: [{ text, style: { fontFamily: "Pinned Mono", fontSize: 20 } }] }] });
      const paragraph = normalizePresentation(p).slides[0].elements[0].content[0];
      try { harness.adapter.measure({ paragraphs: [paragraph], availableWidth: 200, wrap: false }); return false; }
      catch (error) { return String(error.message).includes("tab-stop measurement"); }
    });
    return { failures, plain, external, containers, rejectedScripts, hostCount: document.querySelectorAll("[data-pptkit-measurement]").length };
  });
  for (const failure of result.failures) {
    expect(failure.measurements).toEqual([]);
    expect(failure.diagnostics).toHaveLength(1);
    expect(failure.diagnostics[0]).toMatchObject({ code: "text-measurement-failed", slideId: "slide", elementId: "text" });
    expect(failure.diagnostics[0].message).toContain("tab-stop measurement");
  }
  expect(result.plain.diagnostics.some(d => d.code === "text-measurement-failed")).toBe(false); expect(result.plain.measurements).toHaveLength(1);
  expect(result.external.diagnostics).toEqual([]); expect(result.external.measurements).toHaveLength(1);
  expect(result.containers.measurements).toEqual([]);
  expect(result.containers.diagnostics.filter(d => d.code === "text-measurement-failed").map(d => d.elementId)).toEqual(["shape", "table"]);
  expect(result.rejectedScripts).toEqual([true, true, true, true, true, true, true, true]);
  expect(result.hostCount).toBe(1);
});

test("browser measurement admits plain RTL paragraphs but rejects unsupported RTL tabs", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation, normalizePresentation } = await import('@pptkit/core');
    const paragraph = (direction, text) => {
      const p = createPresentation();
      p.addSlide().addElement({ type: 'text', box: { x: 0, y: 0, width: 250, height: 80 }, content: [{ style: { direction, tabStops: [{ position: 120, align: 'left' }] }, runs: [{ text, style: { fontFamily: 'Pinned Mono', fontSize: 20 } }] }] });
      return normalizePresentation(p).slides[0].elements[0].content[0];
    };
    const plain = harness.adapter.measure({ paragraphs: [paragraph('rtl', 'ABC 123')], availableWidth: 250, wrap: false });
    const ltr = harness.adapter.measure({ paragraphs: [paragraph('ltr', 'ABC\t123')], availableWidth: 250, wrap: false });
    let rejected = false;
    try { harness.adapter.measure({ paragraphs: [paragraph('rtl', 'ABC\t123')], availableWidth: 250, wrap: false }); }
    catch (error) { rejected = error.message.includes('LTR direction'); }
    return { plain, ltr, rejected };
  });
  expect(result.plain.bounds.width).toBeGreaterThan(0);
  expect(result.ltr.bounds.width).toBeGreaterThan(0);
  expect(result.rejected).toBe(true);
});

test("browser measurement rejects run-level RTL overrides it cannot reproduce", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation, normalizePresentation } = await import('@pptkit/core');
    const paragraph = rightToLeft => {
      const p = createPresentation();
      p.addSlide().addElement({ type: 'text', box: { x: 0, y: 0, width: 250, height: 80 }, content: [{ runs: [{ text: '123 -', style: { fontFamily: 'Pinned Mono', fontSize: 20, rightToLeft } }] }] });
      return normalizePresentation(p).slides[0].elements[0].content[0];
    };
    const normal = harness.adapter.measure({ paragraphs: [paragraph(false)], availableWidth: 250, wrap: false });
    let rejected = false;
    try { harness.adapter.measure({ paragraphs: [paragraph(true)], availableWidth: 250, wrap: false }); }
    catch (error) { rejected = error.message.includes('run-level RTL'); }
    return { normal, rejected };
  });
  expect(result.normal.bounds.width).toBeGreaterThan(0);
  expect(result.rejected).toBe(true);
});

test("explicit tab stops restart after hard breaks, including CRLF split across runs and empty lines", async ({ page }) => {
  const result = await page.evaluate(() => {
    const style = { tabStops: [{ position: 120, align: "left" }], lineSpacing: 1.5, indent: 36, hanging: 18 };
    const font = { fontSize: 20 };
    const split = harness.analyze([{ style, runs: [{ text: "A\tB\r", style: font }, { text: "\nC\tWIDE\n\nZ", style: font }] }], { width: 400, wrap: false });
    const trailing = harness.analyze([{ style, runs: [{ text: "A\tB\n", style: font }] }], { width: 400, wrap: false });
    const equivalent = harness.analyze([{ style, runs: [{ text: "A\tB\nC\tWIDE\n\nZ", style: font }] }], { width: 400, wrap: false });
    const repeated = harness.analyze([{ style, runs: [{ text: "\nA\tB\r\rC\tD\r\n", style: font }] }], { width: 400, wrap: false });
    const hanging = harness.analyze([{ style: { ...style, tabStops: [{ position: 0, align: "left" }, { position: 20, align: "left" }] }, runs: [{ text: "A\tB\nA\tB", style: font }] }], { width: 400, wrap: false });
    const emptyRun = harness.analyze([{ style, runs: [{ text: "A", style: font }, { text: "", style: { fontSize: 30 } }, { text: "\tB", style: font }] }], { width: 400, wrap: false });
    const emptyControl = harness.analyze([{ style, runs: [{ text: "A", style: font }, { text: "", style: { fontSize: 30 } }, { text: " B", style: font }] }], { width: 400, wrap: false });
    const styledTab = harness.analyze([{ style, runs: [{ text: "A", style: font }, { text: "\t", style: { fontSize: 30 } }, { text: "B", style: font }] }], { width: 400, wrap: false });
    const styledControl = harness.analyze([{ style, runs: [{ text: "A", style: font }, { text: " ", style: { fontSize: 30 } }, { text: "B", style: font }] }], { width: 400, wrap: false });
    const styledBreak = harness.analyze([{ style, runs: [{ text: "A\tB", style: font }, { text: "\n", style: { fontSize: 30 } }, { text: "C", style: font }] }], { width: 400, wrap: false });
    const breakControl = harness.analyze([{ style, runs: [{ text: "A B", style: font }, { text: "\n", style: { fontSize: 30 } }, { text: "C", style: font }] }], { width: 400, wrap: false });
    return { split, trailing, equivalent, repeated, hanging, emptyRun, emptyControl, styledTab, styledControl, styledBreak, breakControl, expectedHanging: 56 + harness.width("B"), expectedWidth: 156 + harness.width("WIDE") };
  });
  for (const analysis of [result.split, result.trailing, result.equivalent, result.repeated, result.hanging, result.emptyRun, result.emptyControl, result.styledTab, result.styledControl]) expect(analysis.diagnostics).toEqual([]);
  expect(result.split.measurements[0].bounds.width).toBeCloseTo(result.expectedWidth, 1);
  expect(result.split.measurements[0].bounds.height).toBeCloseTo(120, 1);
  expect(result.trailing.measurements[0].bounds.height).toBeCloseTo(60, 1);
  expect(result.repeated.measurements[0].bounds.height).toBeCloseTo(150, 1);
  expect(result.hanging.measurements[0].bounds.width).toBeCloseTo(result.expectedHanging, 1);
  expect(result.emptyRun.measurements[0].bounds.height).toBeCloseTo(result.emptyControl.measurements[0].bounds.height, 1);
  expect(result.styledTab.measurements[0].bounds.height).toBeCloseTo(result.styledControl.measurements[0].bounds.height, 1);
  expect(result.styledBreak.measurements[0].bounds.height).toBeCloseTo(result.breakControl.measurements[0].bounds.height, 1);
  expect(result.split.measurements[0].bounds).toEqual(result.equivalent.measurements[0].bounds);
});

test("tab measurement accepts its exact 4096-control work limit without partial geometry", async ({ page }) => {
  const result = await page.evaluate(() => {
    const row = `A${"\t".repeat(255)}B`, source = `${Array(16).fill(row).join("\n")}\n`;
    const tabStops = Array.from({ length: 256 }, (_, index) => ({ position: 20 * (index + 1), align: "left" }));
    return { controls: [...source].filter(char => char === "\t" || char === "\n" || char === "\r").length,
      analysis: harness.analyze([{ style: { tabStops, lineSpacing: 1.5 }, runs: [{ text: source, style: { fontSize: 20 } }] }], { width: 6000, height: 600, wrap: false }), expectedWidth: 5100 + harness.width("B") };
  });
  expect(result.controls).toBe(4096); expect(result.analysis.diagnostics).toEqual([]);
  expect(result.analysis.measurements[0].bounds.width).toBeCloseTo(result.expectedWidth, 1);
  expect(result.analysis.measurements[0].bounds.height).toBeCloseTo(510, 1);
});

test("explicit no-wrap tab stops measure left, center, right and decimal positions using loaded font advances", async ({ page }) => {
  const result = await page.evaluate(() => {
    const cases = [
      { align: "left", runs: [{ text: "A\tWIDE", style: { fontSize: 20 } }], expected: 120 + harness.width("WIDE") },
      { align: "center", runs: [{ text: "A\tWIDE", style: { fontSize: 20 } }], expected: 120 + harness.width("WIDE") / 2 },
      { align: "right", runs: [{ text: "A\tWIDE", style: { fontSize: 20 } }], expected: 120 },
      { align: "decimal", runs: [{ text: "A\t12", style: { fontSize: 20 } }, { text: ".34", style: { fontSize: 20 } }], expected: 120 + harness.width(".34") },
      { align: "decimal", runs: [{ text: "A\t12", style: { fontSize: 20 } }, { text: ".34", style: { fontSize: 30 } }], expected: 120 + harness.width(".34", 30) },
      { align: "left", runs: [{ text: "A\t", style: { fontSize: 20 } }, { text: "WIDE", style: { fontSize: 30 } }], expected: 120 + harness.width("WIDE", 30) },
      { align: "left", indent: 36, hanging: 0, runs: [{ text: "A\tWIDE", style: { fontSize: 20 } }], expected: 156 + harness.width("WIDE") },
      { align: "left", indent: 36, hanging: 18, runs: [{ text: "A\tWIDE", style: { fontSize: 20 } }], expected: 156 + harness.width("WIDE") },
      { align: "right", indent: 18, hanging: 18, runs: [{ text: "A\tWIDE", style: { fontSize: 20 } }], expected: 138 },
      { align: "left", indent: 36, hanging: 18, position: 0, runs: [{ text: "A\tB", style: { fontSize: 20 } }], expected: 36 + harness.width("B") },
      { align: "left", runs: [{ text: "A\t中文", style: { fontSize: 20, eastAsianFontFamily: "Pinned CJK" } }], expected: 120 + harness.width("中文", 20, "Pinned CJK") },
      { align: "right", indent: 36, hanging: 18, runs: [{ text: "中文\tA", style: { fontSize: 20, eastAsianFontFamily: "Pinned CJK" } }], expected: 156 },
      { align: "decimal", runs: [{ text: "中文\t12.34", style: { fontSize: 20, eastAsianFontFamily: "Pinned CJK" } }], expected: 120 + harness.width(".34") },
      { align: "decimal", runs: [{ text: "A\t", style: { fontSize: 20 } }, { text: "中文", style: { fontSize: 20, fontFamily: "Pinned CJK" } }, { text: ".34", style: { fontSize: 20 } }], expected: 120 + harness.width(".34") },
      { align: "left", runs: [{ text: "あア\t한글", style: { fontSize: 20, fontFamily: "Pinned International" } }], expected: 120 + harness.width("한글", 20, "Pinned International") },
      { align: "right", indent: 36, hanging: 18, runs: [{ text: "한글\tあア", style: { fontSize: 20, fontFamily: "Pinned International" } }], expected: 156 },
      { align: "decimal", runs: [{ text: "あア\t한글", style: { fontSize: 20, fontFamily: "Pinned International" } }, { text: ".34", style: { fontSize: 20 } }], expected: 120 + harness.width(".34") },
    ];
    const measured = cases.map(item => ({ expected: item.expected, analysis: harness.analyze([{ style: { tabStops: [{ position: item.position ?? 120, align: item.align }], lineSpacing: 1.5, ...(item.indent === undefined ? {} : { indent: item.indent, hanging: item.hanging }) }, runs: item.runs }], { width: 400, wrap: false }) }));
    const multi = harness.analyze([{ style: { tabStops: [{ position: 80, align: "left" }, { position: 180, align: "right" }], lineSpacing: 1.5 }, runs: [{ text: "A\tBC\tZ", style: { fontSize: 20 } }] }], { width: 400, wrap: false });
    const consecutive = harness.analyze([{ style: { tabStops: [{ position: 80, align: "left" }, { position: 180, align: "left" }], lineSpacing: 1.5 }, runs: [{ text: "A\t\tZ", style: { fontSize: 20 } }] }], { width: 400, wrap: false });
    const limit = harness.analyze([{ style: { tabStops: Array.from({ length: 256 }, (_, i) => ({ position: 20 * (i + 1), align: "left" })), lineSpacing: 1.5 }, runs: [{ text: `A${"\t".repeat(256)}B`, style: { fontSize: 20 } }] }], { width: 6000, wrap: false });
    const wrapped = harness.analyze([{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: [{ text: "A\tB", style: { fontSize: 20 } }] }], { width: 400, wrap: true });
    return { measured, multi, consecutive, expectedConsecutive: 180 + harness.width("Z"), limit, expectedLimit: 5120 + harness.width("B"), wrapped };
  });
  for (const item of result.measured) {
    expect(item.analysis.diagnostics).toEqual([]); expect(item.analysis.measurements).toHaveLength(1);
    expect(item.analysis.measurements[0].bounds.width).toBeCloseTo(item.expected, 1);
  }
  expect(result.multi.diagnostics).toEqual([]); expect(result.multi.measurements[0].bounds.width).toBeCloseTo(180, 1);
  expect(result.consecutive.diagnostics).toEqual([]); expect(result.consecutive.measurements[0].bounds.width).toBeCloseTo(result.expectedConsecutive, 1);
  expect(result.limit.diagnostics).toEqual([]); expect(result.limit.measurements[0].bounds.width).toBeCloseTo(result.expectedLimit, 1);
  expect(result.wrapped.diagnostics.some(d => d.code === "text-measurement-failed")).toBe(false); expect(result.wrapped.measurements).toHaveLength(1);
});

test("wrapping frames measure bounded tabbed lines and reject unsupported automatic wrapping", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const stops = [{ position: 120, align: "left" }], content = [{ style: { tabStops: stops, lineSpacing: 1.5 }, runs: [{ text: "A\tWIDE", style: { fontSize: 20 } }] }];
    const width = 120 + harness.width("WIDE");
    const fit = harness.analyze(content, { width: width + 0.5, wrap: true });
    const noWrap = harness.analyze(content, { width: width + 0.5, wrap: false });
    const narrow = harness.analyze(content, { width: width - 0.5, wrap: true });
    const overflowingNoWrap = harness.analyze(content, { width: width - 0.5, wrap: false });
    const laterLongLine = harness.analyze([{ style: { tabStops: stops, lineSpacing: 1.5 }, runs: [{ text: "A\tB\nWWWWWWWWWWWWWW", style: { fontSize: 20 } }] }], { width: 150, wrap: true });
    const normal = { style: { lineSpacing: 1.5 }, runs: [{ text: "WWWWWWWWWWWWWWWWWWWW", style: { fontSize: 20 } }] };
    const tabbed = { style: { tabStops: stops, lineSpacing: 1.5 }, runs: [{ text: "A\tB", style: { fontSize: 20 } }] };
    const mixedParagraphs = harness.analyze([normal, tabbed], { width: 150, wrap: true });
    const normalOnly = harness.analyze([normal], { width: 150, wrap: true });
    const tabbedOnly = harness.analyze([tabbed], { width: 150, wrap: true });
    const list = harness.analyze([{ style: { bullet: { type: "number", startAt: 333 }, indent: 36, hanging: 18, tabStops: stops, lineSpacing: 1.5 }, runs: [{ text: "P\tB", style: { fontSize: 20 } }] }], { width: 150, wrap: true });
    const { createPresentation } = await import("@pptkit/core"), { resolveLayout } = await import("@pptkit/layout");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } }), slide = deck.addSlide({ id: "containers" });
    const box = { x: 0, y: 0, width: 300, height: 90 };
    slide.addElement({ id: "shape", type: "shape", shape: "rect", box, text: { content } });
    slide.addElement({ id: "table", type: "table", box, columns: [300], rows: [{ height: 90, cells: [{ content }] }] });
    const longContent = [{ style: { tabStops: stops }, runs: [{ text: "A\tWWWWWWWWWWWW", style: { fontSize: 20 } }] }];
    slide.addElement({ id: "narrow-table", type: "table", box: { ...box, width: 80 }, columns: [80], rows: [{ height: 90, cells: [{ content: longContent }] }] });
    const containers = resolveLayout(deck, { textMeasurement: harness.adapter, diagnostics: { textOverflow: true } }).analysis;
    return { fit, noWrap, narrow, overflowingNoWrap, laterLongLine, mixedParagraphs, normalOnly, tabbedOnly, list, containers };
  });
  expect(result.fit.diagnostics).toEqual([]); expect(result.fit.measurements[0].bounds).toEqual(result.noWrap.measurements[0].bounds);
  expect(result.narrow.diagnostics).toEqual([]); expect(result.narrow.measurements[0].bounds.height).toBeGreaterThan(result.fit.measurements[0].bounds.height);
  for (const analysis of [result.laterLongLine]) {
    expect(analysis.measurements).toEqual([]); expect(analysis.diagnostics[0].code).toBe("text-measurement-failed");
    expect(analysis.diagnostics[0].message).toContain("automatic wrapping");
  }
  expect(result.overflowingNoWrap.measurements).toHaveLength(1);
  expect(result.mixedParagraphs.diagnostics).toEqual([]); expect(result.list.diagnostics).toEqual([]);
  expect(result.mixedParagraphs.measurements[0].bounds.height).toBeCloseTo(result.normalOnly.measurements[0].bounds.height + result.tabbedOnly.measurements[0].bounds.height, 1);
  expect(result.containers.measurements.map(m => m.elementId)).toEqual(["shape", "table"]);
  expect(result.containers.diagnostics.filter(d => d.code === "text-measurement-failed").map(d => d.elementId)).toEqual(["narrow-table"]);
});

test("bounded soft breaks preserve tab alignment when fitting and wrap a single word otherwise", async ({ page }) => {
  const results = await page.evaluate(() => {
    const cases = [];
    for (const align of ["left", "center", "right", "decimal"]) {
      const suffix = align === "decimal" ? "123.45" : "WIDE";
      const content = [{ style: { tabStops: [{ position: 120, align }], lineSpacing: 1.5 }, runs: [{ text: `P\t${suffix}`, style: { fontSize: 20 } }] }];
      cases.push({ align, suffixWidth: harness.width(suffix), wide: harness.analyze(content, { width: 300, wrap: true }), narrow: harness.analyze(content, { width: 130, wrap: true }) });
    }
    const unsupported = ["P\tWIDE  WORD", "P\tWWWWWWWWWWWWWW", "P\tA\tB"].map(text => harness.analyze([{ style: { tabStops: [{ position: 120, align: "left" }, { position: 200, align: "left" }] }, runs: [{ text, style: { fontSize: 20 } }] }], { width: 130, wrap: true }));
    const split = harness.analyze([{ style: { tabStops: [{ position: 120, align: "left" }], lineSpacing: 1.5 }, runs: [{ text: "P\tWI", style: { fontSize: 20 } }, { text: "DE\nQ\tWIDE", style: { fontSize: 20 } }] }], { width: 130, wrap: true });
    const one = harness.analyze([{ style: { tabStops: [{ position: 120, align: "left" }], lineSpacing: 1.5 }, runs: [{ text: "P\tWIDE", style: { fontSize: 20 } }] }], { width: 130, wrap: true });
    const indented = harness.analyze([{ style: { indent: 10, tabStops: [{ position: 120, align: "left" }] }, runs: [{ text: "P\tWIDE", style: { fontSize: 20 } }] }], { width: 130, wrap: true });
    const mixedRuns = [{ text: "P\t", style: { fontSize: 16 } }, { text: "WIDE", style: { fontSize: 28 } }];
    const mixed = harness.analyze([{ style: { tabStops: [{ position: 120, align: "left" }], lineSpacing: 1.5 }, runs: mixedRuns }], { width: 130, wrap: true });
    const mixedBreak = harness.analyze([{ style: { lineSpacing: 1.5 }, runs: [{ ...mixedRuns[0], text: "P\n" }, mixedRuns[1]] }], { width: 130, wrap: true });
    const multi = ["left", "center"].map(align => harness.analyze([{ style: { tabStops: [{ position: 120, align }], lineSpacing: 1.5 }, runs: [{ text: "P\tWI", style: { fontSize: 20 } }, { text: "DE WORD", style: { fontSize: 20 } }] }], { width: 130, wrap: true }));
    const wordLimit = count => harness.analyze([{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: [{ text: `P\t${Array(count).fill("A").join(" ")}`, style: { fontSize: 20 } }] }], { width: 130, wrap: true });
    return { cases, unsupported, split, one, indented, mixed, mixedBreak, multi, word: harness.width("WORD"), first: harness.width("WIDE"), combined: harness.width("WIDE WORD"), limit: wordLimit(256), overLimit: wordLimit(257) };
  });
  for (const item of results.cases) {
    expect(item.wide.diagnostics).toEqual([]); expect(item.narrow.diagnostics).toEqual([]);
    if (item.align === "right") expect(item.narrow.measurements[0].bounds).toEqual(item.wide.measurements[0].bounds);
    else {
      expect(item.narrow.measurements[0].bounds.height).toBeGreaterThan(item.wide.measurements[0].bounds.height);
      expect(item.narrow.measurements[0].bounds.width).toBeCloseTo(item.suffixWidth, 1);
    }
  }
  for (const item of results.unsupported) { expect(item.measurements).toEqual([]); expect(item.diagnostics[0].code).toBe("text-measurement-failed"); }
  expect(results.split.diagnostics).toEqual([]); expect(results.split.measurements[0].bounds.height).toBeCloseTo(2 * results.one.measurements[0].bounds.height, 1);
  expect(results.indented.measurements).toEqual([]); expect(results.indented.diagnostics[0].code).toBe("text-measurement-failed");
  expect(results.mixed.diagnostics).toEqual([]); expect(results.mixed.measurements[0].bounds.height).toBeCloseTo(results.mixedBreak.measurements[0].bounds.height, 1);
  for (const item of results.multi) expect(item.diagnostics).toEqual([]);
  for (const item of results.multi) expect(item.measurements[0].bounds.height).toBeGreaterThan(results.cases[0].wide.measurements[0].bounds.height);
  expect(results.multi[0].measurements[0].bounds.width).toBeCloseTo(results.combined, 1);
  expect(results.multi[1].measurements[0].bounds.width).toBeCloseTo(Math.max(120 - results.combined / 2 + results.first, results.word), 1);
  expect(results.limit.measurements).toHaveLength(1); expect(results.limit.diagnostics.some(d => d.code === "text-measurement-failed")).toBe(false);
  expect(results.overLimit.measurements).toEqual([]); expect(results.overLimit.diagnostics[0].message).toContain("256 words");
});

test("strict tab wrapping rejects automatic breaks while retaining fitting and no-wrap measurement", async ({ page }) => {
  const result = await page.evaluate(() => {
    const options = { document, fonts: harness.fonts, tabWrap: "reject" };
    const provider = harness.createBrowserTextMeasurement(options);
    options.tabWrap = "approximate";
    const content = [{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: [{ text: "P\tWIDE WORD", style: { fontSize: 20 } }] }];
    try {
      let invalid;
      try { harness.createBrowserTextMeasurement({ document, fonts: harness.fonts, tabWrap: "other" }); invalid = false; } catch (error) { invalid = error instanceof TypeError; }
      return { invalid, fitting: harness.analyze(content, { provider, width: 400 }), noWrap: harness.analyze(content, { provider, width: 130, wrap: false }), rejected: harness.analyze(content, { provider, width: 130 }), plain: harness.analyze("WIDE WORD WIDE WORD", { provider, width: 80 }), approximate: harness.analyze(content, { width: 130 }) };
    } finally { provider.dispose(); }
  });
  expect(result.invalid).toBe(true);
  for (const analysis of [result.fitting, result.noWrap, result.plain, result.approximate]) { expect(analysis.measurements).toHaveLength(1); expect(analysis.diagnostics.some(d => d.code === "text-measurement-failed")).toBe(false); }
  expect(result.rejected.measurements).toEqual([]); expect(result.rejected.diagnostics[0].message).toContain("selected tabWrap policy");
});

test("PowerPoint tab profile models bounded breaks and rejects unobserved profiles", async ({ page }) => {
  const result = await page.evaluate(() => {
    const provider = harness.createBrowserTextMeasurement({ document, fonts: harness.fonts, tabWrap: "powerpoint" });
    const content = align => [{ style: { tabStops: [{ position: 120, align }], lineSpacing: 1.5 }, runs: [{ text: "P\tWIDE WORD", style: { fontSize: 20 } }] }];
    const splitLines = [{ style: { tabStops: [{ position: 120, align: "left" }], lineSpacing: 1.5 }, runs: [{ text: "P\tWI", style: { fontSize: 20 } }, { text: "DE WORD\nP\tWIDE WORD", style: { fontSize: 20 } }] }];
    try { return { widths: [130, 135, 150, 180].map(width => harness.analyze(content("left"), { provider, width })), center: harness.analyze(content("center"), { provider, width: 100 }), decimal: harness.analyze([{ style: { tabStops: [{ position: 120, align: "decimal" }], lineSpacing: 1.5 }, runs: [{ text: "P\t123.45 WORD", style: { fontSize: 20 } }] }], { provider, width: 130 }), right: harness.analyze([{ style: { tabStops: [{ position: 200, align: "right" }], lineSpacing: 1.5 }, runs: [{ text: "P\tWIDE WORD", style: { fontSize: 20 } }] }], { provider, width: 100 }), unusual: harness.analyze([{ style: { tabStops: [{ position: 120, align: "left" }] }, runs: [{ text: "P\tWIDE 2 WORD", style: { fontSize: 20 } }] }], { provider, width: 135 }), splitLines: harness.analyze(splitLines, { provider, width: 135 }), expectedWidths: [harness.width("WIDE WORD"), 120 + harness.width("W"), 120 + harness.width("WI"), 120 + harness.width("WIDE")] }; } finally { provider.dispose(); }
  });
  for (const analysis of result.widths) { expect(analysis.diagnostics).toEqual([]); expect(analysis.measurements).toHaveLength(1); }
  for (const analysis of result.widths.slice(0, 3)) expect(analysis.measurements[0].bounds.height).toBeCloseTo(90, 1);
  expect(result.widths[3].measurements[0].bounds.height).toBeCloseTo(60, 1);
  for (const [index, analysis] of result.widths.entries()) expect(analysis.measurements[0].bounds.width).toBeCloseTo(result.expectedWidths[index], 1);
  expect(result.splitLines.diagnostics).toEqual([]); expect(result.splitLines.measurements[0].bounds.height).toBeCloseTo(180, 1);
  for (const [analysis, width] of [[result.center, 100], [result.decimal, 130]]) { expect(analysis.diagnostics).toEqual([]); expect(analysis.measurements[0].bounds.width).toBeCloseTo(width, 1); expect(analysis.measurements[0].bounds.height).toBeCloseTo(60, 1); }
  expect(result.right.diagnostics).toEqual([]); expect(result.right.measurements[0].bounds.width).toBeCloseTo(100, 1);
  expect(result.unusual.measurements).toEqual([]); expect(result.unusual.diagnostics[0].message).toContain("PowerPoint profile");
});

test("overlapping tab alignments retain policy-specific handling without partial measurements", async ({ page }) => {
  const results = await page.evaluate(() => {
    const policies = ["approximate", "powerpoint", "reject"];
    const cases = [{ position: 50, align: "center", text: "P\tWIDE WORD" }, { position: 90, align: "right", text: "P\tWIDE WORD" }, { position: 30, align: "decimal", text: "P\t123.45 WORD" }];
    const measurements = policies.map(tabWrap => {
      const provider = harness.createBrowserTextMeasurement({ document, fonts: harness.fonts, tabWrap });
      try { return cases.map(({ position, align, text }) => harness.analyze([{ style: { tabStops: [{ position, align }], lineSpacing: 1.5 }, runs: [{ text, style: { fontSize: 20 } }] }], { provider, width: 150, wrap: true })); }
      finally { provider.dispose(); }
    });
    const provider = harness.createBrowserTextMeasurement({ document, fonts: harness.fonts, tabWrap: "powerpoint" });
    try {
      const unsupported = harness.analyze([{ style: { tabStops: [{ position: 90, align: "right" }] }, runs: [{ text: "P\tWIDE 2 WORD", style: { fontSize: 20 } }] }], { provider, width: 150 });
      const noWrap = harness.analyze([{ style: { tabStops: [{ position: 90, align: "right" }] }, runs: [{ text: "P\tWIDE WORD", style: { fontSize: 20 } }] }], { provider, width: 150, wrap: false });
      return { measurements, unsupported, noWrap };
    } finally { provider.dispose(); }
  });
  for (const analysis of [...results.measurements[0], ...results.measurements[2], results.noWrap]) { expect(analysis.measurements).toEqual([]); expect(analysis.diagnostics[0].code).toBe("text-measurement-failed"); expect(analysis.diagnostics[0].message).toContain("overlaps its preceding text"); }
  for (const analysis of results.measurements[1]) { expect(analysis.diagnostics).toEqual([]); expect(analysis.measurements).toHaveLength(1); }
  expect(results.unsupported.measurements).toEqual([]); expect(results.unsupported.diagnostics[0].message).toContain("PowerPoint profile");
});

test("PowerPoint tab profile uses the longest fitting initial phrase and bounds its work", async ({ page }) => {
  const result = await page.evaluate(() => {
    const provider = harness.createBrowserTextMeasurement({ document, fonts: harness.fonts, tabWrap: "powerpoint" });
    const content = text => [{ style: { tabStops: [{ position: 120, align: "center" }], lineSpacing: 1.5 }, runs: [{ text, style: { fontSize: 20 } }] }];
    try { return { phrase: harness.analyze(content("P\tI WIDE WORD MORE"), { provider, width: 200 }), expected: 120 + harness.width("I WIDE WORD") / 2, limit: harness.analyze(content(`P\t${Array(256).fill("A").join(" ")}`), { provider, width: 150 }), tooMany: harness.analyze(content(`P\t${Array(257).fill("A").join(" ")}`), { provider, width: 150 }) }; } finally { provider.dispose(); }
  });
  expect(result.phrase.diagnostics).toEqual([]); expect(result.phrase.measurements[0].bounds.width).toBeCloseTo(result.expected, 1); expect(result.phrase.measurements[0].bounds.height).toBeCloseTo(60, 1);
  expect(result.limit.measurements).toHaveLength(1); expect(result.limit.diagnostics.some(d => d.code === "text-measurement-failed")).toBe(false);
  expect(result.tooMany.measurements).toEqual([]); expect(result.tooMany.diagnostics[0].message).toContain("256 words");
});

test("bounded list/tab profile places markers and stops for symbols and Arabic numbers", async ({ page }) => {
  const result = await page.evaluate(() => {
    const bullets = [{ type: "bullet", character: "•" }, { type: "number", startAt: 3 }, { type: "number", startAt: 333 }];
    const cases = [];
    for (const bullet of bullets) for (const [indent, hanging] of [[36, 18], [27, 27], [54, 18]]) for (const align of ["left", "center", "right", "decimal"]) {
      const family = bullet.type === "bullet" ? "Pinned CJK" : "Pinned Mono", text = align === "decimal" ? "12.34" : "WIDE";
      const width = harness.width(text, 20, family), suffix = harness.width(".34", 20, family);
      const expected = align === "left" ? 280 + width : align === "center" ? 280 + width / 2 : align === "right" ? 280 : 280 + suffix;
      const analysis = harness.analyze([{ style: { bullet, indent, hanging, lineSpacing: 1.5, tabStops: [{ position: 280, align }] }, runs: [{ text: `P\t${text}`, style: { fontSize: 20, fontFamily: family } }] }], { width: 600, wrap: false });
      cases.push({ bullet, indent, hanging, align, expected, analysis });
    }
    const unsupported = [
      { bullet: { type: "bullet", character: "*" }, runs: [{ text: "P\tB" }] },
      { bullet: { type: "number", style: "alphaLowerPeriod", startAt: 3 }, runs: [{ text: "P\tB" }] },
      { bullet: { type: "number", style: "arabicParen", startAt: 3 }, runs: [{ text: "P\tB" }] },
      { bullet: { type: "number", startAt: 3 }, runs: [{ text: "" }, { text: "P\tB" }] },
    ].map(item => harness.analyze([{ style: { bullet: item.bullet, indent: 36, hanging: 18, tabStops: [{ position: 280, align: "left" }] }, runs: item.runs.map(run => ({ ...run, style: { fontSize: 20 } })) }], { wrap: false }));
    const overlap = harness.analyze([{ style: { bullet: { type: "number", startAt: 333 }, indent: 36, hanging: 18, tabStops: [{ position: 50, align: "left" }] }, runs: [{ text: "P\tB", style: { fontSize: 20 } }] }], { wrap: false });
    const missingMarkerFont = harness.analyze([{ style: { bullet: { type: "bullet", character: "•" }, indent: 36, hanging: 18, tabStops: [{ position: 120, align: "left" }] }, runs: [{ text: "P\tB", style: { fontSize: 20, fontFamily: "Pinned Mono" } }] }], { wrap: false });
    const hardBreak = harness.analyze([{ style: { bullet: { type: "number", startAt: 3 }, indent: 36, hanging: 18, lineSpacing: 1.5, tabStops: [{ position: 120, align: "left" }] }, runs: [{ text: "P\tB\nP\tB", style: { fontSize: 20 } }] }], { width: 400, wrap: false });
    const boundaries = [];
    for (const [bullet, marker, family] of [[{ type: "bullet", character: "•" }, "•", "Pinned CJK"], [{ type: "number", startAt: 3 }, "3.", "Pinned Mono"], [{ type: "number", startAt: 333 }, "333.", "Pinned Mono"]]) {
      const prefixEnd = 18 + Math.max(18, harness.width(marker, 20, family)) + harness.width("P", 20, family);
      for (const delta of [-0.25, 0.25]) {
        const firstStop = prefixEnd + delta, expected = (delta < 0 ? 280 : firstStop) + harness.width("B", 20, family);
        const analysis = harness.analyze([{ style: { bullet, indent: 36, hanging: 18, lineSpacing: 1.5, tabStops: [{ position: firstStop, align: "left" }, { position: 280, align: "left" }] }, runs: [{ text: "P\tB", style: { fontSize: 20, fontFamily: family } }] }], { width: 600, wrap: false });
        boundaries.push({ bullet, delta, expected, analysis });
      }
    }
    return { cases, unsupported, overlap, missingMarkerFont, hardBreak, expectedHardBreak: 120 + harness.width("B"), boundaries };
  });
  expect(result.cases).toHaveLength(36);
  for (const item of result.cases) {
    expect(item.analysis.diagnostics, JSON.stringify(item)).toEqual([]);
    expect(item.analysis.measurements[0].bounds.width).toBeCloseTo(item.expected, 1);
  }
  for (const analysis of [...result.unsupported, result.overlap, result.missingMarkerFont]) {
    expect(analysis.measurements).toEqual([]); expect(analysis.diagnostics.some(d => d.code === "text-measurement-failed")).toBe(true);
  }
  expect(result.hardBreak.diagnostics).toEqual([]); expect(result.hardBreak.measurements[0].bounds.height).toBeCloseTo(60, 1);
  expect(result.hardBreak.measurements[0].bounds.width).toBeCloseTo(result.expectedHardBreak, 1);
  for (const item of result.boundaries) {
    expect(item.analysis.diagnostics, JSON.stringify(item)).toEqual([]);
    expect(item.analysis.measurements[0].bounds.width).toBeCloseTo(item.expected, 1);
  }
});

test("native SVG draws bounded no-wrap tab chunks at explicit anchors", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const cases = [];
    for (const align of ["left", "center", "right", "decimal"]) {
      const suffix = align === "decimal" ? "12.34" : "WIDE";
      const p = createPresentation({ size: { width: 400, height: 200 }, theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
      p.addSlide().addElement({ type: "text", box: { x: 30, y: 10, width: 300, height: 150 }, frame: { margin: { left: 8, right: 0, top: 0, bottom: 0 }, wrap: false }, content: [{ style: { lineSpacing: 1.5, tabStops: [{ position: 120, align }] }, runs: [{ text: `P\t${suffix}\nQ\t${suffix}`, style: { fontSize: 20 } }] }] });
      const rendered = await renderPresentationToSvg(p), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
      try {
        const parts = [...holder.querySelectorAll("text[data-pptkit-tab-part]")].map(node => ({ part: node.dataset.pptkitTabPart, text: node.textContent, y: Number(node.getAttribute("y")), start: node.getStartPositionOfChar(0).x, end: node.getEndPositionOfChar(node.getNumberOfChars() - 1).x }));
        cases.push({ align, parts, warnings: rendered.warnings.map(warning => warning.code) });
      } finally { holder.remove(); }
    }
    return cases;
  });
  for (const item of result) {
    expect(item.warnings).toContain("text-tab-layout-unsupported");
    const first = item.parts.filter(part => part.y === 26), second = item.parts.filter(part => part.y === 56);
    expect(first[0]).toMatchObject({ part: "prefix", text: "P\t" }); expect(second[0]).toMatchObject({ part: "prefix", text: "Q\t" });
    if (item.align === "decimal") {
      expect(first.map(part => part.text)).toEqual(["P\t", "12", ".34\n"]);
      expect(first[1].end).toBeCloseTo(158, 1); expect(first[2].start).toBeCloseTo(158, 1);
    } else {
      expect(first.map(part => part.text)).toEqual(["P\t", "WIDE\n"]);
      const value = item.align === "left" ? first[1].start : item.align === "right" ? first[1].end : (first[1].start + first[1].end) / 2;
      expect(value).toBeCloseTo(158, 1);
    }
  }
});

test("native SVG tab chunks retain mixed-run styles and split CRLF boundaries", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const deck = createPresentation({ size: { width: 400, height: 240 }, theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } }), slide = deck.addSlide();
    slide.addElement({ id: "styled", type: "text", box: { x: 30, y: 10, width: 300, height: 120 }, frame: { margin: 0, wrap: false }, content: [{ style: { lineSpacing: 1.5, tabStops: [{ position: 120, align: "center" }] }, runs: [{ text: "P\tWI", style: { fontSize: 20, color: "AA0000" } }, { text: "DE\r", style: { fontSize: 20, color: "0000AA", underline: true } }, { text: "\nQ\tB", style: { fontSize: 20, color: "008800" } }] }] });
    slide.addElement({ id: "decimal", type: "text", box: { x: 30, y: 140, width: 300, height: 80 }, frame: { margin: 0, wrap: false }, content: [{ style: { tabStops: [{ position: 120, align: "decimal" }] }, runs: [{ text: "P", style: { fontSize: 20 } }, { text: "\t12", style: { fontSize: 20, color: "AA0000" } }, { text: ".34", style: { fontSize: 20, color: "0000AA" } }] }] });
    const before = JSON.stringify(deck.slides[0].elements), rendered = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
    try {
      const chunks = id => [...holder.querySelectorAll(`[data-pptkit-element-id="${id}"] text[data-pptkit-tab-part]`)].map(node => ({ part: node.dataset.pptkitTabPart, text: node.textContent, y: Number(node.getAttribute("y")), start: node.getStartPositionOfChar(0).x, end: node.getEndPositionOfChar(node.getNumberOfChars() - 1).x, spans: [...node.querySelectorAll("tspan[font-family]")].map(span => ({ text: span.textContent, fill: span.getAttribute("fill"), decoration: span.getAttribute("text-decoration") })) }));
      return { styled: chunks("styled"), decimal: chunks("decimal"), warnings: rendered.warnings.map(w => w.code), html: holder.querySelectorAll("foreignObject").length, unchanged: before === JSON.stringify(deck.slides[0].elements) };
    } finally { holder.remove(); }
  });
  expect(result.html).toBe(0); expect(result.unchanged).toBe(true); expect(result.warnings.filter(code => code === "text-tab-layout-unsupported")).toHaveLength(2);
  expect(result.styled.map(chunk => chunk.text)).toEqual(["P\t", "WIDE\r\n", "Q\t", "B"]);
  expect(result.styled.map(chunk => chunk.y)).toEqual([26, 26, 56, 56]);
  expect((result.styled[1].start + result.styled[1].end) / 2).toBeCloseTo(150, 1);
  expect(result.styled[1].spans).toEqual([{ text: "WI", fill: "#AA0000", decoration: "none" }, { text: "DE", fill: "#0000AA", decoration: "underline" }]);
  expect(result.decimal.map(chunk => chunk.text)).toEqual(["P\t", "12", ".34"]);
  expect(result.decimal[1].end).toBeCloseTo(150, 1); expect(result.decimal[2].start).toBeCloseTo(150, 1);
});

test("native tab chunks preserve hard breaks in source order for text and shapes", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const deck = createPresentation(), slide = deck.addSlide();
    for (const kind of ["text", "shape"]) for (const align of ["left", "center", "right", "decimal"]) {
      const source = align === "decimal" ? "P\t12.34\r\nQ\t56.78\rR\t90.12\nS\t34.56" : "P\t1234\r\nQ\t5678\rR\t9012\nS\t3456";
      const content = [{ style: { tabStops: [{ position: 120, align }] }, runs: [{ text: source, style: { fontSize: 20 } }] }];
      const box = { x: 0, y: 0, width: 300, height: 200 }, frame = { margin: 0, wrap: false };
      slide.addElement(kind === "text" ? { id: `${kind}-${align}`, type: kind, box, frame, content } : { id: `${kind}-${align}`, type: kind, shape: "rect", box, text: { frame, content } });
    }
    const holder = document.createElement("div"); holder.innerHTML = (await renderPresentationToSvg(deck)).slides[0].svg; document.body.append(holder);
    try { return [...holder.querySelectorAll("[data-pptkit-element-id]")].map(node => ({ id: node.getAttribute("data-pptkit-element-id"), text: node.textContent, fallback: !!node.querySelector("foreignObject") })); }
    finally { holder.remove(); }
  });
  expect(result).toHaveLength(8);
  for (const item of result) {
    expect(item.fallback, item.id).toBe(false);
    expect(item.text, item.id).toBe(item.id.endsWith("decimal") ? "P\t12.34\r\nQ\t56.78\rR\t90.12\nS\t34.56" : "P\t1234\r\nQ\t5678\rR\t9012\nS\t3456");
  }
});

test("native SVG anchors multiword no-wrap chunks across runs and shape owners", async ({ page }) => {
  const results = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const results = [];
    for (const type of ["text", "shape"]) for (const align of ["left", "center", "right", "decimal"]) {
      const deck = createPresentation({ size: { width: 600, height: 200 }, theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } }), slide = deck.addSlide();
      const box = { x: 20, y: 10, width: 500, height: 120 }, frame = { margin: 0, wrap: false };
      const suffix = align === "decimal" ? "12.34" : "WIDE WORD";
      const content = [{ style: { tabStops: [{ position: 220, align }] }, runs: [{ text: "FIRST ", style: { fontSize: 20 } }, { text: `WORD\t${suffix.slice(0, 2)}`, style: { fontSize: 20, bold: true } }, { text: suffix.slice(2), style: { fontSize: 20, underline: true } }] }];
      slide.addElement(type === "text" ? { id: "sample", type, box, frame, content } : { id: "sample", type, shape: "rect", box, text: { frame, content } });
      const before = JSON.stringify(slide.elements), rendered = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
      try {
        const root = holder.querySelector('[data-pptkit-element-id="sample"]'), nodes = [...root.querySelectorAll("text[data-pptkit-tab-part]")];
        results.push({ type, align, unchanged: before === JSON.stringify(slide.elements), fallback: !!root.querySelector("foreignObject"), warnings: rendered.warnings.map(w => w.code), parts: nodes.map(node => ({ text: node.textContent, start: node.getStartPositionOfChar(0).x, end: node.getEndPositionOfChar(node.getNumberOfChars() - 1).x })) });
      } finally { holder.remove(); }
    }
    return results;
  });
  for (const item of results) {
    expect(item.unchanged).toBe(true); expect(item.fallback).toBe(false); expect(item.warnings).toContain("text-tab-layout-unsupported");
    expect(item.parts[0].text).toBe("FIRST WORD\t");
    if (item.align === "decimal") {
      expect(item.parts.map(part => part.text)).toEqual(["FIRST WORD\t", "12", ".34"]);
      expect(item.parts[1].end).toBeCloseTo(240, 1); expect(item.parts[2].start).toBeCloseTo(240, 1);
    } else {
      const suffix = item.parts[1]; expect(suffix.text).toBe("WIDE WORD");
      expect(item.align === "left" ? suffix.start : item.align === "right" ? suffix.end : (suffix.start + suffix.end) / 2).toBeCloseTo(240, 1);
    }
  }
});

test("shape text uses bounded SVG tab anchors inside the shape transform", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const deck = createPresentation({ size: { width: 600, height: 400 }, theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } }), slide = deck.addSlide();
    const shape = (id, y, content, transform = {}) => ({ id, type: "shape", shape: "rect", box: { x: 40, y, width: 300, height: 90 }, transform, style: { fill: { type: "solid", color: "DDEEFF" } }, text: { frame: { margin: { left: 8, right: 0, top: 0, bottom: 0 }, wrap: false }, content } });
    const paragraph = runs => [{ style: { lineSpacing: 1.5, tabStops: [{ position: 120, align: "center" }] }, runs }];
    slide.addElement(shape("tab-shape", 30, paragraph([{ text: "P\t", style: { fontSize: 20 } }, { text: "WIDE", style: { fontSize: 20, underline: true } }]), { rotation: 20 }));
    slide.addElement(shape("plain-shape", 150, "Ordinary"));
    slide.addElement(shape("linked-shape", 270, paragraph([{ text: "P\t", style: { fontSize: 20 } }, { text: "WIDE", style: { fontSize: 20 }, action: { type: "url", url: "https://example.com" } }])));
    const before = JSON.stringify(deck.slides[0].elements), rendered = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
    try {
      const root = holder.querySelector('[data-pptkit-element-id="tab-shape"]'), rect = root.querySelector("rect"), suffix = root.querySelector('text[data-pptkit-tab-part="suffix"]');
      const matrix = node => { const m = node.getCTM(); return [m.a, m.b, m.c, m.d, m.e, m.f]; };
      return { warnings: rendered.warnings.map(w => [w.elementId, w.code]), unchanged: before === JSON.stringify(deck.slides[0].elements), part: suffix.textContent, anchor: (suffix.getStartPositionOfChar(0).x + suffix.getEndPositionOfChar(suffix.getNumberOfChars() - 1).x) / 2, rectangleMatrix: matrix(rect), textMatrix: matrix(suffix), native: root.querySelectorAll("foreignObject").length, fallback: ["plain-shape", "linked-shape"].map(id => !!holder.querySelector(`[data-pptkit-element-id="${id}"] foreignObject`)), linked: !!holder.querySelector('[data-pptkit-element-id="linked-shape"] a[href="https://example.com"]') };
    } finally { holder.remove(); }
  });
  expect(result.unchanged).toBe(true); expect(result.part).toBe("WIDE"); expect(result.anchor).toBeCloseTo(168, 1);
  expect(result.rectangleMatrix).toEqual(result.textMatrix); expect(result.native).toBe(0); expect(result.fallback).toEqual([true, true]); expect(result.linked).toBe(true);
  expect(result.warnings.filter(([id, code]) => code === "text-tab-layout-unsupported").map(([id]) => id).sort()).toEqual(["linked-shape", "tab-shape"]);
});

test("shape and text boxes share tab geometry for all supported alignments", async ({ page }) => {
  const cases = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const results = [];
    for (const align of ["left", "center", "right", "decimal"]) {
      for (const verticalAlign of ["top", "middle", "bottom"]) {
        const deck = createPresentation({ size: { width: 600, height: 300 }, theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
        const slide = deck.addSlide(), box = { x: 31, y: 17, width: 300, height: 120 };
        const frame = { margin: { left: 7, right: 5, top: 9, bottom: 11 }, wrap: false, verticalAlign };
        const content = [{ style: { tabStops: [{ position: 120, align }] }, runs: [{ text: "P\t", style: { fontSize: 20 } }, { text: align === "decimal" ? "12.34" : "WIDE", style: { fontSize: 20, bold: true } }] }];
        slide.addElement({ id: "standalone", type: "text", box, frame, content });
        slide.addElement({ id: "shape", type: "shape", shape: "rect", box, text: { frame, content } });
        const rendered = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
        try {
          const geometry = id => [...holder.querySelectorAll(`[data-pptkit-element-id="${id}"] text[data-pptkit-tab-part]`)].map(node => ({ text: node.textContent, x: node.getAttribute("x"), y: node.getAttribute("y"), anchor: node.getAttribute("text-anchor"), start: node.getStartPositionOfChar(0).x, end: node.getEndPositionOfChar(node.getNumberOfChars() - 1).x }));
          results.push({ align, verticalAlign, text: geometry("standalone"), shape: geometry("shape"), fallback: !!holder.querySelector('[data-pptkit-element-id="shape"] foreignObject') });
        } finally { holder.remove(); }
      }
    }
    return results;
  });
  for (const item of cases) {
    expect(item.shape.length, `${item.align}/${item.verticalAlign}`).toBeGreaterThan(0);
    expect(item.shape).toEqual(item.text);
    expect(item.fallback).toBe(false);
  }
});

test("shape tab paragraphs coexist with ordinary paragraphs while unsupported siblings fall back", async ({ page }) => {
  const results = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const tab = { style: { tabStops: [{ position: 120, align: "right" }] }, runs: [{ text: "P\t", style: { fontSize: 20 } }, { text: "WIDE", style: { fontSize: 20, bold: true } }] };
    const plain = { runs: [{ text: "Ordinary paragraph", style: { fontSize: 16 } }] };
    const cases = [
      { id: "before", content: [plain, tab] },
      { id: "after", content: [tab, plain] },
      { id: "linked", content: [tab, { runs: [{ text: "Link", action: { type: "url", url: "https://example.com" } }] }] },
      { id: "mixed", content: [tab, { runs: [{ text: "mixed" }, { text: " text" }] }] },
      { id: "bad-tab", content: [tab, { runs: [{ text: "P\tA\tB" }], style: { tabStops: [{ position: 120, align: "left" }] } }] },
      { id: "plain", content: [plain] },
    ];
    const deck = createPresentation({ size: { width: 600, height: 600 }, theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } }), slide = deck.addSlide();
    for (const [index, item] of cases.entries()) slide.addElement({ id: item.id, type: "shape", shape: "rect", box: { x: 20, y: index * 90, width: 300, height: 80 }, text: { frame: { margin: 0, wrap: false }, content: item.content } });
    const before = JSON.stringify(slide.elements), rendered = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
    try {
      return { unchanged: before === JSON.stringify(slide.elements), cases: cases.map(({ id }) => {
        const root = holder.querySelector(`[data-pptkit-element-id="${id}"]`), suffix = root.querySelector('text[data-pptkit-tab-part="suffix"]');
        return { id, fallback: !!root.querySelector("foreignObject"), text: root.textContent, suffixEnd: suffix && suffix.getEndPositionOfChar(suffix.getNumberOfChars() - 1).x, linked: !!root.querySelector('a[href="https://example.com"]') };
      }) };
    } finally { holder.remove(); }
  });
  expect(results.unchanged).toBe(true);
  for (const id of ["before", "after"]) {
    const result = results.cases.find(item => item.id === id);
    expect(result.fallback).toBe(false); expect(result.text).toContain("Ordinary paragraph"); expect(result.text).toContain("WIDE"); expect(result.suffixEnd).toBeCloseTo(140, 1);
  }
  for (const id of ["linked", "mixed", "bad-tab", "plain"]) expect(results.cases.find(item => item.id === id).fallback, id).toBe(true);
  expect(results.cases.find(item => item.id === "linked").linked).toBe(true);
});

test("SVG tab anchoring leaves unsupported text paths diagnosed and unchanged", async ({ page }) => {
  const results = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const cases = [
      { label: "wrap", wrap: true, text: "P\tWIDE", stops: [{ position: 120, align: "left" }] },
      { label: "two-tabs", wrap: false, text: "P\tA\tB", stops: [{ position: 120, align: "left" }] },
      { label: "trailing-space", wrap: false, text: "P\tWIDE ", stops: [{ position: 120, align: "left" }] },
      { label: "leading-space", wrap: false, text: " P\tWIDE", stops: [{ position: 120, align: "left" }] },
      { label: "repeated-space", wrap: false, text: "P\tWIDE  WORD", stops: [{ position: 120, align: "left" }] },
      { label: "decimal-phrase", wrap: false, text: "P\t12.34 WORD", stops: [{ position: 120, align: "decimal" }] },
      { label: "no-stop", wrap: false, text: "P\tWIDE", stops: [] },
      { label: "decimal-no-dot", wrap: false, text: "P\t123", stops: [{ position: 120, align: "decimal" }] },
      { label: "rtl", wrap: false, text: "P\tאב", stops: [{ position: 120, align: "left" }] },
      { label: "combining", wrap: false, text: "P\tか\u3099", stops: [{ position: 120, align: "left" }] },
      { label: "emoji", wrap: false, text: "P\t😀", stops: [{ position: 120, align: "left" }] },
      { label: "mixed-action", wrap: false, text: "P\tWIDE", stops: [{ position: 120, align: "left" }], mixed: true },
      { label: "mixed-rtl", wrap: false, text: "P\tWIDE\nא", stops: [{ position: 120, align: "left" }], mixedNonAscii: true },
    ];
    return Promise.all(cases.map(async item => {
      const p = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
      const runs = item.mixed ? [{ text: "P\t", style: { fontSize: 20 } }, { text: "WIDE", style: { fontSize: 20 }, action: { type: "url", url: "https://example.com" } }] : item.mixedNonAscii ? [{ text: "P\tWIDE\n", style: { fontSize: 20 } }, { text: "א", style: { fontSize: 20 } }] : [{ text: item.text, style: { fontSize: 20 } }];
      p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 300, height: 100 }, frame: { margin: 0, wrap: item.wrap }, content: [{ style: { tabStops: item.stops }, runs }] });
      p.slides[0].addElement({ id: "shape-fallback", type: "shape", shape: "rect", box: { x: 0, y: 100, width: 300, height: 100 }, text: { frame: { margin: 0, wrap: item.wrap }, content: [{ style: { tabStops: item.stops }, runs }] } });
      const before = JSON.stringify(p.slides[0].elements), rendered = await renderPresentationToSvg(p);
      const holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg;
      return { label: item.label, count: holder.querySelectorAll("[data-pptkit-tab-part]").length, warnings: rendered.warnings.map(warning => warning.code), unchanged: before === JSON.stringify(p.slides[0].elements), html: !!holder.querySelector("foreignObject"), shapeFallback: !!holder.querySelector('[data-pptkit-element-id="shape-fallback"] foreignObject') };
    }));
  });
  for (const item of results) { expect(item.count, item.label).toBe(0); expect(item.warnings).toContain("text-tab-layout-unsupported"); expect(item.unchanged).toBe(true); expect(item.shapeFallback, item.label).toBe(true); }
  expect(results.find(item => item.label === "mixed-action").html).toBe(true);
  expect(results.find(item => item.label === "mixed-rtl").html).toBe(true);
});

test("native SVG honors no-wrap and preserves source spaces like the mixed-run path", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const source = "  M  M  ";
    const cases = [];
    for (const mode of ["native", "mixed", "wrapped"]) {
      const p = createPresentation({ size: { width: 300, height: 300 }, theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
      const runs = mode === "mixed" ? [{ text: source.slice(0, 3), style: { fontSize: 20 } }, { text: source.slice(3), style: { fontSize: 20 } }] : [{ text: source, style: { fontSize: 20 } }];
      p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 35, height: 300 }, frame: { margin: 0, wrap: mode === "wrapped" }, content: [{ runs }] });
      const before = JSON.stringify(p.slides[0].elements), rendered = await renderPresentationToSvg(p);
      if (before !== JSON.stringify(p.slides[0].elements)) throw Error("Renderer changed author text");
      const holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
      try {
        const lines = [...holder.querySelectorAll("[data-pptkit-text-line]")], spans = [...holder.querySelectorAll("foreignObject span[lang]")];
        const advance = lines[0] ? lines[0].getComputedTextLength() : spans.reduce((sum, span) => { const r = document.createRange(); r.selectNodeContents(span); return sum + r.getBoundingClientRect().width; }, 0);
        cases.push({ mode, lineCount: lines.length, text: lines.length ? lines.map(line => line.textContent).join("") : spans.map(span => span.textContent).join(""), advance });
      } finally { holder.remove(); }
    }
    return { cases, expectedAdvance: harness.width(source) };
  });
  expect(result.cases[0].lineCount).toBe(1);
  expect(result.cases[0].text).toBe("  M  M  ");
  expect(result.cases[0].advance).toBeCloseTo(result.expectedAdvance, 1);
  expect(result.cases[1].lineCount).toBe(0);
  expect(result.cases[1].text).toBe("  M  M  ");
  expect(result.cases[1].advance).toBeCloseTo(result.expectedAdvance, 1);
  expect(result.cases[2].lineCount).toBeGreaterThan(1);
  expect(result.cases[2].text).toBe("  M  M  ");
});

test("native SVG preserves repeated CR/LF breaks and the final empty line's vertical space", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const p = createPresentation();
    p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 200, height: 300 }, frame: { margin: 0, wrap: false }, content: [{ style: { lineSpacing: 1.5 }, runs: [{ text: "M\r\n\rM\n\n", style: { fontFamily: "Pinned Mono", fontSize: 20 } }] }, { runs: [{ text: "After", style: { fontFamily: "Pinned Mono", fontSize: 20 } }] }] });
    const holder = document.createElement("div"); holder.innerHTML = (await renderPresentationToSvg(p)).slides[0].svg; document.body.append(holder);
    try { return { text: holder.querySelector('[data-pptkit-element-id]').textContent, lines: [...holder.querySelectorAll("text[data-pptkit-text-line]")].map(line => ({ text: [...line.querySelectorAll("tspan[font-family]")].map(span => span.textContent).join(""), y: Number(line.getAttribute("y")) })) }; }
    finally { holder.remove(); }
  });
  expect(result.text).toBe("M\r\n\rM\n\nAfter");
  expect(result.lines.map(line => line.text)).toEqual(["M", "", "M", "", "", "After"]);
  expect(result.lines.map(line => line.y)).toEqual([16, 46, 76, 106, 136, 166]);
});

test("native forced wrapping keeps combining, ZWJ and flag graphemes intact across selected font fragments", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const p = createPresentation();
    p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 1, height: 500 }, frame: { margin: 0 }, content: [{ runs: [{ text: "a\u0301👩🏽‍💻🇨🇳", style: { fontFamily: "Pinned Mono", fontSize: 20, scriptFontFamilies: { Zyyy: "Pinned CJK", Latn: "Pinned Mono" } } }] }] });
    const holder = document.createElement("div"); holder.innerHTML = (await renderPresentationToSvg(p)).slides[0].svg; document.body.append(holder);
    try { return [...holder.querySelectorAll("[data-pptkit-text-line]")].map(line => line.textContent); }
    finally { holder.remove(); }
  });
  // This checks source cluster integrity, not emoji font coverage or shaping.
  expect(result).toEqual(["a\u0301", "👩🏽‍💻", "🇨🇳"]);
});

test("missing Intl grapheme support uses escaped XHTML text without failing module import", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const descriptor = Object.getOwnPropertyDescriptor(Intl, "Segmenter");
    Object.defineProperty(Intl, "Segmenter", { configurable: true, value: undefined });
    try {
      const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
      const p = createPresentation(); p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 200, height: 100 }, frame: { wrap: false }, content: "  <safe> a\u0301  " });
      const holder = document.createElement("div"); holder.innerHTML = (await renderPresentationToSvg(p)).slides[0].svg; document.body.append(holder);
      try { return { native: holder.querySelectorAll("text[data-pptkit-text-line]").length, fallback: holder.querySelectorAll("foreignObject").length, text: holder.querySelector("span[lang]").textContent, injected: holder.querySelectorAll("safe").length }; }
      finally { holder.remove(); }
    } finally { Object.defineProperty(Intl, "Segmenter", descriptor); }
  });
  expect(result).toEqual({ native: 0, fallback: 1, text: "  <safe> a\u0301  ", injected: 0 });
});

test("native text and chart labels quote literal font families containing numeric words, commas and quotes", async ({ page }) => {
  const result = await page.evaluate(async ({ fontBytes }) => {
    const alias = 'Pinned 123, "Probe"';
    const face = await new FontFace(alias, `url(data:font/woff2;base64,${fontBytes})`).load(); document.fonts.add(face);
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const p = createPresentation({ theme: { fonts: { body: alias, heading: alias } } });
    const slide = p.addSlide();
    slide.addElement({ id: "family", type: "text", box: { x: 0, y: 0, width: 400, height: 70 }, frame: { margin: 0, wrap: false }, content: [{ runs: [{ text: "MMMMiiii", style: { fontFamily: alias, fontSize: 20 } }] }] });
    slide.addElement({ id: "chart", type: "chart", chartType: "bar", title: "MMMMiiii", titleStyle: { fontFamily: alias, fontSize: 20 }, categories: ["A"], series: [{ name: "Value", values: [1] }], box: { x: 0, y: 80, width: 400, height: 250 } });
    const holder = document.createElement("div"); holder.innerHTML = (await renderPresentationToSvg(p)).slides[0].svg; document.body.append(holder);
    try {
      const nodes = [holder.querySelector("tspan"), holder.querySelector('[data-chart-role="title"]')];
      const advances = nodes.map(node => node.getComputedTextLength()), families = nodes.map(node => getComputedStyle(node).fontFamily);
      // Simulate the old unquoted attribute with the same correctly registered face.
      for (const node of nodes) node.setAttribute("font-family", alias);
      return { advances, rawAdvances: nodes.map(node => node.getComputedTextLength()), expected: harness.width("MMMMiiii"), registered: face.family, families };
    }
    finally { holder.remove(); document.fonts.delete(face); }
  }, { fontBytes: latin });
  for (const advance of result.advances) expect(advance, JSON.stringify(result)).toBeCloseTo(result.expected, 1);
  for (const advance of result.rawAdvances) expect(Math.abs(advance - result.expected)).toBeGreaterThan(.05);
  expect(result.families[0]).toContain('Pinned 123,');
  expect(result.families[1]).toBe(result.families[0]);
});

test("mixed SVG preserves trailing empty lines and empty paragraphs with the measured line height", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const content = [{ style: { lineSpacing: 1.5 }, runs: [{ text: "M\n", style: { fontFamily: "Pinned Mono", fontSize: 20 } }, { text: "\n", style: { fontFamily: "Pinned Mono", fontSize: 20 } }] }, { style: { lineSpacing: 1.5 }, runs: [{ text: "", style: { fontFamily: "Pinned Mono", fontSize: 20 } }] }];
    const p = createPresentation(); p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 200, height: 200 }, frame: { margin: 0, wrap: false }, content });
    const holder = document.createElement("div"); holder.innerHTML = (await renderPresentationToSvg(p)).slides[0].svg; document.body.append(holder);
    try {
      const paragraphs = [...holder.querySelector("foreignObject > div").children];
      return { heights: paragraphs.map(paragraph => paragraph.getBoundingClientRect().height), texts: [...holder.querySelectorAll("span[lang]")].map(span => span.textContent), measured: harness.analyze(content, { wrap: false }).measurements[0].bounds.height };
    } finally { holder.remove(); }
  });
  expect(result.heights).toEqual([90, 30]);
  expect(result.heights.reduce((sum, height) => sum + height, 0)).toBeCloseTo(result.measured, 1);
  expect(result.texts).toEqual(["M\n", "\n", ""]);
});

test("text paint alpha composites in native SVG pixels and mixed-run HTML without changing layout", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const samples = [];
    for (const [colorOpacity, opacity] of [[1, 1], [.5, 1], [0, 1], [.5, .5]]) {
      const p = createPresentation({ size: { width: 200, height: 80 } });
      p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 200, height: 80 }, opacity, frame: { margin: 0 },
        content: [{ runs: [{ text: "MMMM", style: { color: "000000", colorOpacity, fontSize: 40 } }] }] });
      const rendered = await renderPresentationToSvg(p);
      const svg = rendered.slides[0].svg, image = new Image();
      const url = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml" }));
      try {
        image.src = url;
        await image.decode();
        const canvas = document.createElement("canvas");
        canvas.width = 200; canvas.height = 80;
        const context = canvas.getContext("2d");
        context.fillStyle = "#FFFFFF"; context.fillRect(0, 0, 200, 80);
        context.drawImage(image, 0, 0, 200, 80);
        const pixels = context.getImageData(0, 0, 200, 80).data;
        let minimum = 255;
        for (let i = 0; i < pixels.length; i += 4) minimum = Math.min(minimum, pixels[i]);
        samples.push(minimum);
      } finally { URL.revokeObjectURL(url); }
    }
    const p = createPresentation();
    p.addSlide().addElement({ type: "text", box: { x: 0, y: 0, width: 300, height: 80 }, opacity: .5,
      content: [{ style: { bullet: { type: "bullet" } }, runs: [0, .25, 1].map(colorOpacity => ({ text: "M", style: { colorOpacity, fontSize: 40 } })) }] });
    const svg = (await renderPresentationToSvg(p)).slides[0].svg;
    const holder = document.createElement("div");
    holder.innerHTML = svg; document.body.append(holder);
    try {
      const spans = [...holder.querySelectorAll("foreignObject span[lang]")];
      return { samples, cssAlpha: spans.map(span => getComputedStyle(span).opacity), bulletAlpha: getComputedStyle(holder.querySelector('span[aria-hidden="true"]')).opacity, widths: spans.map(span => span.getBoundingClientRect().width) };
    } finally { holder.remove(); }
  });
  for (const [index, expected] of [0, 127, 255, 191].entries()) expect(Math.abs(result.samples[index] - expected)).toBeLessThanOrEqual(2);
  expect(result.cssAlpha).toEqual(["0", "0.25", "1"]);
  expect(result.bulletAlpha).toEqual("0");
  expect(result.widths[0]).toBeGreaterThan(0);
  expect(result.widths[1]).toEqual(result.widths[0]);
  expect(result.widths[2]).toEqual(result.widths[0]);
});

test("wrap, explicit repeated/trailing breaks, no-wrap and empty paragraphs have distinct measured geometry", async ({ page }) => {
  const results = await page.evaluate(() => {
    const measure = (text, options) => harness.analyze(text, options).measurements[0].bounds;
    return { wrapped: measure("MMMM", { width: 25 }), unwrapped: measure("MMMM", { width: 25, wrap: false }), breaks: measure("M\n\nM\n"), empty: measure(""), zero: measure("M", { width: 0 }) };
  });
  expect(results.wrapped.height).toBeCloseTo(60, 1);
  expect(results.unwrapped.height).toBeCloseTo(30, 1);
  expect(results.unwrapped.width).toBeGreaterThan(25);
  expect(results.breaks.height).toBeCloseTo(120, 1);
  expect(results.empty.height).toBeCloseTo(30, 1);
  expect(results.empty.width).toBe(0);
  expect(results.zero.width).toBeGreaterThan(0);
});

test("mixed sizes and CJK use pinned fonts and wrap between CJK characters", async ({ page }) => {
  const result = await page.evaluate(() => {
    const content = [{ style: { lineSpacing: 1.5 }, runs: [
      { text: "MM", style: { fontSize: 10 } },
      { text: "中文", style: { fontFamily: "Pinned CJK", fontSize: 20, language: "zh-CN" } },
    ] }];
    return { mixed: harness.analyze(content), expected: harness.width("MM", 10) + harness.width("中文", 20, "Pinned CJK"), wrapped: harness.analyze([{ style: { lineSpacing: 1.5 }, runs: [{ text: "中文测量", style: { fontSize: 20, fontFamily: "Pinned CJK" } }] }], { width: 41 }) };
  });
  expect(result.mixed.diagnostics).toEqual([]);
  expect(result.mixed.measurements[0].bounds.width).toBeCloseTo(result.expected, 1);
  expect(result.wrapped.diagnostics).toEqual([]);
  expect(result.wrapped.measurements[0].bounds.height).toBeCloseTo(60, 1);
});

test("margins, shrink, vertical alignment and paragraph spacing reach the browser provider", async ({ page }) => {
  const result = await page.evaluate(() => {
    const content = [{ style: { spaceBefore: 4, spaceAfter: 6, lineSpacing: 1.5 }, runs: [{ text: "MMMM", style: { fontSize: 20 } }] }];
    return { plain: harness.analyze(content), adjusted: harness.analyze(content, { width: 100, height: 80, margin: 10, fontScale: 0.5, verticalAlign: "bottom" }) };
  });
  expect(result.plain.measurements[0].bounds.height).toBeCloseTo(40, 1);
  const bounds = result.adjusted.measurements[0].bounds;
  expect(bounds.height).toBeCloseTo(25, 1);
  expect(bounds.y + bounds.height).toBeCloseTo(70, 1);
  expect(bounds.width).toBeCloseTo(result.plain.measurements[0].bounds.width / 2, 1);
});

test("numbered markers, hanging indent, alignment and overflowing text are measured", async ({ page }) => {
  const result = await page.evaluate(() => {
    const paragraph = (align, bullet = { type: "none" }) => [{ style: { align, bullet, indent: 10, hanging: 20, lineSpacing: 1.5 }, runs: [{ text: "M", style: { fontSize: 20 } }] }];
    return {
      hanging: harness.analyze(paragraph("left")),
      numbered: harness.analyze(paragraph("left", { type: "number", style: "alphaUpperPeriod", startAt: 27 })),
      beyond: harness.analyze(paragraph("left", { type: "number", style: "alphaUpperPeriod", startAt: 53 })),
      right: harness.analyze(paragraph("right")),
      expectedNumber: harness.width("AA. M"),
      expectedBeyond: harness.width("BA. M"),
      overflow: harness.analyze("MMMM", { width: 25, height: 30 }),
    };
  });
  expect(result.hanging.measurements[0].bounds.x).toBeCloseTo(-10, 1);
  expect(result.numbered.measurements[0].bounds.width).toBeCloseTo(result.expectedNumber, 1);
  expect(result.beyond.measurements[0].bounds.width).toBeCloseTo(result.expectedBeyond, 1);
  expect(result.right.measurements[0].bounds.width).toBeCloseTo(200, 1);
  expect(result.overflow.diagnostics.map(({ code }) => code)).toContain("text-overflow");
});

test("unknown and unloaded fonts never silently use system fallback; declared fallback emits diagnostics", async ({ page }) => {
  const result = await page.evaluate(() => {
    const content = [{ runs: [{ text: "Missing", style: { fontFamily: "Unregistered", fontSize: 20 } }] }];
    const missing = harness.analyze(content);
    const fallback = harness.createBrowserTextMeasurement({ document, fonts: harness.fonts, fallbackFamilies: ["Pinned Mono"] });
    const recovered = harness.analyze(content, { provider: fallback });
    fallback.dispose();
    document.fonts.delete(harness.latinFace);
    const removed = harness.analyze("Gone");
    return { missing, recovered, removed };
  });
  expect(result.missing.diagnostics[0].code).toBe("font-unavailable");
  expect(result.missing.measurements).toEqual([]);
  expect(result.recovered.diagnostics[0].code).toBe("font-fallback");
  expect(result.recovered.measurements).toHaveLength(1);
  expect(result.removed.diagnostics[0].code).toBe("font-unavailable");
});

test("unsupported glyphs, marker coverage, font styles and invalid providers fail explicitly", async ({ page }) => {
  const result = await page.evaluate(() => {
    return {
      glyph: harness.analyze("中文"),
      style: harness.analyze([{ runs: [{ text: "Bold", style: { bold: true } }] }]),
      marker: harness.analyze([{ style: { bullet: { type: "bullet", character: "•" } }, runs: [{ text: "List" }] }]),
      invalid: (() => { try { harness.createBrowserTextMeasurement({ document, fonts: harness.fonts, fallbackFamilies: ["Unknown"] }); return false; } catch { return true; } })(),
    };
  });
  expect(result.glyph.diagnostics[0].code).toBe("font-unavailable");
  expect(result.style.diagnostics[0].code).toBe("font-unavailable");
  expect(result.marker.diagnostics[0].code).toBe("text-measurement-failed");
  expect(result.invalid).toBe(true);
});

test("shadow isolation resists host styles, content is literal, and disposal removes all resources", async ({ page }) => {
  const result = await page.evaluate(() => {
    const text = "<img src=x onerror=alert(1)>";
    const before = harness.analyze(text);
    const style = document.createElement("style");
    style.textContent = "div,span{font-size:99px!important;padding:100px!important;transform:scale(2)!important}*{line-height:9!important}";
    document.head.appendChild(style);
    const after = harness.analyze(text);
    const count = document.querySelectorAll("[data-pptkit-measurement]").length;
    harness.adapter.dispose(); harness.adapter.dispose();
    const disposed = harness.analyze("After disposal");
    return { before, after, count, remaining: document.querySelectorAll("[data-pptkit-measurement]").length, images: document.images.length, disposed };
  });
  expect(result.after).toEqual(result.before);
  expect(result.count).toBe(1);
  expect(result.remaining).toBe(0);
  expect(result.images).toBe(0);
  expect(result.disposed.diagnostics[0].code).toBe("font-resolution-failed");
});

test("detached/hidden and transformed documents do not produce false fitting measurements", async ({ page }) => {
  const result = await page.evaluate(() => {
    document.documentElement.style.transform = "scale(2)";
    const transformed = harness.analyze("M");
    document.documentElement.style.transform = "";
    document.documentElement.style.display = "none";
    const hidden = harness.analyze("M");
    document.documentElement.style.display = "";
    const restored = harness.analyze("M");
    return { transformed, hidden, restored };
  });
  expect(result.transformed.measurements).toEqual([]);
  expect(result.transformed.diagnostics[0].code).toBe("font-resolution-failed");
  expect(result.hidden.measurements).toEqual([]);
  expect(result.hidden.diagnostics[0].code).toBe("text-measurement-failed");
  expect(result.restored.diagnostics).toEqual([]);
});


test("loaded CJK bullet markers and wrapping diagnostics are repeatable without extra hosts", async ({ page }) => {
  const result = await page.evaluate(() => {
    const content = [{ style: { bullet: { type: "bullet", character: "•" }, lineSpacing: 1.5 }, runs: [{ text: "中文测量", style: { fontFamily: "Pinned CJK", fontSize: 20 } }] }];
    const first = harness.analyze(content, { width: 50, height: 30 });
    const second = harness.analyze(content, { width: 50, height: 30 });
    return { first, second, hosts: document.querySelectorAll("[data-pptkit-measurement]").length };
  });
  expect(result.first).toEqual(result.second);
  expect(result.hosts).toBe(1);
  expect(result.first.measurements).toHaveLength(1);
  expect(result.first.diagnostics.map(({ code }) => code)).toEqual(["text-overflow"]);
});

test("font renaming and async coverage cannot bypass availability checks", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const faulty = harness.createBrowserTextMeasurement({ document, fonts: [{ ...harness.fonts[0], covers: async () => { throw new Error("late coverage failure"); } }] });
    const invalid = harness.analyze("M", { provider: faulty });
    faulty.dispose();
    harness.latinFace.family = "Renamed";
    const renamed = harness.analyze("M");
    await new Promise((resolve) => setTimeout(resolve, 0));
    return { invalid, renamed };
  });
  expect(result.invalid.diagnostics[0].code).toBe("font-resolution-failed");
  expect(result.invalid.diagnostics[0].message).toContain("synchronous boolean");
  expect(result.renamed.diagnostics[0].code).toBe("font-unavailable");
});

test("unloaded faces are not started by measurement", async ({ page }) => {
  const result = await page.evaluate(() => {
    const face = new FontFace("Not Loaded", "url(https://unavailable.invalid/font.woff2)");
    document.fonts.add(face);
    const adapter = harness.createBrowserTextMeasurement({ document, fonts: [{ family: "Not Loaded", faces: [face], covers: () => true }] });
    const analysis = harness.analyze([{ runs: [{ text: "M", style: { fontFamily: "Not Loaded" } }] }], { provider: adapter });
    adapter.dispose();
    return { analysis, status: face.status };
  });
  expect(result.analysis.diagnostics[0].code).toBe("font-unavailable");
  expect(result.status).toBe("unloaded");
});


for (const clipped of [false, true]) test(`measured table geometry matches browser cell bounds and exported rows (clip=${clipped})`, async ({ page }) => {
  const result = await page.evaluate(async (clipped) => {
    const { createPresentation, normalizePresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { renderLayoutToSvg } = await import("@pptkit/svg-renderer");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    const cell = (text, rowSpan = 1) => ({ content: [{ style: { lineSpacing: 1.5 }, runs: [{ text, style: { fontSize: 20 } }] }], rowSpan, style: { margin: 2 } });
    deck.addSlide({ id: "s" }).addElement({ type: "table", id: "t", box: { x: 20, y: 30, width: 80, height: 20 }, columns: [40, 40], rows: [
      { height: 10, cells: [cell("MMMMMMMM", 2), cell("M")] },
      { height: 10, cells: [cell("MM")] },
    ] });
    const before = JSON.stringify(normalizePresentation(deck));
    const measured = [];
    const provider = { ...harness.adapter, measure(request) { const value = harness.adapter.measure(request); measured.push({ available: request.availableWidth, ...value }); return value; } };
    const layout = { textMeasurement: provider, tableSizing: [{ elementId: "t", autoRows: [{ row: 0, maxHeight: clipped ? 20 : 50 }, { row: 1, ...(clipped ? { maxHeight: 20 } : {}) }], overflow: clipped ? "clip" : "preserve" }] };
    const resolved = resolveLayout(deck, layout);
    const table = resolved.slides[0].elements[0];
    if (!table.cellLayout) throw new Error(JSON.stringify({ diagnostics: resolved.analysis.diagnostics, measured }));
    const preview = await renderLayoutToSvg(resolved);
    const frame = document.createElement("div");
    frame.innerHTML = preview.slides[0].svg;
    document.body.appendChild(frame);
    const origin = frame.querySelector("foreignObject").getBoundingClientRect();
    const cells = [...frame.querySelectorAll("[data-pptkit-cell]")].map((node) => {
      const rect = node.getBoundingClientRect();
      const textRects = [...node.querySelectorAll("span")].flatMap((span) => {
        const range = document.createRange(); range.selectNodeContents(span); return [...range.getClientRects()];
      });
      const textBottom = Math.max(0, ...textRects.map((textRect) => textRect.bottom - rect.y));
      return { x: rect.x - origin.x, y: rect.y - origin.y, width: rect.width, height: rect.height, textBottom, overflow: getComputedStyle(node).overflow };
    });
    const exported = await generatePptx(deck, { layout });
    return { cells, expected: table.cellLayout.map(({ box }) => box), heights: table.rows.map(({ height }) => height), tableHeight: table.box.height, warnings: exported.warnings.map(({ code }) => code), bytes: [...exported.bytes], unchanged: before === JSON.stringify(normalizePresentation(deck)) };
  }, clipped);
  expect(result.unchanged).toBe(true);
  expect(result.heights).toEqual(clipped ? [20, 20] : [47, 47]);
  expect(result.tableHeight).toBe(clipped ? 40 : 94);
  result.cells.forEach((cell, index) => {
    for (const key of ["x", "y", "width", "height"]) expect(cell[key]).toBeCloseTo(result.expected[index][key], 1);
    expect(cell.overflow).toBe("hidden");
  });
  expect(result.cells.some((cell) => cell.textBottom > cell.height + 0.02)).toBe(clipped);
  expect(result.warnings.some((code) => code === "table-row-overflow")).toBe(clipped);
  const slide = readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
  expect([...slide.matchAll(/<a:tr h="(\d+)">/g)].map((match) => Number(match[1]) / 12700)).toEqual(result.heights);
  expect(slide).toContain(`<a:ext cx="1016000" cy="${result.tableHeight * 12700}"/>`);
});


test("intrinsic text uses loaded font advances and bounded wrapping in exported geometry", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    deck.addSlide({ id: "s" }).addElement({ type: "text", id: "t", box: { x: 10, y: 20, width: 1, height: 1 }, frame: { margin: 5 }, content: [{ style: { lineSpacing: 1.5 }, runs: [{ text: "MMMM", style: { fontSize: 20 } }] }] });
    const options = { textMeasurement: harness.adapter, intrinsicSizing: [{ elementId: "t", width: { type: "content" }, height: { type: "content" } }] };
    const natural = resolveLayout(deck, options);
    const bounded = { ...options, intrinsicSizing: [{ elementId: "t", width: { type: "content", max: 35 }, height: { type: "content" } }] };
    const wrapped = resolveLayout(deck, bounded);
    const exported = await generatePptx(deck, { layout: bounded });
    return { natural: natural.slides[0].elements[0].box, wrapped: wrapped.slides[0].elements[0].box, diagnostics: [...natural.analysis.diagnostics, ...wrapped.analysis.diagnostics], expectedWidth: harness.width("MMMM") + 10, bytes: [...exported.bytes] };
  });
  expect(result.diagnostics).toEqual([]);
  expect(result.natural.width).toBeCloseTo(result.expectedWidth, 1);
  expect(result.natural.height).toBeCloseTo(40, 1);
  expect(result.wrapped.width).toBe(35);
  expect(result.wrapped.height).toBeCloseTo(70, 1);
  const xml = readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
  expect(xml).toContain(`<a:ext cx="${35 * 12700}" cy="${Math.round(result.wrapped.height * 12700)}"/>`);
});

test("wrapped parallelogram intrinsic sizing uses real font metrics across preview and export", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck=createPresentation({size:{width:400,height:300},theme:{fonts:{body:"Pinned Mono",heading:"Pinned Mono"}}});
    deck.addSlide({id:"s"}).addElement({id:"p",type:"shape",shape:"parallelogram",slant:.4,box:{x:10,y:20,width:20,height:20},text:{frame:{margin:5,wrap:true},content:[{style:{lineSpacing:1.5},runs:[{text:"MM MM MM MM",style:{fontSize:20}}]}]}});
    const calls=[], provider={...harness.adapter,measure:r=>{const result=harness.adapter.measure(r);calls.push({width:r.availableWidth,wrap:r.wrap,bounds:result.bounds});return result;}};
    const layout={textMeasurement:provider,intrinsicSizing:[{elementId:"p",width:{type:"content",max:120},height:{type:"content"}}],diagnostics:{textOverflow:true}};
    const resolved=resolveLayout(deck,layout),preview=await renderPresentationToSvg(deck,{layout}),exported=await generatePptx(deck,{layout});
    const frame=document.createElement("div");frame.innerHTML=preview.slides[0].svg;document.body.appendChild(frame);
    const foreign=frame.querySelector("foreignObject");
    return {calls,box:resolved.slides[0].elements[0].box,analysis:resolved.analysis,foreign:{x:Number(foreign.getAttribute("x")),y:Number(foreign.getAttribute("y")),width:Number(foreign.getAttribute("width")),height:Number(foreign.getAttribute("height"))},bytes:[...exported.bytes]};
  });
  expect(result.analysis.diagnostics,JSON.stringify(result.calls)).toEqual([]);
  expect(result.box.width).toBe(120);
  expect(result.analysis.measurements[0].bounds.height).toBeGreaterThan(30);
  const content=result.analysis.measurements[0].contentBox;
  expect(result.foreign.x).toBeCloseTo(10+content.x-5,5);
  expect(result.foreign.y).toBeCloseTo(20+content.y-5,5);
  expect(result.foreign.width).toBeCloseTo(content.width+10,5);
  expect(result.foreign.height).toBeCloseTo(content.height+10,5);
  const xml=readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
  expect(xml).toContain(`<a:ext cx="${120*12700}" cy="${Math.round(result.box.height*12700)}"/>`);
});

test("real-font wrapped parallelogram rejects measured horizontal overflow atomically", async ({ page }) => {
  const result=await page.evaluate(async()=>{
    const {createPresentation}=await import("@pptkit/core"),{resolveLayout}=await import("@pptkit/layout");
    const deck=createPresentation({size:{width:400,height:300},theme:{fonts:{body:"Pinned Mono",heading:"Pinned Mono"}}});
    deck.addSlide({id:"s"}).addElement({id:"p",type:"shape",shape:"parallelogram",slant:.4,box:{x:10,y:20,width:20,height:20},text:{frame:{margin:5,wrap:true},content:[{style:{lineSpacing:1.5},runs:[{text:"MMMM MMMM MMMM",style:{fontSize:20}}]}]}});
    return resolveLayout(deck,{textMeasurement:harness.adapter,intrinsicSizing:[{elementId:"p",width:{type:"content",max:120},height:{type:"content"}}],diagnostics:{textOverflow:true}});
  });
  expect(result.analysis.diagnostics.some(d=>d.code==="intrinsic-shape-height-unavailable")).toBe(true);
  expect(result.slides[0].elements[0].box).toEqual({x:10,y:20,width:20,height:20});
});

test("real-font wrapped parallelogram widens when its first candidate exceeds height limit", async ({ page }) => {
  const result=await page.evaluate(async()=>{
    const {createPresentation}=await import("@pptkit/core"),{resolveLayout}=await import("@pptkit/layout");
    const deck=createPresentation({size:{width:800,height:300},theme:{fonts:{body:"Pinned Mono",heading:"Pinned Mono"}}});
    deck.addSlide({id:"s"}).addElement({id:"p",type:"shape",shape:"parallelogram",slant:.4,box:{x:10,y:20,width:20,height:20},text:{frame:{margin:5,wrap:true},content:[{style:{lineSpacing:1.5},runs:[{text:"MMMM MMMM MMMM",style:{fontSize:20}}]}]}});
    const seen=[],provider={...harness.adapter,measure:r=>{const measured=harness.adapter.measure(r);seen.push({wrap:r.wrap,width:r.availableWidth,height:measured.bounds.height});return measured;}};
    const resolved=resolveLayout(deck,{textMeasurement:provider,intrinsicSizing:[{elementId:"p",width:{type:"content",max:600},height:{type:"content",max:80}}],diagnostics:{textOverflow:true}});
    return {box:resolved.slides[0].elements[0].box,analysis:resolved.analysis,seen};
  });
  expect(result.analysis.diagnostics,JSON.stringify(result.seen)).toEqual([]);
  expect(result.box.width).toBeGreaterThan(300);
  expect(result.box.height).toBeLessThanOrEqual(80);
  expect(result.seen.some(item=>item.wrap&&item.height>30)).toBe(true);
});

test("paginated tables render repeated headers and exact cell frames on every exported page", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { renderLayoutToSvg } = await import("@pptkit/svg-renderer");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck = createPresentation({ size: { width: 200, height: 120 }, theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    const cell = (text) => ({ content: [{ style: { lineSpacing: 1 }, runs: [{ text, style: { fontSize: 10 } }] }], style: { margin: 0 } });
    deck.addSlide({ id: "s" }).addElement({ type: "table", id: "t", box: { x: 0, y: 0, width: 100, height: 130 }, columns: [100], rows: Array.from({ length: 7 }, (_, i) => ({ height: i === 0 ? 10 : 20, cells: [cell(i === 0 ? "Header" : `Body ${i}`)] })) });
    const layout = { tablePagination: [{ elementId: "t", area: { x: 20, y: 30, width: 100, height: 55 }, headerRows: 1 }] };
    const resolved = resolveLayout(deck, layout);
    const preview = await renderLayoutToSvg(resolved);
    const pages = preview.slides.map((slide) => {
      const frame = document.createElement("div"); frame.innerHTML = slide.svg; document.body.appendChild(frame);
      const origin = frame.querySelector("svg").getBoundingClientRect();
      const cells = [...frame.querySelectorAll("[data-pptkit-cell]")].map((node) => {
        const rect = node.getBoundingClientRect();
        return { x: rect.x - origin.x, y: rect.y - origin.y, width: rect.width, height: rect.height, text: node.textContent };
      });
      return cells;
    });
    const exported = await generatePptx(deck, { layout });
    return { pages, fragments: resolved.pagination.fragments, warnings: exported.warnings, count: exported.slideCount, bytes: [...exported.bytes] };
  });
  expect(result.warnings).toEqual([]);
  expect(result.count).toBe(3);
  expect(result.pages.map((cells) => cells.map((cell) => cell.text))).toEqual([["Header", "Body 1", "Body 2"], ["Header", "Body 3", "Body 4"], ["Header", "Body 5", "Body 6"]]);
  const entries = readGeneratedZip(Uint8Array.from(result.bytes));
  result.pages.forEach((cells, index) => {
    cells.forEach((cell, row) => {
      expect(cell.x).toBeCloseTo(20, 1); expect(cell.y).toBeCloseTo([30, 40, 60][row], 1);
      expect(cell.width).toBeCloseTo(100, 1); expect(cell.height).toBeCloseTo(row === 0 ? 10 : 20, 1);
    });
    const xml = entries.get(`ppt/slides/slide${index + 1}.xml`).toString();
    expect([...xml.matchAll(/<a:tr h="(\d+)">/g)].map((match) => Number(match[1]) / 12700)).toEqual([10, 20, 20]);
  });
});

test("native SVG tab anchors preserve mixed font sizes for text and shapes", async ({ page }) => {
  const results = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const deck = createPresentation(), slide = deck.addSlide();
    for (const [index, align] of ["left", "center", "right", "decimal"].entries()) {
      const suffix = align === "decimal" ? ["12", ".34"] : ["WI", "DE"];
      const content = [{ style: { lineSpacing: 1.5, tabStops: [{ position: 140, align }] }, runs: [
        { text: "P\t", style: { fontSize: 16 } }, { text: suffix[0], style: { fontSize: 24, bold: true } }, { text: suffix[1], style: { fontSize: 32 } },
      ] }];
      const frame = { margin: 0, wrap: false }, box = { x: 20, y: index * 100, width: 300, height: 80 };
      slide.addElement({ id: `text-${align}`, type: "text", box, frame, content });
      slide.addElement({ id: `shape-${align}`, type: "shape", shape: "rect", box: { ...box, x: 350 }, text: { frame, content } });
    }
    const before = JSON.stringify(slide.elements), result = await renderPresentationToSvg(deck), holder = document.createElement("div");
    holder.innerHTML = result.slides[0].svg; document.body.append(holder);
    const items = ["left", "center", "right", "decimal"].flatMap(align => ["text", "shape"].map(kind => {
      const element = holder.querySelector(`[data-pptkit-element-id="${kind}-${align}"]`), chunks = [...element.querySelectorAll("text[data-pptkit-tab-part]")];
      return { align, kind, fallback: !!element.querySelector("foreignObject"), parts: chunks.map(node => ({ part: node.dataset.pptkitTabPart, anchor: node.getAttribute("text-anchor"), x: Number(node.getAttribute("x")), y: Number(node.getAttribute("y")), sizes: [...node.querySelectorAll("tspan")].map(span => span.getAttribute("font-size")), start: node.getStartPositionOfChar(0).x, end: node.getEndPositionOfChar(node.getNumberOfChars() - 1).x })) };
    }));
    holder.remove();
    return { items, unchanged: before === JSON.stringify(slide.elements), warnings: result.warnings.filter(w => w.code === "text-tab-layout-unsupported").length };
  });
  expect(results.unchanged).toBe(true); expect(results.warnings).toBe(8);
  for (const item of results.items) {
    expect(item.fallback, `${item.kind}-${item.align}`).toBe(false);
    const prefix = item.parts.find(part => part.part === "prefix");
    expect(prefix.sizes).toEqual(["16", "0"]);
    const target = item.kind === "text" ? 160 : 490;
    expect(item.parts.every(part => part.y === prefix.y)).toBe(true);
    expect(prefix.y).toBe(["left", "center", "right", "decimal"].indexOf(item.align) * 100 + 32);
    if (item.align === "decimal") {
      const integer = item.parts.find(part => part.part === "integer"), decimal = item.parts.find(part => part.part === "decimal");
      expect(integer.sizes).toEqual(["24"]); expect(decimal.sizes).toEqual(["32"]);
      expect(integer.end).toBeCloseTo(target, 1); expect(decimal.start).toBeCloseTo(target, 1);
    } else {
      const suffix = item.parts.find(part => part.part === "suffix");
      expect(suffix.sizes).toEqual(["24", "32"]);
      expect(item.align === "left" ? suffix.start : item.align === "center" ? (suffix.start + suffix.end) / 2 : suffix.end).toBeCloseTo(target, 1);
    }
  }
});

test("native tab anchors preserve left-to-right CJK script font segments", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core"), { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const deck = createPresentation(), slide = deck.addSlide();
    const samples = [["left", "合計\t売上"], ["center", "売上\tカタカナ"], ["right", "매출\t한국어"], ["decimal", "合計\t12.34"]];
    for (const [i, [align, text]] of samples.entries()) {
      const content = [{ style: { tabStops: [{ position: 160, align }] }, runs: [{ text, style: { fontFamily: "Latin Test Face", eastAsianFontFamily: "CJK Test Face", fontSize: 24 } }] }];
      const frame = { margin: 0, wrap: false }, box = { x: 20, y: i * 100, width: 300, height: 70 };
      slide.addElement({ id: `text-${align}`, type: "text", box, frame, content });
      slide.addElement({ id: `shape-${align}`, type: "shape", shape: "rect", box: { ...box, x: 350 }, text: { frame, content } });
    }
    const before = JSON.stringify(slide.elements), rendered = await renderPresentationToSvg(deck), holder = document.createElement("div"); holder.innerHTML = rendered.slides[0].svg; document.body.append(holder);
    const items = samples.flatMap(([align, text]) => ["text", "shape"].map(kind => {
      const element = holder.querySelector(`[data-pptkit-element-id="${kind}-${align}"]`), parts = [...element.querySelectorAll("text[data-pptkit-tab-part]")];
      return { align, text, kind, fallback: !!element.querySelector("foreignObject"), parts: parts.map(node => ({ part: node.dataset.pptkitTabPart, text: node.textContent, y: Number(node.getAttribute("y")), families: [...node.querySelectorAll('tspan[font-family]')].filter(span => span.textContent).map(span => span.getAttribute("font-family")), start: node.getStartPositionOfChar(0).x, end: node.getEndPositionOfChar(node.getNumberOfChars() - 1).x })) };
    }));
    holder.remove(); return { items, unchanged: before === JSON.stringify(slide.elements), warnings: rendered.warnings.filter(w => w.code === "text-tab-layout-unsupported").length };
  });
  expect(result.unchanged).toBe(true); expect(result.warnings).toBe(8);
  for (const item of result.items) {
    expect(item.fallback, `${item.kind}-${item.align}`).toBe(false); expect(item.parts.map(part => part.text).join("")).toBe(item.text);
    expect(item.parts[0].families.join(" ")).toContain("CJK Test Face");
    expect(item.parts.every(part => part.y === ["left", "center", "right", "decimal"].indexOf(item.align) * 100 + 24)).toBe(true);
    const target = item.kind === "text" ? 180 : 510;
    if (item.align === "decimal") {
      expect(item.parts.find(part => part.part === "integer").end).toBeCloseTo(target, 1);
      expect(item.parts.find(part => part.part === "decimal").start).toBeCloseTo(target, 1);
    } else {
      const suffix = item.parts.find(part => part.part === "suffix"); expect(suffix.families.join(" ")).toContain("CJK Test Face");
      expect(item.align === "left" ? suffix.start : item.align === "center" ? (suffix.start + suffix.end) / 2 : suffix.end).toBeCloseTo(target, 1);
    }
  }
});

test("quarter-turn intrinsic sizing uses real wrapped font bounds and exports solved axes", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    const slide = deck.addSlide({ id: "s" });
    for (const [id, textDirection, x] of [["cw", "rotate90", 10], ["ccw", "rotate270", 100]]) {
      slide.addElement({ id, type: "text", box: { x, y: 20, width: 1, height: 1 },
        frame: { textDirection, margin: 5 }, content: [{ style: { lineSpacing: 1.5 }, runs: [{ text: "MMMM", style: { fontSize: 20 } }] }] });
    }
    const layout = { textMeasurement: harness.adapter, intrinsicSizing: ["cw", "ccw"].map(elementId => ({
      elementId, height: { type: "content", max: 35 }, width: { type: "content" },
    })), diagnostics: { textOverflow: true } };
    const resolved = resolveLayout(deck, layout);
    const exported = await generatePptx(deck, { layout });
    const host = document.createElement("div");
    host.innerHTML = (await renderPresentationToSvg(deck, { layout })).slides[0].svg;
    document.body.appendChild(host);
    const preview = Array.from(host.querySelectorAll("foreignObject"), foreign => {
      const range = document.createRange();
      range.selectNodeContents(foreign.firstElementChild);
      const text = range.getBoundingClientRect(), frame = foreign.getBoundingClientRect();
      return { textWidth: text.width, textHeight: text.height, width: frame.width, height: frame.height,
        inside: text.left >= frame.left - 1 && text.right <= frame.right + 1 &&
          text.top >= frame.top - 1 && text.bottom <= frame.bottom + 1 };
    });
    return { preview, boxes: resolved.slides[0].elements.map(e => e.box), diagnostics: resolved.analysis.diagnostics.map(d => d.code), bytes: [...exported.bytes] };
  });
  expect(result.diagnostics).toEqual(["text-frame-direction-measurement-approximate", "text-frame-direction-measurement-approximate"]);
  for (const box of result.boxes) {
    expect(box.width).toBeCloseTo(70, 1);
    expect(box.height).toBe(35);
  }
  expect(result.preview).toHaveLength(2);
  for (const frame of result.preview) {
    expect(frame.inside).toBe(true);
    expect(frame.width / frame.height).toBeCloseTo(2, 2);
    expect(frame.textWidth).toBeGreaterThan(0);
    expect(frame.textHeight).toBeGreaterThan(0);
  }
  const xml = readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
  expect(xml.match(/<a:ext cx="889000" cy="444500"\/>/g)).toHaveLength(2);
  expect(xml).toMatch(/<a:bodyPr[^>]* vert="vert"/);
  expect(xml).toMatch(/<a:bodyPr[^>]* vert="vert270"/);
});

test("East Asian vertical previews lay out mixed runs along the physical height", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    const slide = deck.addSlide();
    const textDirection = "eastAsianVertical", content = [{ runs: [{ text: "日本語" }, { text: "ABC" }] }];
    slide.addElement({ type: "text", id: "v", box: { x: 20, y: 20, width: 100, height: 240 }, frame: { textDirection }, content });
    slide.addElement({ type: "table", id: "t", box: { x: 150, y: 20, width: 100, height: 240 }, columns: [100], rows: [
      { height: 240, cells: [{ content, style: { textDirection } }] },
    ] });
    const preview = await renderPresentationToSvg(deck);
    const host = document.createElement("div");
    host.innerHTML = preview.slides[0].svg;
    document.body.appendChild(host);
    return Array.from(host.querySelectorAll('[style*="writing-mode:vertical-rl"]'), frame => {
      const range = document.createRange(), walker = document.createTreeWalker(frame, NodeFilter.SHOW_TEXT);
      let run;
      while ((run = walker.nextNode()) && run.textContent !== "ABC") {}
      if (!run) throw Error("Latin run is absent");
      const positions = Array.from("ABC", (_, i) => {
        range.setStart(run, i); range.setEnd(run, i + 1);
        const box = range.getBoundingClientRect();
        return { x: box.x, y: box.y, w: box.width, h: box.height };
      });
      return { positions, mode: getComputedStyle(frame).writingMode, orientation: getComputedStyle(frame).textOrientation };
    });
  });
  expect(result).toHaveLength(2);
  for (const { positions, mode, orientation } of result) {
    expect(mode).toBe("vertical-rl");
    expect(orientation).toBe("mixed");
    for (const [i, box] of positions.entries()) {
      expect(box.w).toBeGreaterThan(0); expect(box.h).toBeGreaterThan(0);
      if (i) {
        expect(box.x).toBeCloseTo(positions[0].x, 1);
        expect(box.y).toBeGreaterThan(positions[i - 1].y);
      }
    }
  }
});

test("built-in East Asian measurement tracks vertical advances, wrapping and unsupported profiles", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    function run(height, wrap, kind = "plain") {
      const deck = createPresentation({ theme: { fonts: { body: "Pinned International", heading: "Pinned International" } } });
      const runs = [{ text: kind === "tab" ? "中文\tABCD" : kind === "missing" ? "中文😀" : "中文ABCD", style: { fontSize: 20 } }];
      deck.addSlide().addElement({ type: "text", id: "v", box: { x: 0, y: 0, width: 130, height },
        frame: { textDirection: "eastAsianVertical", margin: 5, wrap },
        content: [{ runs, style: { lineSpacing: 1.5, ...(kind === "list" ? { bullet: { type: "bullet" } } : {}) } }] });
      return resolveLayout(deck, { textMeasurement: harness.adapter, diagnostics: { textOverflow: true } }).analysis;
    }
    return { tall: run(180, false), wrapped: run(55, true), overflow: run(55, false),
      tab: run(180, true, "tab"), list: run(180, true, "list"), missing: run(180, true, "missing"),
      latin: harness.width("ABCD", 20, "Pinned International") };
  });
  const tall = result.tall.measurements[0].bounds, wrapped = result.wrapped.measurements[0].bounds;
  expect(result.tall.diagnostics).toEqual([]);
  expect(result.wrapped.diagnostics).toEqual([]);
  expect(tall.height).toBeCloseTo(40 + result.latin, 0);
  expect(wrapped.width).toBeGreaterThan(tall.width);
  expect(wrapped.height).toBeLessThan(tall.height);
  expect(result.overflow.diagnostics.some(d => d.code === "text-overflow")).toBe(true);
  for (const key of ["tab", "list"]) {
    expect(result[key].measurements).toHaveLength(0);
    expect(result[key].diagnostics.some(d => d.code === "text-measurement-failed")).toBe(true);
  }
  expect(result.missing.measurements).toHaveLength(0);
  expect(result.missing.diagnostics.some(d => d.code === "font-unavailable")).toBe(true);
});

test("East Asian intrinsic sizing uses wrapped browser columns and exports physical extents", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned International", heading: "Pinned International" } } });
    deck.addSlide().addElement({ type: "text", id: "v", box: { x: 20, y: 20, width: 150, height: 180 },
      frame: { textDirection: "eastAsianVertical", margin: 5 },
      content: [{ style: { lineSpacing: 1.5 }, runs: [{ text: "中文ABCD", style: { fontSize: 20 } }] }] });
    const layout = { textMeasurement: harness.adapter, intrinsicSizing: [
      { elementId: "v", height: { type: "content", max: 55 }, width: { type: "content" } },
    ], diagnostics: { textOverflow: true } };
    const resolved = resolveLayout(deck, layout);
    const exported = await generatePptx(deck, { layout });
    return { box: resolved.slides[0].elements[0].box, diagnostics: resolved.analysis.diagnostics, bytes: [...exported.bytes] };
  });
  expect(result.diagnostics).toEqual([]);
  expect(result.box.height).toBe(55);
  expect(result.box.width).toBeGreaterThan(40);
  expect(result.box.width).toBeLessThan(150);
  const xml = readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
  expect(xml).toContain(`<a:ext cx="${Math.round(result.box.width * 12700)}" cy="698500"/>`);
  expect(xml).toMatch(/<a:bodyPr[^>]* vert="eaVert"/);
});

test("East Asian cell auto row fits browser-wrapped columns under a height cap", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned International", heading: "Pinned International" } } });
    deck.addSlide().addElement({ type: "table", id: "t", box: { x: 20, y: 20, width: 150, height: 20 }, columns: [150], rows: [
      { height: 20, cells: [{ style: { textDirection: "eastAsianVertical", margin: 5 },
        content: [{ style: { lineSpacing: 1.5 }, runs: [{ text: "中文ABCD", style: { fontSize: 20 } }] }] }] },
    ] });
    const layout = { textMeasurement: harness.adapter, tableSizing: [{ elementId: "t", autoRows: [{ row: 0, maxHeight: 55 }] }], diagnostics: { textOverflow: true } };
    const resolved = resolveLayout(deck, layout), table = resolved.slides[0].elements[0];
    const { bytes } = await generatePptx(deck, { layout });
    return { height: table.rows[0].height, diagnostics: resolved.analysis.diagnostics.map(d => d.code), bytes: [...bytes] };
  });
  expect(result.height).toBeGreaterThan(20);
  expect(result.height).toBeLessThanOrEqual(55);
  expect(result.diagnostics).toContain("table-text-direction-wrapped-fallback");
  expect(result.diagnostics).not.toContain("text-overflow");
  expect(result.diagnostics).not.toContain("table-cell-overflow");
  const xml = readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
  expect(xml).toContain(`<a:tr h="${Math.round(result.height * 12700)}">`);
  expect(xml).toMatch(/<a:tcPr[^>]* vert="eaVert"/);
});

test("Mongolian and stacked previews use the authored column progression", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { renderPresentationToSvg } = await import("@pptkit/svg-renderer");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    const slide = deck.addSlide();
    for (const [i, textDirection] of ["mongolianVertical", "stackedVertical", "stackedVerticalRtl"].entries()) {
      slide.addElement({ type: "text", box: { x: 20 + i * 110, y: 20, width: 100, height: 100 }, frame: { textDirection },
        content: ["ABCD", "EFGH"].map(text => ({ runs: [{ text, style: { fontSize: 18 } }] })) });
    }
    const svg = (await renderPresentationToSvg(deck)).slides[0].svg;
    const host = document.createElement("div");
    host.innerHTML = svg; document.body.appendChild(host);
    return Array.from(host.querySelectorAll('[style*="writing-mode:vertical-"]'), frame => {
      const columns = Array.from(frame.children, p => {
        const range = document.createRange();
        range.selectNodeContents(p);
        const rect = range.getBoundingClientRect();
        return { x: rect.x, width: rect.width, height: rect.height };
      });
      const style = getComputedStyle(frame);
      return { writingMode: style.writingMode, orientation: style.textOrientation, columns };
    });
  });
  expect(result.map(x => [x.writingMode, x.orientation])).toEqual([
    ["vertical-lr", "mixed"], ["vertical-lr", "upright"], ["vertical-rl", "upright"],
  ]);
  for (const [i, { columns }] of result.entries()) {
    expect(columns).toHaveLength(2);
    for (const column of columns) expect(column.height).toBeGreaterThan(column.width);
    if (i === 2) expect(columns[1].x).toBeLessThan(columns[0].x);
    else expect(columns[1].x).toBeGreaterThan(columns[0].x);
  }
});

test("built-in vertical bounds respect stacked orientation and the starting edge", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    const slide = deck.addSlide();
    for (const textDirection of ["mongolianVertical", "stackedVertical", "stackedVerticalRtl"]) {
      slide.addElement({ type: "text", id: textDirection, box: { x: 0, y: 0, width: 120, height: 160 },
        frame: { textDirection, margin: 5, wrap: false },
        content: [{ style: { lineSpacing: 1.5 }, runs: [{ text: "ABCD", style: { fontSize: 20 } }] }] });
    }
    const analysis = resolveLayout(deck, { textMeasurement: harness.adapter, diagnostics: { textOverflow: true } }).analysis;
    return { ...analysis, advance: harness.width("ABCD", 20) };
  });
  expect(result.diagnostics).toEqual([]);
  expect(result.measurements).toHaveLength(3);
  const [mixed, left, right] = result.measurements.map(m => m.bounds);
  expect(mixed.height).toBeCloseTo(result.advance, 1);
  expect(left.height).toBeGreaterThan(mixed.height);
  expect(right.height).toBeCloseTo(left.height, 1);
  expect(left.x).toBeCloseTo(5, 1);
  expect(right.x + right.width).toBeCloseTo(115, 1);
});

test("remaining vertical modes solve bounded text and table geometry with loaded fonts", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout } = await import("@pptkit/layout");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned Mono", heading: "Pinned Mono" } } });
    const slide = deck.addSlide(), intrinsicSizing = [], tableSizing = [];
    for (const [i, textDirection] of ["mongolianVertical", "stackedVertical", "stackedVerticalRtl"].entries()) {
      const content = [{ style: { lineSpacing: 1.5 }, runs: [{ text: "ABCD", style: { fontSize: 20 } }] }];
      const box = { x: 20 + 155 * i, y: 20, width: 150, height: 20 };
      slide.addElement({ type: "text", id: "v" + i, box, content, frame: { textDirection, margin: 5 } });
      slide.addElement({ type: "table", id: "t" + i, box: { ...box, y: 100 }, columns: [150], rows: [
        { height: 20, cells: [{ content, style: { textDirection, margin: 5 } }] },
      ] });
      intrinsicSizing.push({ elementId: "v" + i, width: { type: "content" }, height: { type: "content", max: 55 } });
      tableSizing.push({ elementId: "t" + i, autoRows: [{ row: 0, maxHeight: 55 }] });
    }
    const layout = { textMeasurement: harness.adapter, intrinsicSizing, tableSizing, diagnostics: { textOverflow: true } };
    const resolved = resolveLayout(deck, layout);
    const bytes = (await generatePptx(deck, { layout })).bytes;
    return { boxes: resolved.slides[0].elements.map(e => e.box), diagnostics: resolved.analysis.diagnostics.map(d => d.code), bytes: [...bytes] };
  });
  expect(result.diagnostics.filter(d => d === "table-text-direction-wrapped-fallback")).toHaveLength(3);
  expect(result.diagnostics.filter(d => /overflow|unavailable|unsatisfiable/.test(d))).toEqual([]);
  const xml = readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
  const rows = [...xml.matchAll(/<a:tr h="(\d+)">/g)].map(m => Number(m[1]));
  for (const [i, mode] of ["mongolianVert", "wordArtVert", "wordArtVertRtl"].entries()) {
    const frame = result.boxes[2 * i], table = result.boxes[2 * i + 1];
    expect(frame.height).toBe(55);
    expect(frame.width).toBeGreaterThan(40);
    expect(frame.width).toBeLessThan(150);
    expect(table.height).toBeGreaterThan(20);
    expect(table.height).toBeLessThanOrEqual(55);
    expect(rows[i]).toBe(Math.round(table.height * 12700));
    expect(xml).toMatch(new RegExp(`<a:bodyPr[^>]* vert="${mode}"`));
    expect(xml).toMatch(new RegExp(`<a:tcPr[^>]* vert="${mode}"`));
  }
});

for (const axis of ["both", "width", "height"]) {
  test(`unwrapped vertical parallelogram fits its inset and exports the solved ${axis}`, async ({ page }) => {
    const result = await page.evaluate(async (axis) => {
      const { createPresentation } = await import("@pptkit/core");
      const { resolveLayout, resolveShapeTextBox } = await import("@pptkit/layout");
      const { generatePptx } = await import("@pptkit/pptx-exporter");
      const deck = createPresentation({ theme: { fonts: { body: "Pinned International", heading: "Pinned International" } } });
      deck.addSlide().addElement({ type: "shape", shape: "parallelogram", slant: 0.4, id: "p", box: { x: 20, y: 20, width: 180, height: 200 },
        text: { frame: { textDirection: "eastAsianVertical", wrap: false, margin: 5 },
          content: [{ style: { lineSpacing: 1.5 }, runs: [{ text: "中文ABCD", style: { fontSize: 20 } }] }] } });
      const layout = { textMeasurement: harness.adapter, intrinsicSizing: [{ elementId: "p", ...(axis !== "height" ? { width: { type: "content" } } : {}),
        ...(axis !== "width" ? { height: { type: "content" } } : {}) }],
        diagnostics: { textOverflow: true } };
      const resolved = resolveLayout(deck, layout), shape = resolved.slides[0].elements[0];
      const { bytes } = await generatePptx(deck, { layout });
      return { box: shape.box, inset: resolveShapeTextBox(shape), analysis: resolved.analysis, bytes: [...bytes] };
    }, axis);
    if (axis === "width") expect(result.box.height).toBe(200);
    if (axis === "height") expect(result.box.width).toBe(180);
    expect(result.analysis.diagnostics).toEqual([]);
    const measured = result.analysis.measurements[0].bounds;
    expect(measured.height).toBeGreaterThan(80);
    expect(measured.height + 10).toBeLessThanOrEqual(result.inset.height + 0.02);
    expect(measured.width + 10).toBeLessThanOrEqual(result.inset.width + 0.02);
    const xml = readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
    expect(xml).toContain(`<a:ext cx="${Math.round(result.box.width * 12700)}" cy="${Math.round(result.box.height * 12700)}"/>`);
    expect(xml).toMatch(/<a:bodyPr[^>]* vert="eaVert"/);
  });
}

test("wrapped East Asian parallelogram fits a height cap and exports the verified frame", async ({ page }) => {
  const result = await page.evaluate(async () => {
    const { createPresentation } = await import("@pptkit/core");
    const { resolveLayout, resolveShapeTextBox } = await import("@pptkit/layout");
    const { generatePptx } = await import("@pptkit/pptx-exporter");
    const deck = createPresentation({ theme: { fonts: { body: "Pinned International", heading: "Pinned International" } } });
    deck.addSlide().addElement({ type: "shape", shape: "parallelogram", slant: 0.4, id: "p", box: { x: 20, y: 20, width: 230, height: 240 },
      text: { frame: { textDirection: "eastAsianVertical", wrap: true, margin: 5 },
        content: [{ style: { lineSpacing: 1.5 }, runs: [{ text: "中文中文中文中文中文中文", style: { fontSize: 20 } }] }] } });
    const layout = { textMeasurement: harness.adapter, intrinsicSizing: [{ elementId: "p", width: { type: "content", max: 220 }, height: { type: "content", max: 115 } }],
      diagnostics: { textOverflow: true } };
    const resolved = resolveLayout(deck, layout), shape = resolved.slides[0].elements[0];
    const { bytes } = await generatePptx(deck, { layout });
    return { box: shape.box, inset: resolveShapeTextBox(shape), analysis: resolved.analysis, bytes: [...bytes] };
  });
  expect(result.analysis.diagnostics).toEqual([]);
  expect(result.box.height).toBeLessThanOrEqual(115);
  expect(result.box.width).toBeLessThanOrEqual(220);
  const measured = result.analysis.measurements[0].bounds;
  expect(measured.width).toBeGreaterThan(40);
  expect(measured.height + 10).toBeLessThanOrEqual(result.inset.height + 0.02);
  expect(measured.width + 10).toBeLessThanOrEqual(result.inset.width + 0.02);
  const xml = readGeneratedZip(Uint8Array.from(result.bytes)).get("ppt/slides/slide1.xml").toString();
  expect(xml).toContain(`<a:ext cx="${Math.round(result.box.width * 12700)}" cy="${Math.round(result.box.height * 12700)}"/>`);
  expect(xml).toMatch(/<a:bodyPr[^>]* vert="eaVert"/);
});
