import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import { importPptx, DEFAULT_SESSION_LIMITS, PptxImportError } from "../dist/index.js";
import { chartFixture, chartXml, series, workbook, encodeChart, frame } from "./chart-fixture.mjs";
import { zip } from "./zip-fixture.mjs";
import { relationships } from "./opc-fixture.mjs";
import { shape, body } from "./content-fixture.mjs";
import { readGeneratedZip } from "../../../tests/support/read-generated-zip.mjs";
import { readOpcSource } from "../dist/opc/read-package.js";
import { prepareTextCandidate } from "../dist/import/prepare-text-candidate.js";
import { snapshot } from "../dist/import/state.js";
import { prepareChartValues } from "../dist/import/prepare-chart-values.js";
import { chartNumber, chartNumericSourcesAgree } from "../dist/import/chart-data.js";
const config = { projection: "content", chartEditSaveProfile: "existing-values-v1" };
const fixture = ({ family = "bar", strict = false, ...extra } = {}) => chartFixture({ strict, xml: chartXml({ family, data: series({ refs: true, values: family === "pie" ? [0, 2.5] : [0, -2.5] }), external: true }), book: workbook({ values: family === "pie" ? [0, 2.5] : [0, -2.5] }), ...extra });
function change(session, edit) { const slide = session.document.slides[0], index = slide.elements.findIndex(e => e.type === "chart"), chart = structuredClone(slide.elements[index]); edit(chart); slide.removeElement(chart.id); slide.insertElement(index, chart); }
const save = async s => { const prepared = await s.prepareSave({ mode: "edit" }); assert.equal(prepared.status, "ready", JSON.stringify(prepared.report)); return s.save(prepared.plan); };

test('shared category cache research fixture imports escaped labels but blocks label source edits', async () => {
  const input = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const session = await importPptx(input, config);
  try {
    assert.equal(session.report.chartLabels, undefined);
    const doc = session.document, slide = doc.slides[0], chart = slide.elements[0];
    assert.deepEqual(chart.categories, [' Q1 ', 'literal _x0041_', '中文 & two']);
    assert.deepEqual(chart.series.map(series => series.name), ['First _x005F_', 'Second & series']);
    const inverse = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
      { type: 'chart-category', slideId: slide.id, elementId: chart.id, categoryIndex: 1, expected: 'literal _x0041_', value: 'Changed' },
    ] }).inverse;
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
    doc.applyEditOperation(inverse); const noop = await session.prepareSave({ mode: 'no-op' }); assert.equal(noop.status, 'ready');
    assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(input));
  } finally { session.dispose(); }
});

test('v4 saves shared category caches with numeric edits across three independent reopens', async () => {
  let bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const v4 = { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' };
  for (const [index, value] of ['First & changed', ' literal _x0041_ ', '中文 😀'].entries()) {
    const before = readGeneratedZip(bytes), session = await importPptx(bytes, v4);
    try {
      const labels = session.report.chartLabels; assert.equal(labels.length, 3); assert.ok(Object.isFrozen(labels) && Object.isFrozen(labels[1].sources));
      assert.equal(labels[1].status, 'bound'); assert.equal(labels[1].sources.length, 2);
      assert.ok(labels[1].sources.every(source => source.workbook.cellReference === 'A3' && source.workbook.preservesWhitespace));
      assert.notEqual(labels[1].sources[0].cache.elementIndex, labels[1].sources[1].cache.elementIndex);
      assert.ok(!JSON.stringify(labels).includes('lexical'));
      change(session, chart => { chart.categories[1] = value; chart.series[0].values[0] = 20 + index; });
      const saved = await save(session), after = readGeneratedZip(saved.bytes);
      assert.equal(saved.report.changes.find(item => item.chartLabel)?.chartLabel.after, value);
      for (const [name, content] of before) if (!['ppt/charts/chart1.xml', 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'].includes(name)) assert.deepEqual(after.get(name), content, name);
      const reopened = await importPptx(saved.bytes, v4);
      try { assert.equal(reopened.document.slides[0].elements[0].categories[1], value); assert.equal(reopened.document.slides[0].elements[0].series[0].values[0], 20 + index); } finally { reopened.dispose(); }
      bytes = saved.bytes;
    } finally { session.dispose(); }
  }
});

test('v4 category save works in both chart namespace dialects and blocks unknown references', async () => {
  const v4 = { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' };
  for (const strict of [false, true]) {
    const session = await importPptx(encodeChart(fixture({ strict })), v4);
    try {
      change(session, chart => { chart.categories[0] = 'updated'; });
      const saved = await save(session), reopened = await importPptx(saved.bytes, v4);
      try { assert.equal(reopened.document.slides[0].elements[0].categories[0], 'updated'); } finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
  const source = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const chart = readGeneratedZip(source).get('ppt/charts/chart1.xml').toString().replace('</c:chartSpace>', '<c:extLst><c:ext uri="future"><c:f>Sheet1!$A$3</c:f></c:ext></c:extLst></c:chartSpace>');
  const { replaceArchiveEntries } = await import('../dist/preservation/replace-archive-entries.js');
  const altered = await replaceArchiveEntries(source, [{ name: 'ppt/charts/chart1.xml', bytes: new TextEncoder().encode(chart) }], { maxOutputBytes: 10_000_000 });
  const session = await importPptx(altered, v4);
  try { assert.ok(session.report.chartLabels.every(item => item.status === 'unmapped' && item.reason)); change(session, item => { item.categories[1] = 'unsafe'; }); const result = await session.prepareSave({ mode: 'edit' }); assert.equal(result.status, 'blocked'); assert.equal(result.plan, undefined); } finally { session.dispose(); }
});

test('v4 preserves boundary whitespace by inserting or upgrading xml:space on the bound worksheet text', async () => {
  const { replaceArchiveEntries } = await import('../dist/preservation/replace-archive-entries.js');
  const original = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const bookName = 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx', sheetName = 'xl/worksheets/sheet1.xml';
  const book = readGeneratedZip(original).get(bookName), sheet = readGeneratedZip(book).get(sheetName).toString();
  const target = '<t xml:space="preserve">literal _x005F_x0041_</t>'; assert.ok(sheet.includes(target));
  for (const attribute of ['', ' xml:space="default"']) {
    const changedSheet = sheet.replace(target, `<t${attribute}>literal _x005F_x0041_</t>`);
    const changedBook = await replaceArchiveEntries(book, [{ name: sheetName, bytes: new TextEncoder().encode(changedSheet) }], { maxOutputBytes: 10_000_000 });
    let bytes = await replaceArchiveEntries(original, [{ name: bookName, bytes: changedBook }], { maxOutputBytes: 10_000_000 });
    const v4 = { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' };
    for (const value of [' leading ', ' trailing ', ' third ']) {
      const session = await importPptx(bytes, v4);
      try {
        change(session, chart => { chart.categories[1] = value; }); const saved = await save(session); bytes = saved.bytes;
        const outputSheet = readGeneratedZip(readGeneratedZip(bytes).get(bookName)).get(sheetName).toString();
        assert.ok(outputSheet.includes(`<t xml:space="preserve">${value}</t>`));
        const reopened = await importPptx(bytes, v4); try { assert.equal(reopened.document.slides[0].elements[0].categories[1], value); assert.ok(reopened.report.chartLabels[1].sources.every(source => source.workbook.preservesWhitespace)); } finally { reopened.dispose(); }
      } finally { session.dispose(); }
    }
  }
});

test('v4 category saves enforce cancellation, stale plans and combined change limits', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const v4 = { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' };
  const session = await importPptx(bytes, v4);
  try {
    change(session, chart => { chart.categories[1] = 'new'; });
    const aborted = new AbortController(); aborted.abort(); await assert.rejects(session.prepareSave({ mode: 'edit', signal: aborted.signal }), { code: 'operation-aborted' });
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready');
    change(session, chart => { chart.categories[1] = 'changed again'; });
    await assert.rejects(session.save(plan.plan), { code: 'save-snapshot-stale' }); plan.plan.dispose();
    assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'ready');
  } finally { session.dispose(); }
  const limited = await importPptx(bytes, { ...v4, sessionLimits: { maxSaveChanges: 1 } });
  try {
    change(limited, chart => { chart.categories[1] = 'new'; chart.series[0].values[0] = 55; });
    await assert.rejects(limited.prepareSave({ mode: 'edit' }), { code: 'save-limit-exceeded' });
  } finally { limited.dispose(); }
});

test('v5 saves series names, categories and numeric values together across reopens', async () => {
  let bytes = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const v5 = { projection: 'content', chartEditSaveProfile: 'existing-values-and-series-names-v5' };
  for (const [round, value] of ['changed & name', ' name _x0041_ ', '名称 😀'].entries()) {
    const session = await importPptx(bytes, v5);
    try {
      assert.deepEqual(session.report.chartSeriesNames.map(item => item.status), ['bound', 'bound']);
      change(session, chart => { chart.series[0].name = value; chart.categories[1] = `category ${round}`; chart.series[0].values[0] = 70 + round; });
      const saved = await save(session); assert.equal(saved.report.changes.find(item => item.chartSeriesName)?.chartSeriesName.after, value); bytes = saved.bytes;
      const reopened = await importPptx(bytes, v5); try { const chart = reopened.document.slides[0].elements[0]; assert.equal(chart.series[0].name, value); assert.equal(chart.categories[1], `category ${round}`); assert.equal(chart.series[0].values[0], 70 + round); } finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});

test('v5 series-name saves cover scatter and bubble while v4 rejects them', async () => {
  for (const family of ['scatter', 'bubble']) {
    const bytes = await readFile(new URL(`../../../tests/fixtures/${family}-name-source.pptx`, import.meta.url));
    const old = await importPptx(bytes, { projection: 'content', chartEditSaveProfile: 'existing-values-and-labels-v4' });
    try { change(old, chart => { chart.series[0].name = 'new name'; }); assert.equal((await old.prepareSave({ mode: 'edit' })).status, 'blocked'); } finally { old.dispose(); }
    const session = await importPptx(bytes, { projection: 'content', chartEditSaveProfile: 'existing-values-and-series-names-v5' });
    try {
      assert.equal(session.report.chartSeriesNames[0].status, 'bound'); change(session, chart => { chart.series[0].name = 'new & name'; });
      const saved = await save(session), reopened = await importPptx(saved.bytes, { projection: 'content' });
      try { assert.equal(reopened.document.slides[0].elements[0].series[0].name, 'new & name'); } finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});

test('v5 blocks aliased series names without partially saving another edit', async () => {
  const { replaceArchiveEntries } = await import('../dist/preservation/replace-archive-entries.js');
  const original = await readFile(new URL('../../../tests/fixtures/chart-label-source.pptx', import.meta.url));
  const name = 'ppt/charts/chart1.xml', chart = readGeneratedZip(original).get(name).toString();
  const altered = chart.replace('Sheet1!$C$1', 'Sheet1!$B$1'); assert.notEqual(altered, chart);
  const bytes = await replaceArchiveEntries(original, [{ name, bytes: new TextEncoder().encode(altered) }], { maxOutputBytes: 10_000_000 });
  const session = await importPptx(bytes, { projection: 'content', chartEditSaveProfile: 'existing-values-and-series-names-v5' });
  try {
    assert.ok(session.report.chartSeriesNames.every(item => item.status === 'unmapped'));
    change(session, item => { item.series[0].name = 'not safe'; item.categories[1] = 'otherwise safe'; });
    const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'blocked'); assert.equal(plan.plan, undefined);
  } finally { session.dispose(); }
});

test('v5 rebinds series names after a retained PowerPoint resave and can save again', async () => {
  const bytes = await readFile(new URL('../../../tests/fixtures/chart-series-name-powerpoint-resave.pptx', import.meta.url));
  const options = { projection: 'content', chartEditSaveProfile: 'existing-values-and-series-names-v5' }, session = await importPptx(bytes, options);
  try {
    assert.deepEqual(session.report.chartSeriesNames.map(item => item.status), ['bound', 'bound']);
    assert.equal(session.report.chartSeriesNames[0].sources[0].workbook.cellReference, 'B1');
    const noOp = await session.prepareSave({ mode: 'no-op' }); assert.equal(noOp.status, 'ready'); assert.deepEqual(Buffer.from((await session.save(noOp.plan)).bytes), Buffer.from(bytes));
    change(session, chart => { chart.series[0].name = 'After PowerPoint'; chart.series[0].values[0] = 12; });
    const saved = await save(session), reopened = await importPptx(saved.bytes, options);
    try { const chart = reopened.document.slides[0].elements[0]; assert.equal(chart.series[0].name, 'After PowerPoint'); assert.equal(chart.series[0].values[0], 12); assert.equal(reopened.report.chartSeriesNames[0].status, 'bound'); } finally { reopened.dispose(); }
  } finally { session.dispose(); }
});

test('underflow cannot turn a nonzero chart value into a writable zero', async () => {
  for(const value of ['1e-400','-1e-400','0.0001e-9999','2e-324'])assert.equal(chartNumber(value),undefined,value);
  for(const value of ['0','-0','0.000e-9999','-0e-9999'])assert.ok(Object.is(chartNumber(value),value.startsWith('-')?-0:0),value);
  assert.equal(chartNumber('5e-324'),Number.MIN_VALUE);
  for(const strict of [false,true]){
    const mismatch=fixture({strict,book:workbook({values:['1e-400',-2.5]})}),s=await importPptx(encodeChart(mismatch),config);
    try{assert.equal(s.document.slides[0].elements[0].series[0].values[0],0);assert.ok(s.report.chartValues.every(item=>item.status==='unmapped'));change(s,c=>{c.series[0].values[0]=1;});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);}finally{s.dispose();}
    const underflow=fixture({strict,xml:chartXml({data:series({refs:true,values:['1e-400',-2.5]}),external:true}),book:workbook({values:['1e-400',-2.5]})}),original=encodeChart(underflow),unmapped=await importPptx(original,config);
    try{assert.ok(!unmapped.document.slides[0].elements.some(element=>element.type==='chart'));const no=await unmapped.prepareSave({mode:'no-op'});assert.equal(no.status,'ready');assert.deepEqual(Buffer.from((await unmapped.save(no.plan)).bytes),Buffer.from(original));}finally{unmapped.dispose();}
    const axis=fixture({strict});axis.parts['deck/chart.xml']=axis.parts['deck/chart.xml'].replace('<c:min val="-5"/>','<c:min val="1e-400"/>');const axisBytes=encodeChart(axis),axisSession=await importPptx(axisBytes,config);try{assert.ok(!axisSession.document.slides[0].elements.some(element=>element.type==='chart'));const no=await axisSession.prepareSave({mode:'no-op'});assert.equal(no.status,'ready');assert.deepEqual(Buffer.from((await axisSession.save(no.plan)).bytes),Buffer.from(axisBytes));}finally{axisSession.dispose();}
  }
});

test('different decimal sources that round to the same double cannot bind', async () => {
  for(const [a,b] of [['9007199254740992','9007199254740993'],['0.1','0.10000000000000001'],['5e-324','4e-324']]){assert.ok(Object.is(Number(a),Number(b)));assert.equal(chartNumericSourcesAgree(a,b),false,`${a} vs ${b}`);}
  for(const [a,b] of [['1.00','1e0'],['-2.50e0','-2.5'],['0.1000','1e-1'],['+0e9999','0'],['-0.0','-0e-9999'],['9007199254740993','9.007199254740993e15']])assert.equal(chartNumericSourcesAgree(a,b),true,`${a} vs ${b}`);
  for(const [a,b] of [['-0','0'],['1e-400','1e-400'],['NaN','NaN']])assert.equal(chartNumericSourcesAgree(a,b),false,`${a} vs ${b}`);
  for(const strict of [false,true]){
    const f=fixture({strict,xml:chartXml({data:series({refs:true,values:['9007199254740992',-2.5]}),external:true}),book:workbook({values:['9007199254740993',-2.5]})}),bytes=encodeChart(f),s=await importPptx(bytes,config);
    try{assert.ok(s.report.chartValues.every(item=>item.status==='unmapped'));assert.ok(s.report.chartValues.every(item=>/agree/.test(item.reason)));change(s,c=>{c.series[0].values[1]=7;});const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);}finally{s.dispose();}
    const equal=fixture({strict,xml:chartXml({data:series({refs:true,values:['1.00',-2.5]}),external:true}),book:workbook({values:['1e0',-2.5]})}),session=await importPptx(encodeChart(equal),config);try{assert.ok(session.report.chartValues.every(item=>item.status==='bound'));change(session,c=>{c.series[0].values[0]=2;});const p=await session.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const reopened=await importPptx((await session.save(p.plan)).bytes,config);try{assert.equal(reopened.document.slides[0].elements[0].series[0].values[0],2);}finally{reopened.dispose();}}finally{session.dispose();}
  }
});

test('numeric source agreement matches independent integer arithmetic across the safe-integer boundary', () => {
  const base=9007199254740900n;
  for(let i=0n;i<200n;i++){
    const value=base+i, same=`${value*10n}e-1`, next=value+1n;
    assert.equal(chartNumericSourcesAgree(String(value),same),true,String(value));
    assert.equal(chartNumericSourcesAgree(String(value),String(next)),false,`${value} vs ${next}`);
    assert.equal(chartNumericSourcesAgree(`-${value}`,`-${value*100n}e-2`),true,String(value));
  }
});

test("bar/line/pie and both dialects synchronize three numeric rounds while retaining all unrelated outer/inner payloads", async () => {
  for (const family of ["bar", "line", "pie"]) for (const strict of [false, true]) {
    const f = fixture({ family, strict }); let bytes = encodeChart(f);
    for (const [round, value] of [12.25, 0, 37].entries()) {
      const before = readGeneratedZip(bytes), session = await importPptx(bytes, config);
      try {
        assert.ok(session.report.chartValues.every(v => v.status === "bound"));
        assert.equal(session.report.chartValues[1].workbook.cellReference, "B3"); change(session, c => { c.series[0].values[1] = value; });
        const saved = await save(session), after = readGeneratedZip(saved.bytes); assert.equal(saved.report.changes.length, 1); assert.equal(saved.report.changes[0].chartValue.after, value);
        assert.deepEqual([...after.keys()], [...before.keys()]);
        for (const [name, data] of before) if (!["deck/chart.xml", "deck/book.xlsx"].includes(name)) assert.deepEqual(after.get(name), data, name);
        assert.equal(after.get("deck/chart.xml").toString(), before.get("deck/chart.xml").toString().replace(/(<c:val>[\s\S]*?<c:pt idx="1"><c:v>)[^<]*(<\/c:v>)/, `$1${value}$2`));
        const a = readGeneratedZip(before.get("deck/book.xlsx")), b = readGeneratedZip(after.get("deck/book.xlsx")); assert.deepEqual([...a.keys()], [...b.keys()]);
        for (const [name, data] of a) if (name !== "xl/sheets/one.xml") assert.deepEqual(b.get(name), data, name);
        assert.equal(b.get("xl/sheets/one.xml").toString(), a.get("xl/sheets/one.xml").toString().replace(/(<c r="B3"[^>]*><v>)[^<]*(<\/v>)/, `$1${value}$2`));
        const reopened = await importPptx(saved.bytes, config); try { assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], value); assert.ok(reopened.report.chartValues.every(v => v.status === "bound")); } finally { reopened.dispose(); }
        bytes = saved.bytes; void round;
      } finally { session.dispose(); }
    }
  }
});

test('chart point operations keep cache and workbook synchronized across three source saves', async () => {
  for (const strict of [false, true]) {
    let input = encodeChart(fixture({ strict }));
    for (const value of [12.25, '-0', 8]) {
      const before = readGeneratedZip(input), session = await importPptx(input, config);
      try {
        const doc = session.document, slide = doc.slides[0], chart = slide.elements[0], current = chart.series[0].values[1];
        const expected = Object.is(current, -0) ? '-0' : current;
        const undo = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
          { type: 'chart-point', slideId: slide.id, elementId: chart.id, seriesIndex: 0, pointIndex: 1, coordinate: 'values', expected, value },
        ] }).inverse;
        const ready = await session.prepareSave({ mode: 'edit' }); assert.equal(ready.status, 'ready', JSON.stringify(ready.report));
        const redo = doc.applyEditOperation(undo).inverse; await assert.rejects(session.save(ready.plan), error => error.code === 'save-snapshot-stale');
        doc.applyEditOperation(JSON.parse(JSON.stringify(redo))); const saved = await save(session), after = readGeneratedZip(saved.bytes);
        for (const [name, payload] of before) if (!['deck/chart.xml', 'deck/book.xlsx'].includes(name)) assert.deepEqual(after.get(name), payload, name);
        const lexical = value === '-0' ? '-0' : String(value);
        assert.match(after.get('deck/chart.xml').toString(), new RegExp(`<c:pt idx="1"><c:v>${lexical}</c:v>`));
        const workbookParts = readGeneratedZip(after.get('deck/book.xlsx'));
        assert.match(workbookParts.get('xl/sheets/one.xml').toString(), new RegExp(`<c r="B3"[^>]*><v>${lexical}</v>`));
        const reopened = await importPptx(saved.bytes, config);
        try { assert.ok(Object.is(reopened.document.slides[0].elements[0].series[0].values[1], value === '-0' ? -0 : value)); }
        finally { reopened.dispose(); }
        input = saved.bytes;
      } finally { session.dispose(); }
    }
    const bad = fixture({ strict, book: workbook({ values: [8, -2.5] }) }), source = encodeChart(bad), session = await importPptx(source, config);
    try {
      const doc = session.document, slide = doc.slides[0], chart = slide.elements[0];
      const inverse = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [
        { type: 'chart-point', slideId: slide.id, elementId: chart.id, seriesIndex: 0, pointIndex: 1, coordinate: 'values', expected: -2.5, value: 4 },
      ] }).inverse;
      assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked'); doc.applyEditOperation(inverse);
      const noop = await session.prepareSave({ mode: 'no-op' }); assert.equal(noop.status, 'ready'); assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(source));
    } finally { session.dispose(); }
  }
});

test('chart label operations cannot bypass numeric-only source-save profile', async () => {
  for (const strict of [false, true]) {
    const source = encodeChart(fixture({ strict })), session = await importPptx(source, config);
    try {
      const doc = session.document, slide = doc.slides[0], chart = slide.elements[0], target = { slideId: slide.id, elementId: chart.id };
      const numeric = { type: 'chart-point', ...target, seriesIndex: 0, pointIndex: 1, coordinate: 'values', expected: -2.5, value: 9 };
      const label = { type: 'chart-category', ...target, categoryIndex: 0, expected: chart.categories[0], value: 'Changed' };
      const inverse = doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [numeric, label] }).inverse;
      assert.equal((await session.prepareSave({ mode: 'edit' })).status, 'blocked');
      doc.applyEditOperation(inverse);
      const noop = await session.prepareSave({ mode: 'no-op' }); assert.equal(noop.status, 'ready'); assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(source));
      doc.applyEditOperation({ schemaVersion: 'pptkit-edit-v1', documentId: doc.id, changes: [numeric] });
      const saved = await save(session), reopened = await importPptx(saved.bytes, config);
      try { assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], 9); assert.equal(reopened.document.slides[0].elements[0].categories[0], chart.categories[0]); }
      finally { reopened.dispose(); }
    } finally { session.dispose(); }
  }
});

test("noncanonical numbers, signed zero, unknown chart neighbors and numeric cell formatting retain exact lexical surroundings", async () => {
  const f = fixture(); f.parts["deck/chart.xml"] = f.parts["deck/chart.xml"].replace('>-2.5</c:v>', '> -2.50e0 </c:v>').replace('</c:chartSpace>', '<c:extLst><c:ext uri="urn:keep"><x:keep xmlns:x="urn:keep"><c:v>-2.5</c:v></x:keep></c:ext></c:extLst></c:chartSpace>');
  const bookParts = readGeneratedZip(f.parts["deck/book.xlsx"]);
  bookParts.set("xl/sheets/one.xml", Buffer.from(bookParts.get("xl/sheets/one.xml").toString().replace('<c r="B3" >', '<c r="B3" s="0" >')));
  f.parts["deck/book.xlsx"] = zip([...bookParts].map(([name, data]) => ({ name, body: data }))).bytes;
  const bytes = encodeChart(f), s = await importPptx(bytes, config);
  try { change(s, c => { c.series[0].values[1] = -0; }); const output = await save(s); const outer = readGeneratedZip(output.bytes);
    assert.ok(outer.get("deck/chart.xml").toString().includes('<c:v>-2.5</c:v></x:keep>')); assert.ok(outer.get("deck/chart.xml").toString().includes('<c:v>-0</c:v>'));
    const savedBook = readGeneratedZip(outer.get("deck/book.xlsx")); assert.ok(savedBook.get("xl/sheets/one.xml").toString().includes('<c r="B3" s="0" ><v>-0</v>'));
    const r = await importPptx(output.bytes, config); try { assert.ok(Object.is(r.document.slides[0].elements[0].series[0].values[1], -0)); assert.ok(r.report.chartValues.every(v => v.status === "bound")); } finally { r.dispose(); }
  } finally { s.dispose(); }
});

test("cache conflict/missing/literal, formula cells, auto-update, external data and opaque scalars remain unmapped and block writes", async () => {
  const cases = [fixture({ book: workbook({ values: [8, -2.5] }) }), fixture({ book: workbook({ formula: true }) }),
    fixture({ xml: chartXml({ data: series({ refs: true, omit: true }), external: true }) }), fixture({ xml: chartXml({ data: series(), external: true }) }),
    fixture({ book: undefined, external: true }), fixture({ xml: chartXml({ data: series({ refs: true }), external: true }).replace('autoUpdate val="0"', 'autoUpdate val="1"') }),
    fixture({ book: workbook({ values: [0], categories: ["Q1"] }) }),
    fixture({ xml: chartXml({ data: series({ refs: true }), external: true }).replace('<c:v>-2.5</c:v>', '<c:v><!-- opaque -->-2.5</c:v>') }),
    fixture({ xml: chartXml({ data: series({ refs: true }), external: true }).replace('</c:chartSpace>', '<c:extLst><c:ext uri="future"><c:f>Sheet1!$B$2</c:f></c:ext></c:extLst></c:chartSpace>') }),
  ];
  for (const f of cases) {
    const original = encodeChart(f), s = await importPptx(original, config);
    try { assert.ok(s.report.chartValues.some(v => v.status === "unmapped")); change(s, c => { c.series[0].values[1] = 5; }); const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "blocked"); assert.ok(p.report.diagnostics.some(d => d.code === "save-edit-unsupported")); assert.equal(p.plan, undefined); }
    finally { s.dispose(); }
  }
});

test("multiple frames sharing one relationship and repeated numeric cell references reject without choosing an owner", async () => {
  for (const f of [fixture({ content: frame({ id: 2 }) + frame({ id: 3 }) }), fixture({ xml: chartXml({ data: series({ idx: 1, order: 0, refs: true }) + series({ idx: 7, order: 1, refs: true }), external: true }) })]) {
    const s = await importPptx(encodeChart(f), config);
    try { assert.ok(s.report.chartValues.every(v => v.status === "unmapped")); assert.ok(s.report.chartValues.every(v => /[Ss]hared|multiple/.test(v.reason))); change(s, c => { c.series[0].values[0] = 9; }); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); }
    finally { s.dispose(); }
  }
});

test("mixed allowed text/numeric edits form one candidate; illegal styles/categories/topology block the complete plan", async () => {
  const bytes = encodeChart(fixture({ content: frame() + shape({ id: "20", txBox: true, text: body() }) }));
  const s = await importPptx(bytes, { ...config, editSaveProfile: "text-runs-v3" });
  try { change(s, c => { c.series[0].values[1] = 20; }); const slide = s.document.slides[0], text = structuredClone(slide.elements[1]); text.content[0].runs[0].text = "Mixed update"; slide.removeElement(text.id); slide.insertElement(1, text); const saved = await save(s); assert.equal(saved.report.changes.length, 2);
    const parts = readGeneratedZip(saved.bytes); assert.match(parts.get("deck/z.xml").toString(), /Mixed update/); assert.match(parts.get("deck/chart.xml").toString(), /<c:v>20<\/c:v>/);
  } finally { s.dispose(); }
  for (const edit of [c => { c.categories[0] = "changed"; }, c => { c.series[0].name = "changed"; }, c => { c.series[0].values.push(3); c.categories.push("new"); }, c => { c.box.width++; }]) {
    const s = await importPptx(bytes, config); try { change(s, c => { c.series[0].values[0] = 4; edit(c); }); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); } finally { s.dispose(); }
  }
});

test("default authority is unchanged; invocation capture, no-op, undo, cancellation/staleness, limits and retry remain scoped", async () => {
  const bytes = encodeChart(fixture()), unselected = await importPptx(bytes, { projection: "content" });
  try { assert.equal(unselected.report.chartValues, undefined); change(unselected, c => { c.series[0].values[0] = 1; }); assert.equal((await unselected.prepareSave({ mode: "edit" })).status, "blocked"); } finally { unselected.dispose(); }
  const options = { ...config }, pending = importPptx(bytes, options); options.chartEditSaveProfile = "invalid"; const s = await pending;
  try {
    const noOp = await s.prepareSave({ mode: "no-op" }); assert.deepEqual((await s.save(noOp.plan)).bytes, new Uint8Array(bytes));
    change(s, c => { c.series[0].values[0] = 7; }); change(s, c => { c.series[0].values[0] = 0; }); assert.equal((await s.prepareSave({ mode: "edit" })).report.diagnostics[0].code, "save-edit-no-changes");
    change(s, c => { c.series[0].values[0] = 7; }); const abort = new AbortController(), cancelled = s.prepareSave({ mode: "edit", signal: abort.signal }); abort.abort(); await assert.rejects(cancelled, { code: "operation-aborted" });
    const plan = await s.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready"); change(s, c => { c.series[0].values[1] = 8; }); await assert.rejects(s.save(plan.plan), { code: "save-snapshot-stale" }); plan.plan.dispose();
  } finally { s.dispose(); }
  const limited = await importPptx(bytes, { ...config, sessionLimits: { maxSaveChanges: 1 } });
  try { change(limited, c => { c.series[0].values = [3, 4]; }); await assert.rejects(limited.prepareSave({ mode: "edit" }), { code: "save-limit-exceeded" }); change(limited, c => { c.series[0].values[1] = -2.5; }); assert.equal((await limited.prepareSave({ mode: "edit" })).status, "ready"); } finally { limited.dispose(); }
  await assert.rejects(importPptx(bytes, { ...config, sessionLimits: { maxChartBindingWork: 1 } }), { code: "import-limit-exceeded" });
  for (const value of [false, true, "other"]) await assert.rejects(importPptx(bytes, { projection: "content", chartEditSaveProfile: value }), TypeError);
  await assert.rejects(importPptx(bytes, { projection: "structure", chartEditSaveProfile: "existing-values-v1" }), TypeError);
});

test("second-layer lexical failure and cancellation expose no partial staged chart/workbook replacement", async () => {
  const bytes = encodeChart(fixture()), source = await readOpcSource(bytes, {}, false, true), session = await importPptx(bytes, config), publicBinding = session.report.chartValues[1]; session.dispose();
  const binding = { ...publicBinding, baselineValue: -2.5, cacheLexical: "-2.5", cacheNamespace: "http://schemas.openxmlformats.org/drawingml/2006/chart", workbookLexical: "WRONG", workbookNamespace: "http://schemas.openxmlformats.org/spreadsheetml/2006/main" };
  const original = [...source.payloads].map(([k, v]) => [k, Buffer.from(v)]);
  await assert.rejects(prepareChartValues(source, [{ binding, value: 9 }], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0), e => e.code === "xml-edit-conflict");
  const abort = new AbortController(); let count = 0;
  await assert.rejects(prepareChartValues(source, [{ binding: { ...binding, workbookLexical: "-2.5" }, value: 9 }], { signal: abort.signal }, DEFAULT_SESSION_LIMITS, async () => { if (++count === 2) abort.abort(); if (abort.signal.aborted) throw new PptxImportError("operation-aborted", "test cancellation"); }, 0), { code: "operation-aborted" });
  for (const [k, v] of original) assert.deepEqual(source.payloads.get(k), new Uint8Array(v));
  const retry = await prepareChartValues(source, [{ binding: { ...binding, workbookLexical: "-2.5" }, value: 9 }], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0); assert.equal(retry.replacements.length, 2);
});

function withTwoCharts(sharedBook = false) {
  const f = fixture({ content: frame({ id: 2 }) + frame({ id: 3, ref: "second" }) });
  f.parts["deck/_rels/z.xml.rels"] = relationships([f.rel("layout", "slideLayout", "layout.xml"), f.rel("chart", "chart", "chart.xml"), f.rel("second", "chart", "second.xml")]);
  f.parts["deck/second.xml"] = f.parts["deck/chart.xml"];
  f.types["deck/second.xml"] = "drawingml.chart";
  f.parts["deck/_rels/second.xml.rels"] = relationships([f.rel("book", "package", sharedBook ? "book.xlsx" : "second.xlsx")]);
  if (!sharedBook) { f.parts["deck/second.xlsx"] = workbook(); f.types["deck/second.xlsx"] = "spreadsheetml.sheet"; }
  return f;
}
function rewriteBook(f, part, edit) {
  const parts = readGeneratedZip(f.parts["deck/book.xlsx"]);
  parts.set(part, Buffer.from(edit(parts.get(part).toString())));
  f.parts["deck/book.xlsx"] = zip([...parts].map(([name, data]) => ({ name, body: data }))).bytes;
  return f;
}

test("distinct chart parts sharing one workbook and numeric label aliases remain read-only", async () => {
  const alias = fixture({ xml: chartXml({ data: series({ name: "0", refs: true }).replace("Sheet1!$B$1", "Sheet1!$B$2"), external: true }), book: workbook({ name: "0" }) });
  for (const f of [withTwoCharts(true), alias,
    rewriteBook(fixture(), "xl/workbook.xml", xml => xml.replace("</workbook>", '<definedNames><definedName name="dependency">Sheet1!$B$2</definedName></definedNames></workbook>'))]) {
    const s = await importPptx(encodeChart(f), config);
    try { assert.ok(s.report.chartValues.some(v => v.status === "unmapped")); change(s, c => { c.series[0].values[0] = 8; }); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); }
    finally { s.dispose(); }
  }
});

test("multiple independent chart/workbook pairs synchronize atomically without name-based identity", async () => {
  const bytes = encodeChart(withTwoCharts()), s = await importPptx(bytes, config);
  try {
    assert.equal(s.report.chartValues.length, 4); assert.ok(s.report.chartValues.every(v => v.status === "bound"));
    const slide = s.document.slides[0];
    for (const [i, element] of [...slide.elements].entries()) { const chart = structuredClone(element); chart.series[0].values[1] = 17 + i; slide.removeElement(chart.id); slide.insertElement(i, chart); }
    const saved = await save(s), parts = readGeneratedZip(saved.bytes);
    assert.equal(saved.report.changes.length, 2);
    for (const [chart, book, value] of [["chart.xml", "book.xlsx", 17], ["second.xml", "second.xlsx", 18]]) {
      assert.ok(parts.get(`deck/${chart}`).toString().includes(`<c:v>${value}</c:v>`));
      assert.ok(readGeneratedZip(parts.get(`deck/${book}`)).get("xl/sheets/one.xml").toString().includes(`<v>${value}</v>`));
    }
    const reopened = await importPptx(saved.bytes, config); try { assert.deepEqual(reopened.document.slides[0].elements.map(e => e.series[0].values[1]), [17, 18]); assert.ok(reopened.report.chartValues.every(v => v.status === "bound")); } finally { reopened.dispose(); }
  } finally { s.dispose(); }
});

test("exact nested comparison exhaustion reports save-limit-exceeded instead of invalid limit arguments", async () => {
  const bytes = encodeChart(fixture()), source = await readOpcSource(bytes, {}, false, true), s = await importPptx(bytes, config);
  try {
    const binding = { ...s.report.chartValues[1], baselineValue: -2.5, cacheLexical: "-2.5", cacheNamespace: "http://schemas.openxmlformats.org/drawingml/2006/chart", workbookLexical: "-2.5", workbookNamespace: "http://schemas.openxmlformats.org/spreadsheetml/2006/main" };
    const staged = await prepareChartValues(source, [{ binding, value: 9 }], {}, DEFAULT_SESSION_LIMITS, async () => {}, 0);
    change(s, c => { c.series[0].values[1] = 9; });
    const bounded = await importPptx(bytes, { ...config, sessionLimits: { maxSaveValidationWork: staged.validationWork } });
    try { change(bounded, c => { c.series[0].values[1] = 9; }); await assert.rejects(bounded.prepareSave({ mode: "edit" }), { code: "save-limit-exceeded" }); } finally { bounded.dispose(); }
    assert.equal((await s.prepareSave({ mode: "edit" })).status, "ready");
  } finally { s.dispose(); }
});

test("candidate rebind rejects a wrong equal-valued worksheet target even when the chart/Core snapshot agrees", async () => {
  const bytes = encodeChart(fixture({ xml: chartXml({ data: series({ refs: true, values: [-2.5, -2.5] }), external: true }), book: workbook({ values: [-2.5, -2.5] }) }));
  const s = await importPptx(bytes, config);
  try {
    const [first, second] = s.report.chartValues;
    const binding = { ...first, workbook: second.workbook, baselineValue: -2.5, cacheLexical: "-2.5", cacheNamespace: "http://schemas.openxmlformats.org/drawingml/2006/chart", workbookLexical: "-2.5", workbookNamespace: "http://schemas.openxmlformats.org/spreadsheetml/2006/main" };
    change(s, c => { c.series[0].values[0] = 99; });
    const candidate = snapshot(s.document, DEFAULT_SESSION_LIMITS.maxSnapshotValues);
    const analysis = { changes: [], diagnostics: [], parts: new Map(), chartValues: [{ binding, value: 99 }] };
    await assert.rejects(prepareTextCandidate(bytes, candidate, analysis, {}, DEFAULT_SESSION_LIMITS, async () => {}), { code: "save-candidate-invalid" });
    // The public pipeline still uses the true binding and can be retried.
    assert.equal((await s.prepareSave({ mode: "edit" })).status, "ready");
  } finally { s.dispose(); }
});


test("starting-number, text and chart/workbook edits form one guarded candidate across dialects and rounds", async () => {
  for (const strict of [false, true]) {
    let bytes = encodeChart(fixture({ strict, content: frame() + shape({ id: "20", txBox: true, text: body() }) })), previous;
    for (const [round, number] of [17, 0, undefined].entries()) {
      const s = await importPptx(bytes, { ...config, editSaveProfile: "text-runs-v3", firstSlideNumberSaveProfile: "attribute-v2" });
      try {
        const before = readGeneratedZip(bytes), slide = s.document.slides[0];
        assert.equal(s.document.firstSlideNumber, previous); s.document.setFirstSlideNumber(number);
        change(s, c => { c.series[0].values[1] = round + 11; });
        const original = slide.elements[1], text = structuredClone(original); text.content[0].runs[0].text += "!"; slide.replaceElement(original.id, text);
        const bad = structuredClone(text); bad.box.x++; slide.replaceElement(original.id, bad); assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); slide.replaceElement(original.id, text);
        const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready", JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 3);
        const next = (await s.save(p.plan)).bytes, after = readGeneratedZip(next);
        for (const [name, data] of before) if (!["deck/main.data", "deck/z.xml", "deck/chart.xml", "deck/book.xlsx"].includes(name)) assert.deepEqual(after.get(name), data, name);
        const source = before.get("deck/main.data").toString(), output = after.get("deck/main.data").toString();
        if (previous === undefined) assert.equal(output.replace(` firstSlideNum="${number}"`, ""), source);
        else if (number === undefined) assert.equal(output, source.replace(`firstSlideNum="${previous}"`, ""));
        else assert.equal(output, source.replace(`firstSlideNum="${previous}"`, `firstSlideNum="${number}"`));
        assert.match(after.get("deck/chart.xml").toString(), new RegExp(`<c:v>${round + 11}</c:v>`));
        const sheet = readGeneratedZip(after.get("deck/book.xlsx")).get("xl/sheets/one.xml").toString(); assert.match(sheet, new RegExp(`<c r="B3"[^>]*><v>${round + 11}</v>`));
        const reopened = await importPptx(next, { ...config, firstSlideNumberSaveProfile: "attribute-v2" });
        try { assert.equal(reopened.document.firstSlideNumber, number); assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], round + 11); assert.ok(reopened.report.chartValues.every(b => b.status === "bound")); } finally { reopened.dispose(); }
        previous = number; bytes = next;
      } finally { s.dispose(); }
    }
  }
});

test("number plus an unbound chart or unsupported change never produces a partial plan", async () => {
  const s = await importPptx(encodeChart(fixture({ book: workbook({ formula: true }) })), { ...config, firstSlideNumberSaveProfile: "attribute-v2" });
  try { s.document.setFirstSlideNumber(17); change(s, c => { c.series[0].values[1] = 9; }); const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "blocked"); assert.equal(p.plan, undefined); assert.ok(p.report.diagnostics.some(d => d.code === "save-edit-unsupported")); }
  finally { s.dispose(); }
});


test("combined number and chart saves enforce shared change limits and stale plans", async () => {
  const bytes = encodeChart(fixture()), opts = { ...config, firstSlideNumberSaveProfile: "attribute-v2" };
  const limited = await importPptx(bytes, { ...opts, sessionLimits: { maxSaveChanges: 1 } });
  try { limited.document.setFirstSlideNumber(17); change(limited, c => { c.series[0].values[1] = 9; }); await assert.rejects(limited.prepareSave({ mode: "edit" }), { code: "save-limit-exceeded" }); } finally { limited.dispose(); }
  const s = await importPptx(bytes, opts);
  try { s.document.setFirstSlideNumber(17); change(s, c => { c.series[0].values[1] = 9; }); const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready"); change(s, c => { c.series[0].values[1] = 10; }); await assert.rejects(s.save(p.plan), { code: "save-snapshot-stale" }); } finally { s.dispose(); }
});


test("native field projection stays immutable when chart, adjacent text and starting number are saved together", async () => {
  const field = '<a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="slidenum"><a:rPr b="1"/><a:t>42</a:t></a:fld>';
  for (const strict of [false, true]) {
    const text = body(`<a:p><a:r><a:t>prefix</a:t></a:r>${field}<a:r><a:t>suffix</a:t></a:r></a:p>`);
    let bytes = encodeChart(fixture({ strict, content: frame() + shape({ id: "20", txBox: true, text }) })), previous;
    for (const [round, start] of [17, 0, undefined].entries()) {
      const options = { ...config, editSaveProfile: "text-runs-v3", firstSlideNumberSaveProfile: "attribute-v2", fieldTextProfile: "slide-number-v1" };
      const s = await importPptx(bytes, options);
      try {
        const slide = s.document.slides[0], original = slide.elements[1], changed = structuredClone(original), fieldRun = changed.content[0].runs[1];
        assert.equal(fieldRun.text, "42"); assert.equal(fieldRun.field.kind, "slide-number"); assert.equal(s.document.firstSlideNumber, previous);
        s.document.setFirstSlideNumber(start); change(s, c => { c.series[0].values[1] = 10 + round; });
        changed.content[0].runs[0].text += "!"; slide.replaceElement(original.id, changed);
        const bad = structuredClone(changed); bad.content[0].runs[1].text = "43"; slide.replaceElement(original.id, bad);
        assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked");
        bad.content[0].runs[1].text = "42"; bad.content[0].runs[1].field.id = "11111111-2222-3333-4444-555555555555"; slide.replaceElement(original.id, bad);
        assert.equal((await s.prepareSave({ mode: "edit" })).status, "blocked"); slide.replaceElement(original.id, changed);
        const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready", JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 3);
        const next = (await s.save(p.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(next);
        for (const [name, payload] of before) if (!["deck/main.data", "deck/z.xml", "deck/chart.xml", "deck/book.xlsx"].includes(name)) assert.deepEqual(after.get(name), payload, name);
        const source = before.get("deck/z.xml").toString(), target = after.get("deck/z.xml").toString();
        assert.equal(target, source.replace(`prefix${"!".repeat(round)}</a:t>`, `prefix${"!".repeat(round + 1)}</a:t>`)); assert.ok(target.includes(field));
        assert.match(after.get("deck/chart.xml").toString(), new RegExp(`<c:v>${10 + round}</c:v>`));
        assert.match(readGeneratedZip(after.get("deck/book.xlsx")).get("xl/sheets/one.xml").toString(), new RegExp(`<c r="B3"[^>]*><v>${10 + round}</v>`));
        const reopened = await importPptx(next, options); try { assert.equal(reopened.document.firstSlideNumber, start); assert.equal(reopened.document.slides[0].elements[1].content[0].runs[1].text, "42"); assert.ok(reopened.report.chartValues.every(b => b.status === "bound")); } finally { reopened.dispose(); }
        previous = start; bytes = next;
      } finally { s.dispose(); }
    }
  }
});


test("chart save with field projections also works without starting-number authority", async () => {
  for (const [profile, nativeType, kind] of [["slide-number-v1", "slidenum", "slide-number"], ["date-and-number-v1", "datetime3", "dynamic-date"], ["date-time-and-number-v1", "datetime13", "dynamic-date"]]) {
    const field = `<a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="${nativeType}"><a:t>cached</a:t></a:fld>`;
    const bytes = encodeChart(fixture({ content: frame() + shape({ id: "20", txBox: true, text: body(`<a:p>${field}</a:p>`) }) }));
    const s = await importPptx(bytes, { ...config, fieldTextProfile: profile });
    try { assert.equal(s.document.slides[0].elements[1].content[0].runs[0].field.kind, kind); change(s, c => { c.series[0].values[1] = 7; }); const saved = await save(s); assert.equal(readGeneratedZip(saved.bytes).get("deck/z.xml").toString(), readGeneratedZip(bytes).get("deck/z.xml").toString()); }
    finally { s.dispose(); }
  }
});

test("field, chart and number transaction respects limits, abort, stale plans and exact no-op", async () => {
  const field = '<a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="slidenum"><a:t>42</a:t></a:fld>';
  const bytes = encodeChart(fixture({ content: frame() + shape({ id: "20", txBox: true, text: body(`<a:p><a:r><a:t>before</a:t></a:r>${field}</a:p>`) }) }));
  const options = { ...config, editSaveProfile: "text-runs-v3", firstSlideNumberSaveProfile: "attribute-v2", fieldTextProfile: "slide-number-v1" };
  const mutate = s => { const slide = s.document.slides[0], text = structuredClone(slide.elements[1]); text.content[0].runs[0].text = "after"; slide.replaceElement(text.id, text); change(s, c => { c.series[0].values[1] = 7; }); s.document.setFirstSlideNumber(17); };
  const limited = await importPptx(bytes, { ...options, sessionLimits: { maxSaveChanges: 2 } });
  try { mutate(limited); await assert.rejects(limited.prepareSave({ mode: "edit" }), { code: "save-limit-exceeded" }); } finally { limited.dispose(); }
  const session = await importPptx(bytes, options);
  try {
    const noop = await session.prepareSave({ mode: "no-op" }); assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(bytes)); noop.plan.dispose();
    mutate(session); const controller = new AbortController(); controller.abort(); await assert.rejects(session.prepareSave({ mode: "edit", signal: controller.signal }), { code: "operation-aborted" });
    const plan = await session.prepareSave({ mode: "edit" }); assert.equal(plan.status, "ready", JSON.stringify(plan.report));
    const slide = session.document.slides[0], changed = structuredClone(slide.elements[1]); changed.content[0].runs[1].text = "tampered"; slide.replaceElement(changed.id, changed);
    await assert.rejects(session.save(plan.plan), { code: "save-snapshot-stale" }); const blocked = await session.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.equal(blocked.plan, undefined);
    changed.content[0].runs[1].text = "42"; slide.replaceElement(changed.id, changed); const retry = await session.prepareSave({ mode: "edit" }); assert.equal(retry.status, "ready"); assert.deepEqual((await session.save(retry.plan)).bytes, (await session.save(plan.plan)).bytes);
  } finally { session.dispose(); }
});

test("mixed-v5 atomically saves descriptions, text, chart values and numbering across three reopens", async () => {
  const opts = { ...config, editSaveProfile: "text-runs-v3", descriptionEditSaveProfile: "mixed-v5", firstSlideNumberSaveProfile: "attribute-v2" };
  for (const strict of [false, true]) {
    let bytes = encodeChart(fixture({ strict, content: frame() + shape({ id: "20", txBox: true, text: body() }) }));
    for (const [round, number] of [17, 0, undefined].entries()) {
      const before = readGeneratedZip(bytes), s = await importPptx(bytes, opts);
      try {
        s.document.setFirstSlideNumber(number);
        change(s, c => { c.series[0].values[1] = round + 11; c.accessibility = { ...c.accessibility, description: `Chart ${round}` }; });
        const slide = s.document.slides[0], old = slide.elements[1], edited = structuredClone(old);
        edited.content[0].runs[0].text += "!"; edited.accessibility = { ...edited.accessibility, description: `Text ${round}` }; slide.replaceElement(old.id, edited);
        const p = await s.prepareSave({ mode: "edit" }); assert.equal(p.status, "ready", JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 5);
        const next = (await s.save(p.plan)).bytes, after = readGeneratedZip(next);
        assert.deepEqual([...after.keys()], [...before.keys()]);
        for (const [name, data] of before) if (!["deck/main.data", "deck/z.xml", "deck/chart.xml", "deck/book.xlsx"].includes(name)) assert.deepEqual(after.get(name), data, name);
        const beforeBook = readGeneratedZip(before.get("deck/book.xlsx")), afterBook = readGeneratedZip(after.get("deck/book.xlsx"));
        for (const [name, data] of beforeBook) if (name !== "xl/sheets/one.xml") assert.deepEqual(afterBook.get(name), data, name);
        assert.match(afterBook.get("xl/sheets/one.xml").toString(), new RegExp(`<c r="B3"[^>]*><v>${round + 11}</v>`));
        const reopened = await importPptx(next, opts);
        try {
          assert.equal(reopened.document.firstSlideNumber, number);
          assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], round + 11);
          assert.equal(reopened.document.slides[0].elements[0].accessibility.description, `Chart ${round}`);
          assert.equal(reopened.document.slides[0].elements[1].accessibility.description, `Text ${round}`);
          assert.equal(reopened.document.slides[0].elements[1].content[0].runs[0].text, edited.content[0].runs[0].text);
        } finally { reopened.dispose(); }
        bytes = next;
      } finally { s.dispose(); }
    }
  }
});

test("mixed-v5 rejects unbound chart edits and enforces a shared change limit", async () => {
  const opts = { ...config, editSaveProfile: "text-runs-v3", descriptionEditSaveProfile: "mixed-v5", firstSlideNumberSaveProfile: "attribute-v2" };
  const bad = await importPptx(encodeChart(fixture({ book: workbook({ formula: true }) })), opts);
  try {
    bad.document.setFirstSlideNumber(17); change(bad, c => { c.series[0].values[1] = 9; c.accessibility = { ...c.accessibility, description: "Changed" }; });
    const p = await bad.prepareSave({ mode: "edit" }); assert.equal(p.status, "blocked"); assert.equal(p.plan, undefined);
  } finally { bad.dispose(); }
  const limited = await importPptx(encodeChart(fixture()), { ...opts, sessionLimits: { maxSaveChanges: 1 } });
  try {
    limited.document.setFirstSlideNumber(17); change(limited, c => { c.accessibility = { ...c.accessibility, description: "Changed" }; });
    await assert.rejects(limited.prepareSave({ mode: "edit" }), { code: "save-limit-exceeded" });
  } finally { limited.dispose(); }
  await assert.rejects(importPptx(encodeChart(fixture()), { ...opts, descriptionEditSaveProfile: "mixed-v4" }), TypeError);
  await assert.rejects(importPptx(encodeChart(fixture()), { ...opts, editSaveProfile: undefined }), TypeError);
});

test("mixed-v5 blocks unsupported geometry and stale snapshots without partial saves", async () => {
  const opts = { ...config, editSaveProfile: "text-runs-v3", descriptionEditSaveProfile: "mixed-v5", firstSlideNumberSaveProfile: "attribute-v2" };
  const s = await importPptx(encodeChart(fixture({ content: frame() + shape({ id: "20", txBox: true, text: body() }) })), opts);
  try {
    s.document.setFirstSlideNumber(17); change(s, c => { c.series[0].values[1] = 9; c.accessibility = { ...c.accessibility, description: "Changed" }; });
    const slide = s.document.slides[0], original = slide.elements[1], edited = structuredClone(original);
    edited.content[0].runs[0].text += "!"; edited.box.x++; slide.replaceElement(original.id, edited);
    const blocked = await s.prepareSave({ mode: "edit" }); assert.equal(blocked.status, "blocked"); assert.equal(blocked.plan, undefined);
    edited.box.x--; slide.replaceElement(original.id, edited);
    const ready = await s.prepareSave({ mode: "edit" }); assert.equal(ready.status, "ready", JSON.stringify(ready.report));
    change(s, c => { c.series[0].values[1] = 10; });
    await assert.rejects(s.save(ready.plan), { code: "save-snapshot-stale" });
  } finally { s.dispose(); }
});

test('mixed drawing names and descriptions share source nodes with text, chart and numbering for three rounds', async () => {
  const opts = { ...config, editSaveProfile: 'text-runs-v3', descriptionEditSaveProfile: 'mixed-v5', drawingNameSaveProfile: 'existing-v1', decorativeSaveProfile: 'existing-v1', firstSlideNumberSaveProfile: 'attribute-v2' };
  for (const strict of [false, true]) {
    const ext = '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="0"/></a:ext></a:extLst>';
    const textShape = shape({ id: '20', txBox: true, text: body() }).replace('<p:cNvPr id="20" name="Object 20" descr="Description" />', `<p:cNvPr id="20" name="Object 20" descr="Description">${ext}</p:cNvPr>`);
    let bytes = encodeChart(fixture({ strict, content: frame() + textShape }));
    for (const [round, number] of [17, 0, undefined].entries()) {
      const before = readGeneratedZip(bytes), s = await importPptx(bytes, opts);
      try {
        const slide = s.document.slides[0], oldNames = slide.elements.map(e => e.name), oldDescriptions = slide.elements.map(e => e.accessibility?.description);
        s.document.setFirstSlideNumber(number);
        for (const [i, old] of [...slide.elements].entries()) {
          const edit = structuredClone(old); edit.name = `Name ${round} ${i}`; edit.accessibility = { ...edit.accessibility, description: `Alt ${round} ${i}` };
          if (edit.type === 'chart') edit.series[0].values[1] = 10 + round; else { edit.content[0].runs[0].text += '!'; edit.accessibility.decorative = round % 2 === 0; } slide.replaceElement(old.id, edit);
        }
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report)); assert.equal(p.plan.report.changes.length, 8);
        const output = (await s.save(p.plan)).bytes, after = readGeneratedZip(output);
        for (const [name, payload] of before) if (!['deck/main.data', 'deck/z.xml', 'deck/chart.xml', 'deck/book.xlsx'].includes(name)) assert.deepEqual(after.get(name), payload, name);
        let xml = before.get('deck/z.xml').toString();
        for (let i = 0; i < 2; i++) { xml = xml.replace(`name="${oldNames[i]}"`, `name="Name ${round} ${i}"`); if (oldDescriptions[i] !== undefined) xml = xml.replace(`descr="${oldDescriptions[i]}"`, `descr="Alt ${round} ${i}"`); else xml = xml.replace(`name="Name ${round} ${i}"`, `name="Name ${round} ${i}" descr="Alt ${round} ${i}"`); }
        xml = xml.replace(`val="${round % 2 === 0 ? '0' : '1'}"`, `val="${round % 2 === 0 ? '1' : '0'}"`);
        const text = s.document.slides[0].elements[1].content[0].runs[0].text; const escaped = text.replaceAll("&", "&amp;"); xml = xml.replace(`>${escaped.slice(0, -1)}</a:t>`, `>${escaped}</a:t>`);
        assert.equal(after.get('deck/z.xml').toString(), xml);
        const reopened = await importPptx(output, opts);
        try { assert.equal(reopened.document.firstSlideNumber, number); assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], 10 + round); assert.equal(reopened.document.slides[0].elements[1].content[0].runs[0].text, text); assert.equal(reopened.document.slides[0].elements[1].accessibility.decorative, round % 2 === 0); for (const [i, e] of reopened.document.slides[0].elements.entries()) { assert.equal(e.name, `Name ${round} ${i}`); assert.equal(e.accessibility.description, `Alt ${round} ${i}`); } } finally { reopened.dispose(); }
        bytes = output;
      } finally { s.dispose(); }
    }
  }
});


test("custom scalar and chart workbook edits share a single guarded transaction", async () => {
  for (const strict of [false, true]) {
    const base = readGeneratedZip(encodeChart(fixture({ strict })));
    const rootBase = strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    const ns = strict ? "http://purl.oclc.org/ooxml/officeDocument/customProperties" : "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
    const vt = strict ? "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes" : "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
    base.set('_rels/.rels', Buffer.from(base.get('_rels/.rels').toString().replace('</Relationships>', `<Relationship Id="custom" Type="${rootBase}" Target="docProps/custom.xml"/></Relationships>`)));
    base.set('[Content_Types].xml', Buffer.from(base.get('[Content_Types].xml').toString().replace('</Types>', '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')));
    base.set('docProps/custom.xml', Buffer.from(`<Properties xmlns="${ns}" xmlns:v="${vt}"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="K"><v:lpwstr>old</v:lpwstr></property></Properties>`));
    let bytes = zip([...base].map(([name, body]) => ({ name, body }))).bytes;
    const opts = { ...config, customPropertySaveProfile: 'scalar-v1' };
    for (let round = 0; round < 3; round++) {
      const s = await importPptx(bytes, opts);
      try {
        s.document.setMetadata({ ...s.document.metadata, customProperties: [{ name: 'K', type: 'string', value: `new ${round}` }] });
        change(s, c => { c.series[0].values[1] = round + 10; });
        change(s, c => { c.box.x++; });
        const blocked = await s.prepareSave({ mode: 'edit' }); assert.equal(blocked.status, 'blocked'); assert.equal(blocked.plan, undefined);
        change(s, c => { c.box.x--; });
        const p = await s.prepareSave({ mode: 'edit' }); assert.equal(p.status, 'ready', JSON.stringify(p.report));
        const output = (await s.save(p.plan)).bytes, before = readGeneratedZip(bytes), after = readGeneratedZip(output);
        for (const [name, data] of before) if (!['docProps/custom.xml', 'deck/chart.xml', 'deck/book.xlsx'].includes(name)) assert.deepEqual(after.get(name), data, name);
        const reopened = await importPptx(output, opts);
        try { assert.equal(reopened.document.metadata.customProperties[0].value, `new ${round}`); assert.equal(reopened.document.slides[0].elements[0].series[0].values[1], round + 10); } finally { reopened.dispose(); }
        bytes = output;
      } finally { s.dispose(); }
    }
  }
});

test('chart workbook edits and metadata deletion/reinsertion commit atomically over three rounds', async () => {
  for (const strict of [false, true]) {
    const base=readGeneratedZip(encodeChart(fixture({strict})));
    base.set('_rels/.rels',Buffer.from(base.get('_rels/.rels').toString().replace('</Relationships>','<Relationship Id="core" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>')));
    base.set('[Content_Types].xml',Buffer.from(base.get('[Content_Types].xml').toString().replace('</Types>','<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')));
    base.set('docProps/core.xml',Buffer.from('<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:subject>Subject</dc:subject><!--sentinel--><cp:keywords>old</cp:keywords><cp:revision>7</cp:revision></cp:coreProperties>'));
    let bytes=zip([...base].map(([name,body])=>({name,body}))).bytes;const opts={...config,metadataSaveProfile:'metadata-keywords-remove-v15'};
    for(let round=0;round<3;round++){
      const s=await importPptx(bytes,opts);try{
        s.document.setMetadata({...s.document.metadata,subject:round===1?'new subject':undefined,keywords:round===1?['new keyword']:[]});change(s,c=>{c.series[0].values[1]=round+11;});
        change(s,c=>{c.box.x++;});const blocked=await s.prepareSave({mode:'edit'});assert.equal(blocked.status,'blocked');assert.equal(blocked.plan,undefined);change(s,c=>{c.box.x--;});
        const p=await s.prepareSave({mode:'edit'});assert.equal(p.status,'ready',JSON.stringify(p.report));const output=(await s.save(p.plan)).bytes,before=readGeneratedZip(bytes),after=readGeneratedZip(output);
        for(const [name,payload] of before)if(!['docProps/core.xml','deck/chart.xml','deck/book.xlsx'].includes(name))assert.deepEqual(after.get(name),payload,name);
        assert.match(after.get('docProps/core.xml').toString(),/<!--sentinel-->/);assert.equal(after.get('deck/chart.xml').toString(),before.get('deck/chart.xml').toString().replace(/(<c:val>[\s\S]*?<c:pt idx="1"><c:v>)[^<]*(<\/c:v>)/,`$1${round+11}$2`));
        const innerBefore=readGeneratedZip(before.get('deck/book.xlsx')),inner=readGeneratedZip(after.get('deck/book.xlsx'));for(const [name,payload] of innerBefore)if(name!=='xl/sheets/one.xml')assert.deepEqual(inner.get(name),payload,name);assert.equal(inner.get('xl/sheets/one.xml').toString(),innerBefore.get('xl/sheets/one.xml').toString().replace(/(<c r="B3"[^>]*><v>)[^<]*(<\/v>)/,`$1${round+11}$2`));
        const reopened=await importPptx(output,opts);try{assert.equal(reopened.document.slides[0].elements[0].series[0].values[1],round+11);assert.equal(reopened.document.metadata.subject,round===1?'new subject':undefined);assert.deepEqual(reopened.document.metadata.keywords??[],round===1?['new keyword']:[]);assert.equal(reopened.document.metadata.revision,7);}finally{reopened.dispose();}bytes=output;
      }finally{s.dispose();}
    }
  }
});
