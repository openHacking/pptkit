import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance, editTable } from '../dist/index.js';
import { tableOriginsDocument } from '../../../tests/fixtures/table-origins.mjs';
const fixture = () => tableOriginsDocument(createPresentation);
const source = path => ({ kind: 'element', id: 'table-0', path: `slides.0.elements.0.${path}` });
const preset = suffix => ({ kind: 'preset', id: 'cell."preset', path: `textStylePresets["cell.\\\"preset"].${suffix}` });
function origin(r, id, property) {
  const rows = r.origins.filter(o => o.target.kind === 'element' && o.target.id === id && o.property === property);
  assert.equal(rows.length, 1, `${id}/${property}`); return rows[0];
}
const cellOrigin = (r, rIndex, cIndex, suffix) => origin(r, 'table-0', `rows.${rIndex}.cells.${cIndex}.${suffix}`);
const builtin = path => ({ kind: 'default', path });

test('table text direction origins follow regions, explicit horizontal overrides and merged cells', () => {
  const d=createPresentation();d.addSlide({id:'slide'}).addElement({id:'table-0',type:'table',box:{x:0,y:0,width:90,height:40},columns:[30,30,30],rows:[{height:20,cells:[{content:'Merged',rowSpan:2,style:{textDirection:'horizontal'}},{content:'Whole'},{content:'Last'}]},{height:20,cells:[]}],style:{wholeTable:{textDirection:'rotate90'},lastColumn:{textDirection:'rotate270'}}});
  const before=structuredClone(d.slides[0].elements[0]),r=resolvePresentationInheritance(d),table=r.presentation.slides[0].elements[0];
  assert.deepEqual(table.rows.map(row=>row.cells.map(c=>c.style.textDirection)),[['horizontal','rotate90','rotate270'],['rotate90','rotate270']]);
  assert.deepEqual(cellOrigin(r,0,0,'style.textDirection').source,source('rows.0.cells.0.style.textDirection'));
  assert.deepEqual(cellOrigin(r,0,1,'style.textDirection').source,source('style.wholeTable.textDirection'));
  assert.deepEqual(cellOrigin(r,0,2,'style.textDirection').source,source('style.lastColumn.textDirection'));
  assert.deepEqual(cellOrigin(r,1,0,'style.textDirection').source,source('style.wholeTable.textDirection'));
  assert.deepEqual(cellOrigin(r,1,1,'style.textDirection').source,source('style.lastColumn.textDirection'));
  assert.deepEqual(d.slides[0].elements[0],before);
  const plain=createPresentation();plain.addSlide().addElement({id:'plain',type:'table',box:{x:0,y:0,width:10,height:10},columns:[10],rows:[{height:10,cells:[{content:''}]}]});
  const p=resolvePresentationInheritance(plain);assert.equal(p.presentation.slides[0].elements[0].rows[0].cells[0].style.textDirection,undefined);assert.equal(p.origins.some(o=>o.target.id==='plain'&&o.property.endsWith('.textDirection')),false);
  const local=structuredClone(plain.slides[0].elements[0]);local.rows[0].cells[0].style={textDirection:'horizontal'};plain.slides[0].replaceElement('plain',local);
  const q=resolvePresentationInheritance(plain);assert.deepEqual(origin(q,'plain','rows.0.cells.0.style.textDirection').source,{kind:'element',id:'plain',path:'slides.0.elements.0.rows.0.cells.0.style.textDirection'});
});

test('regional cell sources follow real precedence without changing ordinary IR, shared themes or authored grids', () => {
  const d = fixture(), before = JSON.stringify(d.slides[0].elements[0]), r = resolvePresentationInheritance(d);
  assert.deepEqual(r.presentation, normalizePresentation(d));
  assert.equal(JSON.stringify(d.slides[0].elements[0]), before);
  const table = r.presentation.slides[0].elements[0];
  assert.deepEqual(table.rows.map(row => row.cells.length), [3,2,4,4]);
  assert.deepEqual(cellOrigin(r, 3, 3, 'content.0.runs.0.style.fontSize').source, source('style.bottomRight.text.run.fontSize'));
  assert.deepEqual(cellOrigin(r, 0, 1, 'content.0.runs.0.style.fontSize').source, source('style.wholeTable.text.run.fontSize'));
  assert.deepEqual(cellOrigin(r, 2, 1, 'content.0.runs.0.style.fontSize').source, source('style.bandedColumns.first.text.run.fontSize'));
  assert.deepEqual(table.rows[2].cells[1].content[0].runs[0].style.fontFamily, { theme: 'body' });
  assert.equal(r.presentation.layouts.find(l => l.id === 'layout-0').theme.fonts.body, 'Georgia');
  assert.equal(r.presentation.layouts.find(l => l.id === 'layout-1').theme.fonts.body, 'Verdana');
});

test('scalar/partial margins, preset and cell overrides retain independent declaration sources', () => {
  const r = resolvePresentationInheritance(fixture());
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.margin.left').source, preset('frame.margin.left'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.margin.right').source, source('rows.0.cells.0.style.margin.right'));
  for (const side of ['top','bottom']) assert.deepEqual(cellOrigin(r, 0, 0, `style.margin.${side}`).source, source('style.wholeTable.margin'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.verticalAlign').source, source('rows.0.cells.0.style.verticalAlign'));
  const plain = origin(r, 'plain-table', 'rows.0.cells.0.style.margin.left');
  assert.deepEqual(plain.source, builtin('defaults.table.cell.margin.left'));
  assert.equal(r.presentation.slides[2].elements[0].children[0].rows[0].cells[0].style.margin.left, 7);
  assert.deepEqual(origin(r, 'plain-table', 'rows.0.cells.0.style.verticalAlign').source, preset('frame.verticalAlign'));
});

test('uniform strokes and sparse edges retain fieldwise sources including fallback copies and atomic paint defaults', () => {
  const r = resolvePresentationInheritance(fixture()), c = r.presentation.slides[0].elements[0].rows[0].cells[0];
  for (const side of ['left','right','top','bottom']) {
    assert.deepEqual(cellOrigin(r, 0, 0, `style.borders.${side}.width`).source, source('rows.0.cells.0.style.stroke.width'));
    assert.equal(c.style.borders[side].width, 0);
  }
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.borders.left.dash').source, source('rows.0.cells.0.style.borders.left.dash'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.borders.right.dash').source, source('style.borders.innerVertical.dash'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.borders.bottom.dash').source, source('style.firstRow.borders.bottom.dash'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.borders.top.paint.color').source, source('style.borders.outer.paint.color'));
  assert.equal(c.style.borders.top.paint.opacity, 1);
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.borders.top.paint.opacity').source, builtin('defaults.table.cell.borders.top.paint.opacity'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'style.borders.right.paint.opacity').source, source('style.wholeTable.stroke.paint.opacity'));
  assert.deepEqual(origin(r, 'plain-table', 'rows.0.cells.0.style.borders.left.width').source, { kind: 'element', id: 'plain-table', path: 'slides.2.elements.0.children.0.rows.0.cells.0.style.stroke.width' });
  assert.deepEqual(c.style.fill, { type: 'none' });
  assert.equal(r.origins.some(o => o.target.id === 'table-0' && o.property.startsWith('rows.0.cells.0.style.fill.') && o.property !== 'rows.0.cells.0.style.fill.type'), false);
});

test('raw text merges preserve preset, paragraph/run and whole-map clear barriers with reopened slots', () => {
  const r = resolvePresentationInheritance(fixture());
  assert.deepEqual(cellOrigin(r, 0, 0, 'content.0.runs.0.style.bold').source, source('rows.0.cells.0.content.0.runs.0.style.bold'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'content.0.runs.1.style.bold').source, preset('run.bold'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'content.0.runs.1.style.eastAsianFontFamily').source, preset('run.eastAsianFontFamily'));
  assert.equal(cellOrigin(r, 0, 0, 'content.0.runs.1.style.eastAsianFontFamily').state, 'cleared');
  assert.equal(cellOrigin(r, 0, 0, 'content.0.runs.1.style.scriptFontFamilies.Hans').state, 'cleared');
  assert.deepEqual(cellOrigin(r, 0, 0, 'content.0.runs.0.style.scriptFontFamilies.Hans').source, preset('run.scriptFontFamilies.Hans'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'content.0.runs.0.style.scriptFontFamilies.Arab').source, source('rows.0.cells.0.content.0.runs.0.style.scriptFontFamilies.Arab'));
  assert.deepEqual(cellOrigin(r, 1, 0, 'content.0.runs.0.style.scriptFontFamilies.Arab').source, source('style.bandedRows.first.text.run.scriptFontFamilies'));
  assert.equal(cellOrigin(r, 1, 0, 'content.0.runs.0.style.scriptFontFamilies.Arab').state, 'cleared');
  assert.deepEqual(cellOrigin(r, 0, 0, 'content.0.style.bullet.startAt').source, builtin('defaults.text.paragraph.bullet.startAt'));
  assert.deepEqual(cellOrigin(r, 0, 0, 'content.0.style.bullet.type').source, source('rows.0.cells.0.content.0.style.bullet.type'));
  assert.deepEqual(cellOrigin(r, 0, 1, 'content.0.style.bullet.character').source, builtin('defaults.text.paragraph.bullet.character'));
});

test('implicit blanks and generated insertion styles have only genuine regional/preset/default sources', () => {
  const d = fixture(), r = resolvePresentationInheritance(d), t = r.presentation.slides[0].elements[0];
  assert.equal(d.slides[0].elements[0].rows[2].cells.length, 0);
  for (const c of t.rows[2].cells) assert.equal(c.content[0].runs[0].text, '');
  assert.deepEqual(cellOrigin(r, 2, 0, 'content.0.style.align').source, source('style.firstColumn.text.paragraph.align'));
  assert.deepEqual(cellOrigin(r, 2, 3, 'style.verticalAlign').source, source('style.lastColumn.verticalAlign'));
  assert.equal(r.origins.some(o => o.target.id === 'table-0' && /^rows\.2\./.test(o.property) && o.source.kind !== 'default' && /\.rows\.2\.cells\./.test(o.source.path)), false);
  assert.equal(t.rows[0].cells[2].content[0].runs[0].text, '');
  assert.deepEqual(cellOrigin(r, 0, 2, 'content.0.runs.0.style.bold').source, source('style.wholeTable.text.run.bold'));
  assert.equal(t.rows[3].cells[1].content[0].runs[0].text, '');
  assert.deepEqual(cellOrigin(r, 3, 1, 'content.0.runs.0.style.bold').source, source('style.wholeTable.text.run.bold'));
});

function leaves(value, path) {
  const atomic = /\.(?:color|fontFamily|eastAsianFontFamily|complexScriptFontFamily)$/.test(path) || /\.scriptFontFamilies\.[^.]+$/.test(path);
  return value !== null && typeof value === 'object' && !atomic ? Object.entries(value).flatMap(([key, child]) => leaves(child, `${path}.${key}`)) : [path];
}
const read = (value, path) => path.split('.').reduce((v, key) => v?.[key], value);
function checkSourceLeaves(d, r) {
  const elements = new Map(); const walk = values => { for (const e of values) { elements.set(e.id, e); if (e.type === 'group') walk(e.children); } };
  r.presentation.layouts.forEach(l => walk(l.elements)); r.presentation.slides.forEach(s => walk(s.elements));
  let declarations = 0, clears = 0, coverage = 0;
  for (const e of elements.values()) if (e.type === 'table') e.rows.forEach((row, ri) => row.cells.forEach((cell, ci) => {
    const prefix = `rows.${ri}.cells.${ci}`;
    const required = leaves(cell.style, `${prefix}.style`);
    cell.content.forEach((p, pi) => { required.push(...leaves(p.style, `${prefix}.content.${pi}.style`)); p.runs.forEach((run, xi) => { required.push(...leaves(run.style, `${prefix}.content.${pi}.runs.${xi}.style`)); if (run.style.colorOpacity === undefined) required.push(`${prefix}.content.${pi}.runs.${xi}.style.colorOpacity`); }); });
    for (const key of required) { origin(r, e.id, key); coverage++; }
  }));
  for (const o of r.origins) {
    const target = elements.get(o.target.id);
    if (o.target.kind !== 'element' || target?.type !== 'table') continue;
    if (o.source.kind === 'default') { assert.match(o.source.path, /^defaults\.(?:table\.cell|text\.(?:paragraph|run))\./); continue; }
    const declared = o.source.kind === 'preset' ? read(d.textStylePresets[o.source.id], o.source.path.slice(`textStylePresets[${JSON.stringify(o.source.id)}].`.length)) : read(d, o.source.path);
    const effective = read(target, o.property);
    if (o.state === 'cleared') { assert.equal(declared, null, JSON.stringify(o)); assert.equal(effective, undefined, JSON.stringify(o)); clears++; }
    else { assert.deepEqual(effective === undefined && o.property.endsWith('.colorOpacity') ? 1 : effective, declared, JSON.stringify(o)); declarations++; }
  }
  const identities = r.origins.map(o => JSON.stringify([o.target, o.property])); assert.equal(new Set(identities).size, identities.length);
  return { declarations, clears, coverage };
}

test('every normalized formatting leaf is covered once and each authored source independently reads the selected value/null', () => {
  const d = fixture(), r = resolvePresentationInheritance(d), counts = checkSourceLeaves(d, r);
  assert.ok(counts.declarations > 800); assert.ok(counts.clears > 100); assert.ok(counts.coverage > 1200);
});

test('raw border fallback origin copies survive later partial uniform strokes and atomic replacements', () => {
  const d = createPresentation();
  d.addSlide({ elements: [{ id: 'edge', type: 'table', box: { x: 0, y: 0, width: 100, height: 40 }, columns: [100], rows: [{ cells: [{ content: '' }] }], style: { wholeTable: { stroke: { width: 2, paint: { type: 'solid', color: '111111', opacity: 0.4 } } }, firstRow: { borders: { left: { dash: 'dash' } } }, lastRow: { stroke: { width: 0 } }, bottomRight: { stroke: { paint: { type: 'none' } }, borders: { top: { paint: { type: 'solid', color: '222222' } } } } } }] });
  const r = resolvePresentationInheritance(d); checkSourceLeaves(d, r);
  assert.deepEqual(origin(r, 'edge', 'rows.0.cells.0.style.borders.left.width').source, { kind: 'element', id: 'edge', path: 'slides.0.elements.0.style.lastRow.stroke.width' });
  assert.equal(r.origins.some(o => o.target.id === 'edge' && o.property === 'rows.0.cells.0.style.borders.left.paint.color'), false);
  assert.deepEqual(origin(r, 'edge', 'rows.0.cells.0.style.borders.top.paint.opacity').source, builtin('defaults.table.cell.borders.top.paint.opacity'));
});

test('reset and table structural edits recompute current sources without mutating previous snapshots', () => {
  const d = fixture(), slide = d.slides[0], before = resolvePresentationInheritance(d), saved = JSON.stringify(before.origins), t = slide.elements[0];
  const edited = editTable(t, [{ type: 'insert-row', index: 2, height: 45 }], { frame: 'fit-grid' }); slide.replaceElement(t.id, edited.table);
  const after = resolvePresentationInheritance(d); checkSourceLeaves(d, after);
  assert.equal(after.presentation.slides[0].elements[0].rows.length, 5);
  assert.equal(JSON.stringify(before.origins), saved);
  assert.throws(() => { cellOrigin(before, 0, 0, 'style.margin.left').source.path = 'bad'; }, TypeError);
  const current = slide.elements[0], { style: _rules, ...legacy } = current; slide.replaceElement(t.id, legacy);
  const reset = resolvePresentationInheritance(d); checkSourceLeaves(d, reset);
  assert.equal(reset.origins.some(o => o.target.id === t.id && /\.style\.(?:banded|wholeTable|firstRow|lastRow|bottomRight)/.test(o.source.path)), false);
  assert.deepEqual(reset.presentation, normalizePresentation(d));
});

test('equal paints retain distinct rule owners and default-valued presets retain explicit sources', () => {
  const d = createPresentation({ textStylePresets: { equal: { frame: { margin: 0 }, run: { bold: false } } } });
  d.addSlide({ elements: [{ id: 'equal', type: 'table', box: { x: 0, y: 0, width: 100, height: 40 }, columns: [50, 50], rows: [{ cells: [{ content: '' }, { content: '', textStylePreset: 'equal' }] }], style: { wholeTable: { fill: { type: 'solid', color: 'FFFFFF' } }, firstColumn: { fill: { type: 'solid', color: 'FFFFFF' } }, lastColumn: { fill: { type: 'solid', color: 'FFFFFF' } } } }] });
  const r = resolvePresentationInheritance(d); checkSourceLeaves(d, r);
  assert.equal(origin(r, 'equal', 'rows.0.cells.0.style.fill.color').source.path, 'slides.0.elements.0.style.firstColumn.fill.color');
  assert.equal(origin(r, 'equal', 'rows.0.cells.1.style.fill.color').source.path, 'slides.0.elements.0.style.lastColumn.fill.color');
  assert.deepEqual(origin(r, 'equal', 'rows.0.cells.1.content.0.runs.0.style.bold').source, { kind: 'preset', id: 'equal', path: 'textStylePresets["equal"].run.bold' });
});


test('96 independently partitioned grids attribute equal-valued fills to the actual geometric winning region', () => {
  let seed = 0x21a9c83;
  const random = n => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % n; };
  const names = ['wholeTable', 'firstColumn', 'lastColumn', 'firstRow', 'lastRow', 'topLeft', 'topRight', 'bottomLeft', 'bottomRight'];
  for (let round = 0; round < 96; round++) {
    const height = 1+random(6), width = 1+random(6), rects = [{ r: 0, c: 0, h: height, w: width }];
    for (let cut = 0; cut < 10; cut++) {
      const i = random(rects.length), a = rects[i];
      if (a.h > 1 && (a.w === 1 || random(2) === 0)) { const h = 1+random(a.h-1); rects.splice(i,1,{...a,h},{...a,r:a.r+h,h:a.h-h}); }
      else if (a.w > 1) { const w = 1+random(a.w-1); rects.splice(i,1,{...a,w},{...a,c:a.c+w,w:a.w-w}); }
    }
    const style = {}, region = () => ({ fill: { type: 'solid', color: '112233' }, text: { run: { fontSize: 17 } } });
    for (const name of names) if (random(3) !== 0) style[name] = region();
    style.bandedRows = { size: 1+random(2), first: region(), second: region() };
    style.bandedColumns = { size: 1+random(2), first: region(), second: region() };
    const d = createPresentation(), rows = Array.from({ length: height }, () => ({ cells: [] }));
    rects.sort((a,b) => a.r-b.r || a.c-b.c);
    for (const a of rects) rows[a.r].cells.push({ content: `${a.r}:${a.c}`, rowSpan: a.h, colSpan: a.w });
    d.addSlide({ elements: [{ id: 'partition', type: 'table', box: { x: 0, y: 0, width: width*30, height: height*20 }, columns: Array(width).fill(30), rows, style }] });
    const r = resolvePresentationInheritance(d), table = r.presentation.slides[0].elements[0];
    assert.deepEqual(r.presentation, normalizePresentation(d)); checkSourceLeaves(d, r);
    for (const a of rects) {
      let winner = style.wholeTable ? 'wholeTable' : undefined;
      for (const [position, count, first, last, name] of [[a.c,width,'firstColumn','lastColumn','bandedColumns'],[a.r,height,'firstRow','lastRow','bandedRows']]) {
        if ((style[first] && position===0) || (style[last] && position===count-1)) continue;
        const parity = Math.floor((position-(style[first]?1:0))/style[name].size)%2;
        winner = `${name}.${parity===0?'first':'second'}`;
      }
      const touching = { firstColumn: a.c===0, lastColumn: a.c+a.w===width, firstRow: a.r===0, lastRow: a.r+a.h===height, topLeft: a.r===0&&a.c===0, topRight: a.r===0&&a.c+a.w===width, bottomLeft: a.r+a.h===height&&a.c===0, bottomRight: a.r+a.h===height&&a.c+a.w===width };
      for (const name of names.slice(1)) if (touching[name] && style[name]) winner = name;
      const ci = table.rows[a.r].cells.findIndex(c => c.content[0].runs[0].text===`${a.r}:${a.c}`);
      for (const [property, leaf, defaultPath] of [['style.fill.type','fill.type','defaults.table.cell.fill.type'],['content.0.runs.0.style.fontSize','text.run.fontSize','defaults.text.run.fontSize']]) {
        const o = origin(r, 'partition', `rows.${a.r}.cells.${ci}.${property}`);
        assert.deepEqual(o.source, winner ? { kind: 'element', id: 'partition', path: `slides.0.elements.0.style.${winner}.${leaf}` } : builtin(defaultPath), `${round}/${a.r}/${a.c}`);
      }
    }
  }
});


test('aliased authored cell objects retain distinct occurrence paths rather than last-reference ownership', () => {
  const cell = { content: [{ runs: [{ text: 'Same', style: { bold: false } }] }], style: { margin: 0 } };
  const d = createPresentation(); d.addSlide({ elements: [{ id: 'aliased', type: 'table', box: { x: 0, y: 0, width: 100, height: 40 }, columns: [50,50], rows: [{ cells: [cell,cell] }], style: {} }] });
  assert.equal(d.slides[0].elements[0].rows[0].cells[0], d.slides[0].elements[0].rows[0].cells[1]);
  const r = resolvePresentationInheritance(d); checkSourceLeaves(d, r);
  for (let ci=0;ci<2;ci++) {
    assert.deepEqual(origin(r, 'aliased', `rows.0.cells.${ci}.content.0.runs.0.style.bold`).source, { kind: 'element', id: 'aliased', path: `slides.0.elements.0.rows.0.cells.${ci}.content.0.runs.0.style.bold` });
    assert.deepEqual(origin(r, 'aliased', `rows.0.cells.${ci}.style.margin.left`).source, { kind: 'element', id: 'aliased', path: `slides.0.elements.0.rows.0.cells.${ci}.style.margin` });
  }
});
