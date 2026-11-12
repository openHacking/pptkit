import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance } from '../packages/core/dist/index.js';
import { resolveLayout, resolveNormalizedLayout } from '../packages/layout/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderLayoutToSvg } from '../packages/svg-renderer/dist/index.js';
import { inspectPptx, importPptx } from '../packages/pptx-parser/dist/index.js';
import { multiMasterDocument } from '../tests/fixtures/multi-master.mjs';
import { editMasterDefinitions } from '../tests/fixtures/master-definition-edit.mjs';
import { chartNotesOriginsDocument } from '../tests/fixtures/chart-notes-origins.mjs';
import { textOriginsDocument } from '../tests/fixtures/text-origins.mjs';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

async function run(api) {
  const document = multiMasterDocument(api.createPresentation), normalized = api.normalizePresentation(document), layout = api.resolveLayout(document);
  const svg = await api.renderLayoutToSvg(layout), generated = await api.generatePptx(document);
  return { normalized, layout, svg, bytes: [...generated.bytes], warnings: generated.warnings };
}

test('master themes stay isolated in Node/ESM/global authoring, layout, SVG, native XML and source no-op', async t => {
  const expected = await run({ createPresentation, normalizePresentation, resolveLayout, generatePptx, renderLayoutToSvg });
  assert.deepEqual(expected.warnings, []);
  for (const [i, color, font] of [[0, 'CC2200', 'Georgia'], [1, '0044CC', 'Times New Roman'], [2, 'CC2200', 'Georgia'], [3, '112233', 'Arial']]) {
    assert.ok(expected.svg.slides[i].svg.includes('#' + color)); assert.ok(expected.svg.slides[i].svg.includes(font));
  }
  assert.ok(expected.svg.slides[0].svg.includes('#F8E8E0')); assert.ok(expected.svg.slides[1].svg.includes('#E0E8F8'));
  const bytes = Uint8Array.from(expected.bytes), inspection = await inspectPptx(bytes); assert.equal(inspection.slideMasters.length, 4);
  const session = await importPptx(bytes, { projection: 'structure' });
  try { const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready'); assert.deepEqual([...((await session.save(plan.plan)).bytes)], expected.bytes); } finally { session.dispose(); }
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', r => { requests.push(r.request().url()); return r.abort(); });
  const imports = {}, globals = [];
  for (const name of ['core', 'layout', 'pptx-exporter', 'svg-renderer', 'pptx-parser']) {
    imports['@pptkit/' + name] = 'data:text/javascript;base64,' + Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`, 'utf8')).toString('base64');
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`, 'utf8'));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const format of ['esm', 'global']) {
    if (format === 'global') for (const content of globals) await page.addScriptTag({ content });
    const actual = await page.evaluate(async ({ format, fixture }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'layout', 'pptx-exporter', 'svg-renderer'].map(n => import('@pptkit/' + n))));
      const document = Function('createPresentation', `return (${fixture})(createPresentation)`)(api.createPresentation), normalized = api.normalizePresentation(document), layout = api.resolveLayout(document);
      return { normalized, layout, svg: await api.renderLayoutToSvg(layout), bytes: [...(await api.generatePptx(document)).bytes], warnings: [] };
    }, { format, fixture: multiMasterDocument.toString() });
    assert.deepEqual(actual, expected);
  }
  const truth = await page.evaluate(parts => {
    const P = 'http://schemas.openxmlformats.org/presentationml/2006/main', A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const doc = part => new DOMParser().parseFromString(parts[part], 'application/xml');
    const relationship = (part, type) => [...doc(part).getElementsByTagName('Relationship')].filter(n => n.getAttribute('Type').endsWith('/' + type));
    const resolve = (part, target) => new URL(target, 'https://source.invalid/' + part).pathname.slice(1);
    const rows = [];
    for (let i = 1; i <= 4; i++) {
      const layout = resolve('ppt/slides/slide' + i + '.xml', relationship('ppt/slides/_rels/slide' + i + '.xml.rels', 'slideLayout')[0].getAttribute('Target'));
      const master = resolve(layout, relationship(layout.replace('/slideLayouts/', '/slideLayouts/_rels/') + '.rels', 'slideMaster')[0].getAttribute('Target'));
      const theme = resolve(master, relationship(master.replace('/slideMasters/', '/slideMasters/_rels/') + '.rels', 'theme')[0].getAttribute('Target'));
      const color = doc(theme).getElementsByTagNameNS(A, 'accent1')[0].firstElementChild.getAttribute('val');
      rows.push({ layout, master, theme, color });
    }
    const ids = [1, 2, 3, 4].flatMap(i => [...doc('ppt/slideMasters/slideMaster' + i + '.xml').getElementsByTagNameNS(P, 'sldLayoutId')].map(n => n.getAttribute('id')));
    const masterIds = [...doc('ppt/presentation.xml').getElementsByTagNameNS(P, 'sldMasterId')].map(n => n.getAttribute('id'));
    if (ids.length !== 5 || new Set([...ids, ...masterIds]).size !== ids.length + masterIds.length) throw Error('Native master/layout identities collide');
    return rows;
  }, Object.fromEntries([...readGeneratedZip(bytes)].filter(([n]) => n.endsWith('.xml') || n.endsWith('.rels')).map(([n, b]) => [n, b.toString()])));
  assert.deepEqual(truth.map(r => r.color), ['CC2200', '0044CC', 'CC2200', '112233']);
  assert.deepEqual(truth.map(r => r.master), ['ppt/slideMasters/slideMaster2.xml', 'ppt/slideMasters/slideMaster3.xml', 'ppt/slideMasters/slideMaster2.xml', 'ppt/slideMasters/slideMaster1.xml']);
  assert.deepEqual(requests, []);
});

async function textOriginSnapshot(api, fixture) {
  const d = fixture(api.createPresentation), resolution = api.resolvePresentationInheritance(d);
  const normalized = api.normalizePresentation(d);
  if (JSON.stringify(normalized) !== JSON.stringify(resolution.presentation)) throw new Error('Style trace changed normalized values');
  if (!Object.isFrozen(resolution.origins) || !resolution.origins.every(o => Object.isFrozen(o.source))) throw new Error('Mutable style origin');
  const layout = api.resolveNormalizedLayout(normalized), output = await api.generatePptx(d);
  return { normalized, origins: resolution.origins, layout, svg: await api.renderLayoutToSvg(layout), bytes: [...output.bytes], warnings: output.warnings };
}

test('text format declarations and clear barriers match through Node/browser IR, report, Layout, SVG and bytes', async t => {
  const expected = await textOriginSnapshot({ createPresentation, normalizePresentation, resolvePresentationInheritance, resolveNormalizedLayout, generatePptx, renderLayoutToSvg }, textOriginsDocument);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', r => { requests.push(r.request().url()); return r.abort(); });
  const imports = {}, globals = [];
  for (const name of ['core', 'layout', 'pptx-exporter', 'svg-renderer']) {
    imports['@pptkit/' + name] = 'data:text/javascript;base64,' + Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`, 'utf8')).toString('base64');
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`, 'utf8'));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const format of ['esm', 'global']) {
    if (format === 'global') for (const content of globals) await page.addScriptTag({ content });
    const actual = await page.evaluate(async ({ format, fixture, snapshot }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'layout', 'pptx-exporter', 'svg-renderer'].map(n => import('@pptkit/' + n))));
      return Function('api', 'fixture', `return (${snapshot})(api, (${fixture}))`)(api);
    }, { format, fixture: textOriginsDocument.toString(), snapshot: textOriginSnapshot.toString() });
    assert.deepEqual(actual, expected);
  }
  const parts = readGeneratedZip(Uint8Array.from(expected.bytes)), xml = parts.get('ppt/slides/slide1.xml').toString();
  const layoutXml = [...parts].find(([name, bytes]) => name.startsWith('ppt/slideLayouts/') && name.endsWith('.xml') && bytes.toString().includes('<p:cSld name="A">'))[1].toString();
  const native = await page.evaluate(({ xml, layoutXml }) => {
    const P = 'http://schemas.openxmlformats.org/presentationml/2006/main', A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const d = new DOMParser().parseFromString(xml, 'application/xml');
    const shape = [...d.getElementsByTagNameNS(P, 'sp')].find(s => s.getElementsByTagNameNS(P, 'cNvPr')[0].getAttribute('name') === 'body');
    const frame = shape.getElementsByTagNameNS(A, 'bodyPr')[0], run = shape.getElementsByTagNameNS(A, 'r')[0], rPr = run.getElementsByTagNameNS(A, 'rPr')[0];
    const layout = new DOMParser().parseFromString(layoutXml, 'application/xml');
    const shared = [...layout.getElementsByTagNameNS(P, 'sp')].find(s => s.getElementsByTagNameNS(P, 'cNvPr')[0].getAttribute('name') === 'shared-text');
    return { text: run.getElementsByTagNameNS(A, 't')[0].textContent, bold: rPr.getAttribute('b'), italic: rPr.getAttribute('i'), underline: rPr.getAttribute('u'), strike: rPr.getAttribute('strike'), alpha: rPr.getElementsByTagNameNS(A, 'alpha')[0].getAttribute('val'), margins: ['lIns', 'rIns', 'tIns', 'bIns'].map(k => frame.getAttribute(k)), wrap: frame.getAttribute('wrap'), fontScale: frame.getElementsByTagNameNS(A, 'normAutofit')[0].getAttribute('fontScale'), sharedText: shared.getElementsByTagNameNS(A, 't')[0].textContent, sharedPlaceholderCount: shared.getElementsByTagNameNS(P, 'ph').length, layoutPlaceholderCount: layout.getElementsByTagNameNS(P, 'ph').length };
  }, { xml, layoutXml });
  assert.deepEqual(native, { text: '', bold: '0', italic: '0', underline: 'none', strike: 'noStrike', alpha: '0', margins: ['0', '0', '0', '0'], wrap: 'none', fontScale: '100000', sharedText: 'Shared', sharedPlaceholderCount: 0, layoutPlaceholderCount: 1 });
  await page.setContent(expected.svg.slides[0].svg);
  await page.evaluate(() => document.fonts.ready);
  const marker = await page.evaluate(() => {
    const fo = [...document.querySelectorAll('foreignObject')].find(el => el.textContent.includes('Next') && el.textContent.includes('Plain'));
    const spans = [...fo.querySelectorAll('span')], bullet = spans.find(s => s.getAttribute('aria-hidden') === 'true'), next = spans.find(s => s.textContent === 'Next');
    const css = getComputedStyle(bullet), base = fo.getBoundingClientRect();
    return { markerLeft: bullet.getBoundingClientRect().left-base.left, textLeft: next.getBoundingClientRect().left-base.left, opacity: css.opacity, fontSize: css.fontSize, weight: css.fontWeight, decoration: css.textDecorationLine };
  });
  assert.deepEqual(marker, { markerLeft: 0, textLeft: 27, opacity: '1', fontSize: '24px', weight: '700', decoration: 'line-through' });
  assert.deepEqual(requests, []);
});

test('shared definition edits preserve full Node/ESM/global output parity and native ownership after rollback', async t => {
  const api = { createPresentation, normalizePresentation, resolvePresentationInheritance, resolveLayout, resolveNormalizedLayout, generatePptx, renderLayoutToSvg };
  async function snapshot(api, d) {
    const normalized = api.normalizePresentation(d), layout = api.resolveLayout(d), output = await api.generatePptx(d);
    const resolution = api.resolvePresentationInheritance(d);
    assert.deepEqual(resolution.presentation, normalized);
    assert.deepEqual(api.resolveNormalizedLayout(resolution.presentation), layout);
    return { normalized, origins: resolution.origins, layout, svg: await api.renderLayoutToSvg(layout), bytes: [...output.bytes], warnings: output.warnings };
  }
  const d = editMasterDefinitions(multiMasterDocument(createPresentation)), expected = await snapshot(api, d);
  assert.deepEqual(expected.warnings, []);
  const parts = readGeneratedZip(Uint8Array.from(expected.bytes));
  for (const i of [0, 4]) {
    const element = expected.normalized.slides[i].elements[0];
    assert.equal(element.content[0].runs[0].style.bold, false);
    assert.equal(element.content[0].runs[0].style.italic, false);
    assert.equal(element.content[0].style.spaceBefore, 0);
  }
  assert.equal([...parts.keys()].filter(n => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(n)).length, 5);
  assert.equal([...parts.keys()].filter(n => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n)).length, 5);
  for (const i of [0, 4]) { assert.ok(expected.svg.slides[i].svg.includes('#118855')); assert.ok(expected.svg.slides[i].svg.includes('Verdana')); assert.ok(expected.svg.slides[i].svg.includes('#EEFFEE')); }
  for (let i = 1; i <= 5; i++) assert.ok(!parts.get(`ppt/slides/slide${i}.xml`).toString().includes('shared-mark'));
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', r => { requests.push(r.request().url()); return r.abort(); });
  const imports = {}, globals = [];
  for (const name of ['core', 'layout', 'pptx-exporter', 'svg-renderer']) {
    imports['@pptkit/' + name] = 'data:text/javascript;base64,' + Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`, 'utf8')).toString('base64');
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`, 'utf8'));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const format of ['esm', 'global']) {
    if (format === 'global') for (const content of globals) await page.addScriptTag({ content });
    const actual = await page.evaluate(async ({ format, fixture, edit }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'layout', 'pptx-exporter', 'svg-renderer'].map(n => import('@pptkit/' + n))));
      const d = Function('createPresentation', `return (${fixture})(createPresentation)`)(api.createPresentation);
      Function('document', `return (${edit})(document)`)(d);
      const normalized = api.normalizePresentation(d), layout = api.resolveLayout(d), output = await api.generatePptx(d);
      const resolution = api.resolvePresentationInheritance(d);
      if (JSON.stringify(resolution.presentation) !== JSON.stringify(normalized)) throw new Error('Origin-enabled IR differs');
      if (!Object.isFrozen(resolution.origins) || !resolution.origins.every(r => Object.isFrozen(r.source))) throw new Error('Mutable origins');
      return { normalized, origins: resolution.origins, layout, svg: await api.renderLayoutToSvg(layout), bytes: [...output.bytes], warnings: output.warnings };
    }, { format, fixture: multiMasterDocument.toString(), edit: editMasterDefinitions.toString() });
    assert.deepEqual(actual, expected);
  }
  const truth = await page.evaluate(parts => {
    const P = 'http://schemas.openxmlformats.org/presentationml/2006/main', A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const doc = part => new DOMParser().parseFromString(parts[part], 'application/xml');
    const relation = (part, type) => [...doc(part).getElementsByTagName('Relationship')].find(n => n.getAttribute('Type').endsWith('/' + type)).getAttribute('Target');
    const resolve = (part, target) => new URL(target, 'https://source.invalid/' + part).pathname.slice(1);
    return [1, 2, 3, 4, 5].map(i => {
      const slide = `ppt/slides/slide${i}.xml`, layout = resolve(slide, relation(`ppt/slides/_rels/slide${i}.xml.rels`, 'slideLayout'));
      const master = resolve(layout, relation(layout.replace('/slideLayouts/', '/slideLayouts/_rels/') + '.rels', 'slideMaster'));
      const theme = resolve(master, relation(master.replace('/slideMasters/', '/slideMasters/_rels/') + '.rels', 'theme'));
      return { layout, master, color: doc(theme).getElementsByTagNameNS(A, 'accent1')[0].firstElementChild.getAttribute('val'), layoutShapes: doc(layout).getElementsByTagNameNS(P, 'sp').length, slideShapes: doc(slide).getElementsByTagNameNS(P, 'sp').length, masterBackgrounds: doc(master).getElementsByTagNameNS(P, 'bg').length, slideBackgrounds: doc(slide).getElementsByTagNameNS(P, 'bg').length };
    });
  }, Object.fromEntries([...parts].filter(([n]) => n.endsWith('.xml') || n.endsWith('.rels')).map(([n, b]) => [n, b.toString()])));
  assert.deepEqual(truth.map(r => r.color), ['118855', '0044CC', '112233', '112233', '118855']);
  assert.equal(truth[0].layout, truth[4].layout); assert.equal(truth[0].master, truth[4].master);
  assert.deepEqual(truth.map(r => r.layoutShapes), [2, 1, 1, 1, 2]);
  assert.deepEqual(truth.map(r => r.slideShapes), [1, 1, 1, 1, 1]);
  assert.deepEqual(truth.map(r => r.masterBackgrounds), [1, 1, 1, 1, 1]);
  assert.deepEqual(truth.map(r => r.slideBackgrounds), [0, 0, 0, 0, 0]);
  assert.deepEqual(requests, []);
});


test('chart/notes origins agree across Node/ESM/global and independently serialized theme domains', async t => {
  const expected = await textOriginSnapshot({ createPresentation, normalizePresentation, resolvePresentationInheritance, resolveNormalizedLayout, generatePptx, renderLayoutToSvg }, chartNotesOriginsDocument);
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', r => { requests.push(r.request().url()); return r.abort(); });
  const imports = {}, globals = [];
  for (const name of ['core', 'layout', 'pptx-exporter', 'svg-renderer']) {
    imports['@pptkit/' + name] = 'data:text/javascript;base64,' + Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`, 'utf8')).toString('base64');
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`, 'utf8'));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const format of ['esm', 'global']) {
    if (format === 'global') for (const content of globals) await page.addScriptTag({ content });
    const actual = await page.evaluate(async ({ format, fixture, snapshot }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'layout', 'pptx-exporter', 'svg-renderer'].map(n => import('@pptkit/' + n))));
      return Function('api', `return (${snapshot})(api, (${fixture}))`)(api);
    }, { format, fixture: chartNotesOriginsDocument.toString(), snapshot: textOriginSnapshot.toString() });
    assert.deepEqual(actual, expected);
  }
  const parts = readGeneratedZip(Uint8Array.from(expected.bytes));
  const truth = await page.evaluate(parts => {
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main', C = 'http://schemas.openxmlformats.org/drawingml/2006/chart', P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const doc = name => new DOMParser().parseFromString(parts[name], 'application/xml');
    const attributes = n => Object.fromEntries([...n.attributes].map(a => [a.name, a.value]));
    const charts = Object.keys(parts).filter(n => /^ppt\/charts\/chart\d+\.xml$/.test(n)).sort((a,b) => Number(a.match(/chart(\d+)/)[1])-Number(b.match(/chart(\d+)/)[1])).map(name => {
      const d = doc(name), category = d.getElementsByTagNameNS(C, 'catAx')[0], value = d.getElementsByTagNameNS(C, 'valAx')[0];
      const style = n => { const r = n.getElementsByTagNameNS(A, 'defRPr')[0] ?? n.getElementsByTagNameNS(A, 'rPr')[0]; return { ...attributes(r), font: r.getElementsByTagNameNS(A, 'latin')[0].getAttribute('typeface'), color: r.getElementsByTagNameNS(A, 'srgbClr')[0].getAttribute('val') }; };
      const title = d.getElementsByTagNameNS(C, 'title')[0];
      return { category: category ? style(category) : null, value: value ? style(value) : null, title: title ? style(title) : null };
    });
    const notes = [1,2,3,4].map(index => {
      const d = doc(`ppt/notesSlides/notesSlide${index}.xml`), body = [...d.getElementsByTagNameNS(P, 'sp')].find(n => n.getElementsByTagNameNS(P, 'ph')[0]?.getAttribute('type') === 'body');
      if (!body) return [];
      return [...body.getElementsByTagNameNS(A, 'r')].map(run => { const p = run.getElementsByTagNameNS(A, 'rPr')[0]; return { text: run.getElementsByTagNameNS(A, 't')[0].textContent, ...attributes(p), font: p.getElementsByTagNameNS(A, 'latin')[0].getAttribute('typeface') }; });
    });
    return { charts, notes };
  }, Object.fromEntries([...parts].filter(([n]) => n.endsWith('.xml')).map(([n,b]) => [n,b.toString()])));
  assert.equal(truth.charts.length, 11);
  for (const [index, font, color] of [[0, 'Georgia', 'CC2200'], [1, 'Verdana', '0044CC'], [2, 'Georgia', 'CC2200']]) for (let c = 0; c < 3; c++) {
    const chart = truth.charts[index*3+c];
    if (c !== 2) { assert.equal(chart.category.font, font); assert.equal(chart.category.color, color); assert.equal(chart.category.sz, '900'); assert.equal(chart.category.b, '0'); assert.equal(chart.category.i, '1'); assert.equal(chart.value.font, 'Arial'); assert.equal(chart.value.color, '595959'); assert.equal(chart.value.sz, '1400'); } else { assert.equal(chart.category, null); assert.equal(chart.value, null); }
    assert.equal(chart.title.font, font); assert.equal(chart.title.color, '112233'); assert.equal(chart.title.sz, '1200'); assert.equal(chart.title.b, '0');
  }
  for (let index = 0; index < 3; index++) {
    assert.equal(truth.notes[index][0].text, ''); assert.equal(truth.notes[index][0].b, '0'); assert.equal(truth.notes[index][0].sz, '1800');
    assert.equal(truth.notes[index][1].text, 'Notes ' + index); assert.equal(truth.notes[index][1].font, 'Tahoma'); assert.equal(truth.notes[index][1].i, '1');
  }
  assert.deepEqual(truth.notes[3], []); assert.deepEqual(requests, []);
});
