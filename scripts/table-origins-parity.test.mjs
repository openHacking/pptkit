import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation, normalizePresentation, resolvePresentationInheritance } from '../packages/core/dist/index.js';
import { resolveNormalizedLayout } from '../packages/layout/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { renderLayoutToSvg } from '../packages/svg-renderer/dist/index.js';
import { tableOriginsDocument } from '../tests/fixtures/table-origins.mjs';
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';

async function snapshot(api, fixture) {
  const d = fixture(api.createPresentation), r = api.resolvePresentationInheritance(d), normalized = api.normalizePresentation(d);
  if (JSON.stringify(r.presentation) !== JSON.stringify(normalized)) throw Error('Traced IR changed values');
  if (!Object.isFrozen(r.origins) || !r.origins.every(o => Object.isFrozen(o.source))) throw Error('Mutable origins');
  const layout = api.resolveNormalizedLayout(normalized), out = await api.generatePptx(d);
  return { normalized, origins: r.origins, layout, svg: await api.renderLayoutToSvg(layout), bytes: [...out.bytes], warnings: out.warnings };
}

test('table sources, clear barriers, effective output and bytes match Node/ESM/global and independent native/SVG DOM', async t => {
  const expected = await snapshot({ createPresentation, normalizePresentation, resolvePresentationInheritance, resolveNormalizedLayout, generatePptx, renderLayoutToSvg }, tableOriginsDocument);
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', r => { requests.push(r.request().url()); return r.abort(); });
  const imports = {}, globals = [];
  for (const name of ['core','layout','pptx-exporter','svg-renderer']) {
    imports['@pptkit/'+name] = 'data:text/javascript;base64,'+Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`,'utf8')).toString('base64');
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`,'utf8'));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const format of ['esm','global']) {
    if (format === 'global') for (const content of globals) await page.addScriptTag({ content });
    const actual = await page.evaluate(async ({ format, fixture, snapshot }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core','layout','pptx-exporter','svg-renderer'].map(n => import('@pptkit/'+n))));
      return Function('api', `return (${snapshot})(api, (${fixture}))`)(api);
    }, { format, fixture: tableOriginsDocument.toString(), snapshot: snapshot.toString() });
    assert.deepEqual(actual, expected);
  }
  const parts = readGeneratedZip(Uint8Array.from(expected.bytes));
  const truth = await page.evaluate(({ xml, svg }) => {
    const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const d = new DOMParser().parseFromString(xml, 'application/xml'), p = new DOMParser().parseFromString(svg, 'image/svg+xml');
    const attrs = node => Object.fromEntries([...node.attributes].map(a => [a.name,a.value]));
    const rows = [...d.getElementsByTagNameNS(A,'tr')];
    const cell = rows[0].getElementsByTagNameNS(A,'tc')[0], style = cell.getElementsByTagNameNS(A,'tcPr')[0];
    const direct = (n, name) => [...n.children].find(c => c.namespaceURI===A && c.localName===name);
    const lines = Object.fromEntries(['lnL','lnR','lnT','lnB'].map(k => { const n=direct(style,k); return [k, { ...attrs(n), color: n.getElementsByTagNameNS(A,'srgbClr')[0]?.getAttribute('val'), alpha: n.getElementsByTagNameNS(A,'alpha')[0]?.getAttribute('val') ?? '100000' }]; }));
    const runs = [...cell.getElementsByTagNameNS(A,'r')].map(n => ({ text: n.getElementsByTagNameNS(A,'t')[0].textContent, ...attrs(n.getElementsByTagNameNS(A,'rPr')[0]) }));
    const preview = p.querySelector('[data-pptkit-cell="0:0"]');
    const implicit = p.querySelector('[data-pptkit-cell="2:0"]');
    return { cell: attrs(style), noFill: !!direct(style,'noFill'), lines, runs, previewStyle: preview.getAttribute('style'), previewFrameStyle: preview.firstElementChild.getAttribute('style'), previewText: [...preview.querySelectorAll('span[lang]')].map(n=>n.textContent).join(''), implicitText: [...implicit.querySelectorAll('span[lang]')].map(n=>n.textContent).join(''), rowCount: rows.length };
  }, { xml: parts.get('ppt/slides/slide1.xml').toString(), svg: expected.svg.slides[0].svg });
  assert.equal(truth.rowCount, 4); assert.equal(truth.noFill, true);
  assert.equal(truth.cell.marL,'0'); assert.equal(truth.cell.marR,'0'); assert.equal(truth.cell.marT,'101600'); assert.equal(truth.cell.anchor,'b');
  // Origins describe logical cell candidates; Layout chooses the shared physical edge.
  assert.deepEqual(Object.fromEntries(Object.entries(truth.lines).map(([key,line]) => [key,line.w])), { lnL: '0', lnR: '12700', lnT: '0', lnB: '38100' });
  for (const edge of Object.values(expected.normalized.slides[0].elements[0].rows[0].cells[0].style.borders)) assert.equal(edge.width, 0);
  assert.equal(truth.lines.lnT.color,'FF0000'); assert.equal(truth.lines.lnT.alpha,'100000');
  assert.equal(truth.lines.lnR.color,'111111'); assert.equal(truth.lines.lnR.alpha,'40000');
  assert.equal(truth.runs[0].b,'0'); assert.equal(truth.runs[0].text,''); assert.equal(truth.runs[1].text,'Merged'); assert.equal(truth.runs[1].i,'0');
  assert.equal(truth.previewText,'Merged'); assert.equal(truth.implicitText,''); assert.match(truth.previewFrameStyle,/padding:8px 0px 8px 0px/);
  assert.equal(expected.warnings.filter(w=>w.code==='table-style-materialized').length,2);
  assert.deepEqual(requests,[]);
});
