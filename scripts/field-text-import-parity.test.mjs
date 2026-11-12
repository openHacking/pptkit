import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { normalizePresentation } from '../packages/core/dist/index.js';
import { importPptx } from '../packages/pptx-parser/dist/index.js';
import { encodeStructure } from '../packages/pptx-parser/test/structure-fixture.mjs';
import { notesFixture, notesShape } from '../packages/pptx-parser/test/notes-fixture.mjs';
import { body } from '../packages/pptx-parser/test/content-fixture.mjs';

test('opt-in page-number field projection and no-op agree across Node and browser entries', async t => {
  const id = '38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A', field = `<a:p><a:r><a:t>Page </a:t></a:r><a:fld id="{${id}}" type="slidenum"><a:t>42</a:t></a:fld><a:fld id="{${id}}" type="datetime3"><a:t>date cache</a:t></a:fld><a:fld id="{${id}}" type="datetime13"><a:t>time cache</a:t></a:fld></a:p>`, bytes = encodeStructure(notesFixture(notesShape({ text: body(field) })));
  const options = { projection: 'content', fieldTextProfile: 'date-time-and-number-v1', editSaveProfile: 'text-runs-v2' }, session = await importPptx(bytes, options);
  let expected;
  try { const plan = await session.prepareSave({ mode: 'no-op' }), document = normalizePresentation(session.document), report = session.report, original = Array.from((await session.save(plan.plan)).bytes), slide = session.document.slides.find(s => Array.isArray(s.notes) && s.notes[0]?.runs.some(run => run.field)), notes = structuredClone(slide.notes), bad = structuredClone(notes); bad[0].runs[1].text = '99'; slide.setNotes(bad); const blocked = await session.prepareSave({ mode: 'edit' }); notes[0].runs[0].text = 'Edited '; slide.setNotes(notes); const edited = await session.prepareSave({ mode: 'edit' }); assert.equal(edited.status, 'ready'); expected = { document, report, bytes: original, blocked: blocked.status, saved: Array.from((await session.save(edited.plan)).bytes) }; } finally { session.dispose(); }
  const core = await readFile('packages/core/dist/browser/index.js', 'utf8'), coreUrl = `data:text/javascript;base64,${Buffer.from(core).toString('base64')}`;
  const parser = (await readFile('packages/pptx-parser/dist/browser/index.js', 'utf8')).replaceAll('"@pptkit/core"', JSON.stringify(coreUrl)), parserUrl = `data:text/javascript;base64,${Buffer.from(parser).toString('base64')}`;
  const browser = await chromium.launch(); t.after(() => browser.close()); const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  for (const name of ['core', 'pptx-parser']) await page.addScriptTag({ content: await readFile(`packages/${name}/dist/browser/global.js`, 'utf8') });
  const results = await page.evaluate(async ({ coreUrl, parserUrl, bytes, options }) => {
    const core = await import(coreUrl), parser = await import(parserUrl), results = [];
    for (const [c, p] of [[core, parser], [PPTKit, PPTKit]]) { const session = await p.importPptx(new Uint8Array(bytes), options); try { const plan = await session.prepareSave({ mode: 'no-op' }), document = c.normalizePresentation(session.document), report = session.report, original = Array.from((await session.save(plan.plan)).bytes), slide = session.document.slides.find(s => Array.isArray(s.notes) && s.notes[0]?.runs.some(run => run.field)), notes = structuredClone(slide.notes), bad = structuredClone(notes); bad[0].runs[1].text = '99'; slide.setNotes(bad); const blocked = await session.prepareSave({ mode: 'edit' }); notes[0].runs[0].text = 'Edited '; slide.setNotes(notes); const edited = await session.prepareSave({ mode: 'edit' }); if (edited.status !== 'ready') throw Error(JSON.stringify(edited.report)); results.push({ document, report, bytes: original, blocked: blocked.status, saved: Array.from((await session.save(edited.plan)).bytes) }); } finally { session.dispose(); } }
    return results;
  }, { coreUrl, parserUrl, bytes: Array.from(bytes), options });
  assert.equal(expected.blocked, 'blocked'); for (const result of results) assert.deepEqual(result, expected); assert.deepEqual(requests, []);
});
