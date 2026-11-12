import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../../tests/support/read-generated-zip.mjs';
const path = process.argv[2] ?? '.tmp/bidi-run-native/run-rtl.pptx';
const bytes = await readFile(path);
function inspect(data) {
  const xml = readGeneratedZip(data).get('ppt/slides/slide1.xml').toString();
  return [...xml.matchAll(/<p:sp>.*?<p:cNvPr[^>]*name="([^"]+)"(.*?)<\/p:sp>/gs)].map(match => ({ name: match[1], runs: [...match[2].matchAll(/<a:r>(.*?)<\/a:r>/gs)].map(run => ({ text: /<a:t[^>]*>(.*?)<\/a:t>/s.exec(run[1])?.[1], rtl: /<a:rtl\b/.test(run[1]) })) }));
}
const original = inspect(bytes);
assert.deepEqual(original.map(item => item.name), ['ltr-control', 'ltr-run-rtl', 'rtl-control', 'rtl-run-rtl']);
for (const item of original) {
  assert.deepEqual(item.runs.map(run => run.text), ['ABC ', 'مرحبا', ' - ', '123', ' XYZ']);
  assert.deepEqual(item.runs.map(run => run.rtl), item.name.endsWith('-run-rtl') ? [false, true, true, true, false] : [false, false, false, false, false]);
}
const session = await importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v3' });
try {
  const slide = session.document.slides[0], elements = slide.elements;
  assert.equal(elements.length, 4);
  assert.deepEqual(elements.map(element => element.content[0].style.direction), ['ltr', 'ltr', 'rtl', 'rtl']);
  const declarations = session.report.textStyleInheritance.properties.filter(item => item.property === 'run.rightToLeft' && item.source.part === '/ppt/slides/slide1.xml');
  assert.equal(declarations.length, 6); assert.ok(declarations.every(item => item.status === 'mapped' && item.value === true));
  const noOp = await session.prepareSave({ mode: 'no-op' }); assert.equal(noOp.status, 'ready'); assert.deepEqual((await session.save(noOp.plan)).bytes, new Uint8Array(bytes));
  const fresh = inspect((await generatePptx(session.document)).bytes);
  assert.deepEqual(fresh.map(item => item.runs.map(run => run.rtl)), original.map(item => item.runs.map(run => run.rtl)));
  const changed = structuredClone(elements[1]); changed.content[0].runs[2].text = ' + '; slide.replaceElement(elements[1].id, changed);
  const edit = await session.prepareSave({ mode: 'edit' }); assert.equal(edit.status, 'ready');
  const after = inspect((await session.save(edit.plan)).bytes);
  const expected = structuredClone(original); expected[1].runs[2].text = ' + ';
  assert.deepEqual(after, expected);
  console.log(JSON.stringify({ shapes: original.length, rtlRuns: 6, mappedDeclarations: declarations.length, freshExportRtl: 'preserved', noOp: 'byte-identical', editedRun: 'ltr-run-rtl/2', rtlPreserved: true }));
} finally { session.dispose(); }
