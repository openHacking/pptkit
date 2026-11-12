import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../../tests/support/read-generated-zip.mjs';

const path = process.argv[2] ?? '.tmp/bidi-direction-native/direction.pptx';
const bytes = await readFile(path);
function shapes(data) {
  const xml = readGeneratedZip(data).get('ppt/slides/slide1.xml').toString();
  return [...xml.matchAll(/<p:sp>.*?<p:cNvPr[^>]*name="([^"]+)".*?<a:pPr\b([^>]*)>.*?<a:t>([^<]*)<\/a:t>.*?<\/p:sp>/gs)].map(match => ({ name: match[1], rtl: /\brtl="([^"]+)"/.exec(match[2])?.[1], text: match[3] }));
}
const original = shapes(bytes);
assert.deepEqual(original.map(item => [item.name, item.rtl]), [['latin-first-ltr','0'],['latin-first-rtl','1'],['rtl-first-ltr','0'],['rtl-first-rtl','1']]);
assert.deepEqual(original.map(item => item.text), ['ABC (שלום 123) XYZ', 'ABC (שלום 123) XYZ', 'שלום (ABC 123) סוף', 'שלום (ABC 123) סוף']);
const session = await importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v3' });
try {
  const slide = session.document.slides[0];
  assert.equal(slide.elements.length, 4);
  assert.deepEqual(slide.elements.map(element => element.content[0].style.direction), ['ltr', 'rtl', 'ltr', 'rtl']);
  const regenerated = shapes((await generatePptx(session.document)).bytes);
  assert.equal(regenerated.length, 4);
  assert.deepEqual(regenerated.map(item => item.rtl), ['0', '1', '0', '1']);
  const declarations = session.report.textStyleInheritance.properties.filter(item => item.property === 'paragraph.direction' && item.source.part === '/ppt/slides/slide1.xml');
  assert.equal(declarations.length, 4);
  const noOp = await session.prepareSave({ mode: 'no-op' });
  assert.equal(noOp.status, 'ready'); assert.deepEqual((await session.save(noOp.plan)).bytes, new Uint8Array(bytes));
  const element = slide.elements[1], changed = structuredClone(element);
  changed.content[0].runs[0].text += '!'; slide.replaceElement(element.id, changed);
  const edit = await session.prepareSave({ mode: 'edit' }); assert.equal(edit.status, 'ready');
  const after = shapes((await session.save(edit.plan)).bytes);
  assert.deepEqual(after, original.map((item, i) => i === 1 ? { ...item, text: item.text + '!' } : item));
  console.log(JSON.stringify({ names: original.map(item => item.name), rtl: original.map(item => item.rtl), mappedSlideDeclarations: 4, freshExportRtl: 'preserved', noOp: 'byte-identical', editedRun: original[1].name, rtlPreservedAfterEdit: true }));
} finally { session.dispose(); }
