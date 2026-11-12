import assert from 'node:assert/strict';
import test from 'node:test';
import { createPresentation } from '@pptkit/core';
import { generatePptx } from '../dist/index.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const box = { x: 10, y: 20, width: 300, height: 80 };
const on = { bold: true, italic: true, underline: true, strike: true };
const off = { bold: false, italic: false, underline: false, strike: false };
const reset = /b="0" i="0" u="none" strike="noStrike"/;

test('layout placeholders carry paragraph/run defaults and slide-local false values explicitly cancel all four styles', async () => {
  const d = createPresentation();
  d.defineSlideLayout({ id: 'styled', name: 'Styled', placeholders: [{ key: 'body', kind: 'body', box, textStyle: { paragraph: { spaceBefore: 12, align: 'center', bullet: { type: 'number', startAt: 3 } }, run: { ...on, fontSize: 24, color: '123456' } } }], elements: [{ id: 'shared-art', type: 'text', placeholderKey: 'body', content: 'VISIBLE' }] });
  d.addSlide({ layoutId: 'styled', elements: [{ id: 'body', type: 'text', placeholderKey: 'body', content: [{ runs: [{ text: 'Plain', style: off }] }] }] });
  const output = await generatePptx(d); assert.deepEqual(output.warnings, []);
  const parts = readGeneratedZip(output.bytes), layout = [...parts].find(([name, bytes]) => name.startsWith('ppt/slideLayouts/') && name.endsWith('.xml') && bytes.toString().includes('name="Styled"'))[1].toString();
  assert.match(layout, /<a:pPr algn="ctr" marL="342900" indent="-342900">/);
  assert.match(layout, /<a:spcBef><a:spcPts val="1200"\/><\/a:spcBef>/);
  assert.match(layout, /<a:buAutoNum type="arabicPeriod" startAt="3"\/>/);
  assert.match(layout, /<a:defRPr lang="en-US" sz="2400" b="1" i="1" u="sng" strike="sngStrike">/);
  assert.match(layout, /<a:srgbClr val="123456"/);
  const placeholderShape = [...layout.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].find(match => match[0].includes('name="body"'))[0];
  assert.doesNotMatch(placeholderShape, /<a:t>/); // Shared prompt text is not copied or invented.
  const artwork = [...layout.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].find(match => match[0].includes('name="shared-art"'))[0];
  assert.match(artwork, /<a:t>VISIBLE<\/a:t>/);
  assert.doesNotMatch(artwork, /<p:ph\b/, 'Layout artwork uses placeholder geometry/style without becoming a duplicate native prompt');
  assert.equal((layout.match(/<p:ph\b/g) ?? []).length, 1);
  assert.match(parts.get('ppt/slides/slide1.xml').toString(), reset);
});

test('resolved false attributes are explicit in text, shape text, table cells and empty-note insertion styles', async () => {
  const d = createPresentation(), content = [{ runs: [{ text: '', style: { ...off, fontSize: 22 } }] }];
  d.addSlide({ notes: content, elements: [
    { id: 'text', type: 'text', box, content },
    { id: 'shape', type: 'shape', shape: 'rect', box, text: { content } },
    { id: 'table', type: 'table', box, columns: [300], rows: [{ height: 80, cells: [{ content }] }] },
  ] });
  const parts = readGeneratedZip((await generatePptx(d)).bytes);
  const slide = parts.get('ppt/slides/slide1.xml').toString(), notes = parts.get('ppt/notesSlides/notesSlide1.xml').toString();
  for (const [xml, count] of [[slide, 3], [notes, 1]]) {
    const ends = [...xml.matchAll(/<a:endParaRPr\b([^>]*)>/g)].filter(match => match[1].includes('sz="2200"'));
    assert.equal(ends.length, count);
    for (const end of ends) assert.match(end[1], reset);
  }
});

test('nonempty table paragraph marks retain their final resolved font slots without copying size or actions', async () => {
  const d = createPresentation();
  const content = [
    { runs: [{ text: 'Theme' }] },
    { runs: [{ text: 'First', style: { fontFamily: 'First Face' } }, { text: 'Final', style: { fontFamily: 'Arial', eastAsianFontFamily: 'East Face', complexScriptFontFamily: 'Complex Face', fontSize: 14 }, action: { type: 'url', url: 'https://example.com/' } }] },
  ];
  d.addSlide({ elements: [
    { id: 'table', type: 'table', box, columns: [300], rows: [{ height: 80, cells: [{ content }] }] },
    { id: 'text', type: 'text', box, content },
  ] });
  const xml = readGeneratedZip((await generatePptx(d)).bytes).get('ppt/slides/slide1.xml').toString();
  const table = xml.match(/<a:tbl>[\s\S]*?<\/a:tbl>/)[0];
  const ends = [...table.matchAll(/<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>/g)].map(m => m[0]);
  assert.equal(ends.length, 2);
  assert.match(ends[0], /<a:latin typeface="\+mn-lt"\/>/);
  assert.equal(ends[1], '<a:endParaRPr lang="en-US"><a:latin typeface="Arial"/><a:ea typeface="East Face"/><a:cs typeface="Complex Face"/></a:endParaRPr>');
  for (const end of ends) assert.doesNotMatch(end, /\bsz=|hlink|solidFill/);
  assert.equal((xml.replace(table, '').match(/<a:endParaRPr lang="en-US"\/>/g) ?? []).length, 2, 'ordinary text emission remains unchanged');
});
