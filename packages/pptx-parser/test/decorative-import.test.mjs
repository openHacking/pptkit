import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { contentFixture, shape, body } from './content-fixture.mjs';
import { encodeStructure } from './structure-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const uri = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';
const ns = 'http://schemas.microsoft.com/office/drawing/2017/decorative';
const ext = (value = '1', extra = '') => `<a:ext uri="${uri}"><d:decorative xmlns:d="${ns}"${value === undefined ? '' : ` val="${value}"`}${extra}/></a:ext>`;
const document = (contents, strict = false) => {
  const element = shape({ txBox: true, text: body() }).replace('/><p:cNvSpPr', `><a:extLst>${contents}</a:extLst></p:cNvPr><p:cNvSpPr`);
  return encodeStructure(contentFixture(element, strict));
};

test('explicit decorative boolean projects in both dialects without changing source bytes', async () => {
  for (const strict of [false, true]) for (const [lexical, expected] of [['1', true], ['true', true], ['0', false], ['false', false]]) {
    const bytes = document(ext(lexical), strict), session = await importPptx(bytes, { projection: 'content' });
    try {
      const element = session.document.slides[0].elements[0];
      assert.equal(element.accessibility.decorative, expected);
      assert.equal(element.accessibility.description, 'Description');
      assert.ok(session.report.findings.some(f => f.feature === 'drawing.decorative' && f.tier === 'editable' && f.source.path?.includes('decorative[1]')));
      assert.ok(!session.report.findings.some(f => f.feature === 'drawing.unmapped-property' && f.source.path?.includes('extLst')));
      const ready = await session.prepareSave({ mode: 'no-op' }); assert.equal(ready.status, 'ready');
      assert.deepEqual((await session.save(ready.plan)).bytes, new Uint8Array(bytes));
    } finally { session.dispose(); }
  }
});

test('ambiguous or unknown extension syntax stays source-only and does not set a decorative flag', async () => {
  const invalid = [
    ext('1') + ext('0'), ext('', ''), ext('maybe'), ext('1', ' other="x"'),
    `<a:ext uri="${uri}"><d:decorative xmlns:d="${ns}"/></a:ext>`,
    `<a:ext uri="${uri}"><d:decorative xmlns:d="urn:other" val="1"/></a:ext>`,
    `<a:ext uri="${uri}"><d:decorative xmlns:d="${ns}" val="1"><d:other/></d:decorative></a:ext>`,
    `<a:ext uri="${uri}"><d:decorative xmlns:d="${ns}" val="1"/>unexpected</a:ext>`,
  ];
  for (const contents of invalid) {
    const session = await importPptx(document(contents), { projection: 'content' });
    try {
      assert.equal(session.document.slides[0].elements[0].accessibility.decorative, undefined);
      assert.ok(session.report.findings.some(f => f.feature === 'drawing.unmapped-property' && f.source.path?.includes('extLst')));
    } finally { session.dispose(); }
  }
});

test('unknown neighboring extensions remain reported alongside a mapped decorative flag', async () => {
  const session = await importPptx(document(`<a:ext uri="{OTHER}"><x:unknown xmlns:x="urn:other"/></a:ext>${ext()}`), { projection: 'content' });
  try {
    assert.equal(session.document.slides[0].elements[0].accessibility.decorative, true);
    assert.ok(session.report.findings.some(f => f.feature === 'drawing.unmapped-property' && f.source.path?.includes('extLst')));
  } finally { session.dispose(); }
});

test('projected decorative flag does not grant source edit-save authority', async () => {
  const bytes = document(ext());
  const session = await importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v1' });
  try {
    const slide = session.document.slides[0], original = structuredClone(slide.elements[0]);
    slide.replaceElement(original.id, { ...original, accessibility: { ...original.accessibility, decorative: false } });
    const blocked = await session.prepareSave({ mode: 'edit' });
    assert.equal(blocked.status, 'blocked');
    slide.replaceElement(original.id, original);
    const restored = await session.prepareSave({ mode: 'no-op' }); assert.equal(restored.status, 'ready');
    assert.deepEqual((await session.save(restored.plan)).bytes, new Uint8Array(bytes));
  } finally { session.dispose(); }
});

test('unsupported drawing does not claim an editable decorative projection', async () => {
  const source = shape({ preset: 'hexagon' }).replace('/><p:cNvSpPr', `><a:extLst>${ext()}</a:extLst></p:cNvPr><p:cNvSpPr`);
  const session = await importPptx(encodeStructure(contentFixture(source)), { projection: 'content' });
  try {
    assert.equal(session.document.slides[0].elements.length, 0);
    assert.ok(!session.report.findings.some(f => f.feature === 'drawing.decorative'));
    assert.ok(session.report.findings.some(f => f.feature === 'drawing.sp' && f.tier === 'unsupported'));
  } finally { session.dispose(); }
});

test('text preservation retains decorative and unknown neighboring extensions for three rounds', async () => {
  for (const strict of [false, true]) {
    let bytes = new Uint8Array(document(`<a:ext uri="{OTHER}"><x:unknown xmlns:x="urn:other"/></a:ext>${ext()}`, strict));
    for (let round = 1; round <= 3; round++) {
      const session = await importPptx(bytes, { projection: 'content', editSaveProfile: 'text-runs-v1' });
      try {
        const slide = session.document.slides[0], original = slide.elements[0], replacement = structuredClone(original);
        assert.equal(original.accessibility.decorative, true);
        const beforeText = replacement.content[0].runs[0].text;
        replacement.content[0].runs[0].text = `round-${round}`;
        slide.replaceElement(original.id, replacement);
        const plan = await session.prepareSave({ mode: 'edit' }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        const output = await session.save(plan.plan), before = readGeneratedZip(bytes), after = readGeneratedZip(output.bytes);
        for (const [name, payload] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), payload, name);
        const encodedBefore = round === 1 ? ' A &amp; B ' : beforeText;
        assert.equal(after.get('deck/z.xml').toString(), before.get('deck/z.xml').toString().replace(`>${encodedBefore}</a:t>`, `>round-${round}</a:t>`));
        bytes = output.bytes;
      } finally { session.dispose(); }
    }
    const reopened = await importPptx(bytes, { projection: 'content' });
    try {
      assert.equal(reopened.document.slides[0].elements[0].accessibility.decorative, true);
      assert.equal(reopened.document.slides[0].elements[0].content[0].runs[0].text, 'round-3');
      assert.ok(reopened.report.findings.some(f => f.feature === 'drawing.unmapped-property' && f.source.path?.includes('extLst')));
    } finally { reopened.dispose(); }
  }
});
