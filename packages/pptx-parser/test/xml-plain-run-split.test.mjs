import assert from 'node:assert/strict';
import test from 'node:test';
import { captureXmlTextSource } from '../dist/preservation/xml-text-source.js';
import { replaceArchiveEntries } from '../dist/preservation/replace-archive-entries.js';
import { importPptx } from '../dist/index.js';
import { contentFixture, shape, body } from './content-fixture.mjs';
import { encodeStructure } from './structure-fixture.mjs';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

const ns = 'urn:a';
const encode = (text, encoding) => encoding === 'utf8' ? Buffer.from(text) : encoding === 'utf8bom' ? Buffer.concat([Buffer.from([239, 187, 191]), Buffer.from(text)]) : Buffer.concat([
  Buffer.from(encoding === 'utf16le' ? [255, 254] : [254, 255]),
  encoding === 'utf16le' ? Buffer.from(text, 'utf16le') : Buffer.from(text, 'utf16le').swap16(),
]);
const decode = (bytes, encoding) => new TextDecoder(encoding === 'utf8' || encoding === 'utf8bom' ? 'utf-8' : encoding === 'utf16le' ? 'utf-16le' : 'utf-16be').decode(bytes);

test('plain run split preserves lexical surroundings and encoding in a private XML part', async () => {
  const run = "<a:r><a:t xml:space='preserve'>A&amp;B🙂</a:t></a:r>";
  const xml = `<?xml version="1.0"?><a:p xmlns:a="${ns}"><!--before-->${run}<x:future xmlns:x="urn:x"/><a:endParaRPr/></a:p>`;
  const expected = xml.replace(run, "<a:r><a:t xml:space='preserve'>A&amp;</a:t></a:r><a:r><a:t xml:space='preserve'>B🙂</a:t></a:r>");
  for (const encoding of ['utf8', 'utf8bom', 'utf16le', 'utf16be']) {
    const bytes = encode(xml, encoding), source = await captureXmlTextSource(bytes);
    const request = { elementIndex: 3, namespaceUri: ns, expectedText: 'A&B🙂', offset: 2 };
    const result = await source.splitSimpleRun(request);
    assert.equal(decode(result.bytes, encoding), expected);
    assert.deepEqual(Buffer.from(result.bytes), encode(expected, encoding));
    if (encoding === 'utf8bom') assert.deepEqual([...result.bytes.slice(0, 3)], [239, 187, 191]);
    if (encoding !== 'utf8' && encoding !== 'utf8bom') assert.deepEqual([...result.bytes.slice(0, 2)], encoding === 'utf16le' ? [255, 254] : [254, 255]);
    assert.equal(source.readElement({ elementIndex: 2, namespaceUri: ns, local: 'r' }), run);
    const exact = await source.splitSimpleRun(request, { maxOutputBytes: result.bytes.length });
    assert.deepEqual(exact.bytes, result.bytes);
    await assert.rejects(source.splitSimpleRun(request, { maxOutputBytes: result.bytes.length - 1 }), { code: 'xml-edit-limit-exceeded' });
    await assert.rejects(source.splitSimpleRun(request, { maxWorkUnits: 1 }), { code: 'xml-edit-limit-exceeded' });
    const reopened = await captureXmlTextSource(result.bytes);
    assert.deepEqual(reopened.elements.filter(item => item.local === 't').map(item => item.text), ['A&', 'B🙂']);
    result.bytes[0] ^= 0xff;
    assert.deepEqual(Buffer.from(exact.bytes), encode(expected, encoding));
    assert.equal(source.readElement({ elementIndex: 2, namespaceUri: ns, local: 'r' }), run);
    reopened.dispose(); source.dispose();
  }
});

test('simple run split rejects unsupported formatting, opaque markup, invalid ancestry and lifecycle failures', async () => {
  for (const run of [
    '<a:r><a:rPr u="sng"/><a:t>ab</a:t></a:r>',
    '<a:r><a:rPr b="1"><a:hlinkClick r:id="rId1" xmlns:r="urn:r"/></a:rPr><a:t>ab</a:t></a:r>',
    '<a:r><a:rPr b="1"><!--opaque--></a:rPr><a:t>ab</a:t></a:r>',
    '<a:r custom="opaque"><a:t>ab</a:t></a:r>',
    '<a:r><a:t custom="opaque">ab</a:t></a:r>',
    '<a:r><!--opaque--><a:t>ab</a:t></a:r>',
    '<a:r><a:t>ab</a:t><a:br/></a:r>',
  ]) {
    const source = await captureXmlTextSource(encode(`<a:p xmlns:a="${ns}">${run}</a:p>`, 'utf8'));
    const atom = source.elements.find(item => item.local === 't');
    await assert.rejects(source.splitSimpleRun({ elementIndex: atom.elementIndex, namespaceUri: ns, expectedText: 'ab', offset: 1 }), { code: 'xml-edit-unsupported' });
    source.dispose();
  }
  const source = await captureXmlTextSource(encode(`<a:p xmlns:a="${ns}"><a:r><a:t>e\u0301🙂</a:t></a:r></a:p>`, 'utf8'));
  const request = { elementIndex: 3, namespaceUri: ns, expectedText: 'e\u0301🙂', offset: 2 };
  await assert.rejects(source.splitSimpleRun({ ...request, offset: 1 }), { code: 'split-invalid' });
  await assert.rejects(source.splitSimpleRun({ ...request, expectedText: 'wrong' }), { code: 'xml-edit-conflict' });
  const controller = new AbortController(); controller.abort();
  await assert.rejects(source.splitSimpleRun(request, { signal: controller.signal }), { code: 'operation-aborted' });
  source.dispose();
  await assert.rejects(source.splitSimpleRun(request), { code: 'xml-source-disposed' });
  const longText = 'a'.repeat(20000), pendingSource = await captureXmlTextSource(encode(`<a:p xmlns:a="${ns}"><a:r><a:t>${longText}</a:t></a:r></a:p>`, 'utf8'));
  const pending = pendingSource.splitSimpleRun({ elementIndex: 3, namespaceUri: ns, expectedText: longText, offset: 19999 });
  pendingSource.dispose();
  await assert.rejects(pending, { code: 'xml-source-disposed' });
  const bounded = await captureXmlTextSource(encode(`<a:p xmlns:a="${ns}"><a:r><a:t>ab</a:t></a:r></a:p>`, 'utf8'), { limits: { maxNodes: 4 } });
  await assert.rejects(bounded.splitSimpleRun({ elementIndex: 3, namespaceUri: ns, expectedText: 'ab', offset: 1 }), { code: 'xml-limit-exceeded' });
  bounded.dispose();
});

test('simple run split duplicates selected direct scalar formatting without changing siblings', async () => {
  const properties = `<a:rPr b='1' i="0" sz='2400' lang="en-US"/>`;
  const run = `<a:r>${properties}<a:t xml:space="preserve">A&amp;B🙂</a:t></a:r>`;
  const xml = `<a:p xmlns:a="${ns}"><a:pPr algn="ctr"/>${run}<a:endParaRPr lang="fr-FR"/></a:p>`;
  const expected = xml.replace(run, `<a:r>${properties}<a:t xml:space="preserve">A&amp;</a:t></a:r><a:r>${properties}<a:t xml:space="preserve">B🙂</a:t></a:r>`);
  for (const encoding of ['utf8', 'utf8bom', 'utf16le', 'utf16be']) {
    const source = await captureXmlTextSource(encode(xml, encoding));
    const atom = source.elements.find(item => item.local === 't');
    const result = await source.splitSimpleRun({ elementIndex: atom.elementIndex, namespaceUri: ns, expectedText: 'A&B🙂', offset: 2 });
    assert.deepEqual(Buffer.from(result.bytes), encode(expected, encoding));
    const reopened = await captureXmlTextSource(result.bytes);
    assert.equal(reopened.elements.filter(item => item.local === 'rPr').length, 2);
    assert.deepEqual(reopened.elements.filter(item => item.local === 't').map(item => item.text), ['A&', 'B🙂']);
    reopened.dispose(); source.dispose();
  }
});

test('three successive plain XML splits preserve source order and allow independent recapture', async () => {
  let bytes = encode(`<a:p xmlns:a="${ns}"><a:r><a:t>abcd</a:t></a:r></a:p>`, 'utf8');
  for (const [index, text, offset] of [[3, 'abcd', 2], [3, 'ab', 1], [7, 'cd', 1]]) {
    const source = await captureXmlTextSource(bytes);
    const result = await source.splitSimpleRun({ elementIndex: index, namespaceUri: ns, expectedText: text, offset });
    source.dispose(); bytes = result.bytes;
  }
  const reopened = await captureXmlTextSource(bytes);
  assert.deepEqual(reopened.elements.filter(item => item.local === 't').map(item => item.text), ['a', 'b', 'c', 'd']);
  reopened.dispose();
});

test('equal text in separate runs demonstrates why raw index alone is not owner authorization', async () => {
  const xml = `<a:p xmlns:a="${ns}"><a:r><a:t>ab</a:t></a:r><a:r><a:t>ab</a:t></a:r></a:p>`;
  const source = await captureXmlTextSource(encode(xml, 'utf8'));
  const result = await source.splitSimpleRun({ elementIndex: 5, namespaceUri: ns, expectedText: 'ab', offset: 1 });
  assert.equal(decode(result.bytes, 'utf8'), `<a:p xmlns:a="${ns}"><a:r><a:t>ab</a:t></a:r><a:r><a:t>a</a:t></a:r><a:r><a:t>b</a:t></a:r></a:p>`);
  source.dispose();
});

test('manual package rehearsal reimports three simple splits without changing other parts', async () => {
  const props = '<a:rPr b="1" sz="2400" lang="en-US"/>';
  const originalRun = `<a:r>${props}<a:t>abcd</a:t></a:r>`;
  for (const strict of [false, true]) {
    const fixture = contentFixture(shape({ txBox: true, text: body(`<a:p>${originalRun}<a:endParaRPr lang="en-US"/></a:p>`), props: '<a:effectLst><a:glow rad="12700"><a:srgbClr val="112233"/></a:glow></a:effectLst>' }), strict);
    const sourceBytes = encodeStructure(fixture), initial = Buffer.from(sourceBytes);
    let bytes = sourceBytes;
    for (const [round, text, offset, expectedRuns] of [[1, 'abcd', 2, ['ab', 'cd']], [2, 'ab', 1, ['a', 'b', 'cd']], [3, 'cd', 1, ['a', 'b', 'c', 'd']]]) {
      const before = readGeneratedZip(bytes), payload = before.get('deck/z.xml');
      const xml = await captureXmlTextSource(payload);
      const atom = xml.elements.find(item => item.local === 't' && item.text === text);
      assert.ok(atom, `round ${round} lacks the expected atom`);
      const rawRun = xml.readElement({ elementIndex: atom.parentElementIndex, namespaceUri: atom.namespaceUri, local: 'r' });
      const split = await xml.splitSimpleRun({ elementIndex: atom.elementIndex, namespaceUri: atom.namespaceUri, expectedText: text, offset });
      xml.dispose();
      const halves = [text.slice(0, offset), text.slice(offset)];
      const expectedXml = payload.toString().replace(rawRun, halves.map(value => rawRun.replace(`>${text}</a:t>`, `>${value}</a:t>`)).join(''));
      assert.equal(Buffer.from(split.bytes).toString(), expectedXml);
      const candidate = await replaceArchiveEntries(bytes, [{ name: 'deck/z.xml', bytes: split.bytes }], { maxOutputBytes: 8 * 1024 * 1024 });
      const after = readGeneratedZip(candidate);
      assert.equal(after.size, before.size);
      for (const [name, part] of before) if (name !== 'deck/z.xml') assert.deepEqual(after.get(name), part, name);
      const session = await importPptx(candidate, { projection: 'content', editSaveProfile: 'text-runs-v1' });
      try {
        const runs = session.document.slides[0].elements[0].content[0].runs;
        assert.deepEqual(runs.map(run => run.text), expectedRuns);
        assert.ok(runs.every(run => run.style.bold === true && run.style.fontSize === 24));
        assert.equal(session.report.textRuns.length, expectedRuns.length);
        assert.ok(session.report.textRuns.every(binding => binding.status === 'bound'));
        const noop = await session.prepareSave({ mode: 'no-op' });
        assert.equal(noop.status, 'ready');
        assert.deepEqual(Buffer.from((await session.save(noop.plan)).bytes), Buffer.from(candidate));
        noop.plan.dispose();
      } finally { session.dispose(); }
      bytes = candidate;
    }
    assert.deepEqual(Buffer.from(sourceBytes), initial);
  }
});
