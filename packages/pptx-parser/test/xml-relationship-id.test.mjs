import test from 'node:test';
import assert from 'node:assert/strict';
import { editXmlNumericAttribute, editXmlTextAttribute, editXmlTextAttributes, replaceXmlNumericAttribute, replaceXmlRelationshipId } from '../dist/preservation/replace-xml-relationship-id.js';

const relationship = 'urn:relations', element = 'urn:drawing';
const before = `<?xml version="1.0"?><root xmlns:a="${element}" xmlns:r="${relationship}" xmlns:q="${relationship}"><!-- r:embed='rId1' --><a:blip r:embed='rId1'/><a:blip q:embed = "r&#x49;d1" x="keep &amp; &gt;"/><tail/></root>`;
const expected = before.replace('q:embed = "r&#x49;d1"', 'q:embed = "rId_new"');
const request = { elementIndex: 3, elementNamespace: element, elementLocal: 'blip', attributeNamespace: relationship, attributeLocal: 'embed', expected: 'rId1', replacement: 'rId_new' };
const options = { maxOutputBytes: 8192 };
const utf16be = text => { const le = Buffer.from(text, 'utf16le'); for (let i = 0; i < le.length; i += 2) [le[i], le[i + 1]] = [le[i + 1], le[i]]; return Buffer.concat([Buffer.from([0xfe, 0xff]), le]); };
const encodings = [text => Buffer.from(text), text => Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text)]), text => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')]), utf16be];

test('relationship attribute edit retains lexical bytes and exact namespace identity in all supported encodings', async () => {
  for (const encode of encodings) {
    const source = encode(before), snapshot = Buffer.from(source);
    const result = await replaceXmlRelationshipId(source, request, options);
    assert.deepEqual(Buffer.from(result), encode(expected));
    assert.deepEqual(source, snapshot);
    const noop = await replaceXmlRelationshipId(source, { ...request, replacement: 'rId1' }, options);
    assert.deepEqual(Buffer.from(noop), snapshot);
  }
});

test('relationship edits reject stale identities, namespace mutation, unsafe IDs and resource violations', async () => {
  const input = Buffer.from(before), code = expected => error => error.code === expected;
  for (const change of [{ elementIndex: 4 }, { elementNamespace: 'urn:wrong' }, { attributeNamespace: 'urn:wrong' }, { expected: 'stale' }]) await assert.rejects(replaceXmlRelationshipId(input, { ...request, ...change }, options), code('xml-edit-conflict'));
  for (const replacement of ['a:b', 'a b', 'x"/><bad', '', '0bad', 'é']) await assert.rejects(replaceXmlRelationshipId(input, { ...request, replacement }, options), code('xml-edit-unsupported'));
  await assert.rejects(replaceXmlRelationshipId(input, { ...request, attributeNamespace: 'http://www.w3.org/2000/xmlns/' }, options), code('xml-edit-unsupported'));
  await assert.rejects(replaceXmlRelationshipId(input, request, { maxOutputBytes: 5 }), code('xml-edit-limit-exceeded'));
  await assert.rejects(replaceXmlRelationshipId(input, { ...request, replacement: 'rId1' }, { maxOutputBytes: 5 }), code('xml-edit-limit-exceeded'));
  await assert.rejects(replaceXmlRelationshipId(input, request, { ...options, limits: { maxInputBytes: 5 } }), code('xml-limit-exceeded'));
  const controller = new AbortController(); controller.abort();
  await assert.rejects(replaceXmlRelationshipId(input, request, { ...options, signal: controller.signal }), code('operation-aborted'));
  if (typeof SharedArrayBuffer !== 'undefined') await assert.rejects(replaceXmlRelationshipId(new Uint8Array(new SharedArrayBuffer(10)), request, options), /ordinary byte buffer/);
});

test('relationship edit snapshots caller bytes and scalar request before asynchronous parsing', async () => {
  const source = Buffer.from(before), mutable = { ...request };
  const pending = replaceXmlRelationshipId(source, mutable, options);
  source.fill(0); mutable.replacement = 'changed';
  assert.equal(Buffer.from(await pending).toString(), expected);
});

test('three sequential relationship edits preserve all non-target source text including inactive branches', async () => {
  let text = `<root xmlns:a="${element}" xmlns:r="${relationship}" xmlns:mc="urn:mc"><mc:AlternateContent><mc:Choice><a:blip r:embed="hidden"/></mc:Choice><mc:Fallback><a:blip r:embed='rId1'/></mc:Fallback></mc:AlternateContent><a:blip r:embed="rId1"/></root>`;
  let bytes = Buffer.from(text), old = 'rId1';
  for (const replacement of ['rId2', 'rId3', 'rId4']) {
    bytes = Buffer.from(await replaceXmlRelationshipId(bytes, { ...request, elementIndex: 6, expected: old, replacement }, options));
    text = text.replace(`r:embed='${old}'`, `r:embed='${replacement}'`);
    assert.equal(bytes.toString(), text);
    old = replacement;
  }
});

test('numeric attribute primitive preserves all other bytes in every supported encoding', async () => {
  const sourceText = `<p:presentation xmlns:p="urn:p" firstSlideNum = '+00&#49;' other="unchanged"><child firstSlideNum="1"/></p:presentation>`;
  const expectedText = sourceText.replace("firstSlideNum = '+00&#49;'", "firstSlideNum = '17'");
  const target = { elementIndex: 1, elementNamespace: 'urn:p', elementLocal: 'presentation', attributeNamespace: '', attributeLocal: 'firstSlideNum', expected: '+001', replacement: '17' };
  for (const encode of encodings) {
    const input = encode(sourceText), snapshot = Buffer.from(input), output = await replaceXmlNumericAttribute(input, target, options);
    assert.deepEqual(Buffer.from(output), encode(expectedText)); assert.deepEqual(input, snapshot);
    assert.deepEqual(Buffer.from(await replaceXmlNumericAttribute(output, { ...target, expected: '17', replacement: '17' }, options)), Buffer.from(output));
  }
});

test('numeric attribute primitive rejects stale source, unsafe lexical forms and invalid budgets', async () => {
  const input = Buffer.from('<p:presentation xmlns:p="urn:p" firstSlideNum="1"/>');
  const target = { elementIndex: 1, elementNamespace: 'urn:p', elementLocal: 'presentation', attributeNamespace: '', attributeLocal: 'firstSlideNum', expected: '1', replacement: '2' }, code = expected => error => error.code === expected;
  for (const change of [{ expected: '2' }, { elementNamespace: 'urn:wrong' }, { attributeNamespace: 'urn:wrong' }, { elementIndex: 2 }]) await assert.rejects(replaceXmlNumericAttribute(input, { ...target, ...change }, options), code('xml-edit-conflict'));
  for (const replacement of ['', '-1', '+1', '01', '1.0', '1e2', '2147483648', '1"/><bad', '１']) await assert.rejects(replaceXmlNumericAttribute(input, { ...target, replacement }, options), code('xml-edit-unsupported'));
  await assert.rejects(replaceXmlNumericAttribute(input, target, { maxOutputBytes: 5 }), code('xml-edit-limit-exceeded'));
  const controller = new AbortController(); controller.abort(); await assert.rejects(replaceXmlNumericAttribute(input, target, { ...options, signal: controller.signal }), code('operation-aborted'));
  const mutable = { ...target }, caller = Buffer.from(input), pending = replaceXmlNumericAttribute(caller, mutable, options);
  caller.fill(0); mutable.replacement = '9999';
  assert.equal(Buffer.from(await pending).toString(), '<p:presentation xmlns:p="urn:p" firstSlideNum="2"/>');
});

test('numeric edits select one namespace-bound MCE branch and preserve every other occurrence for three rounds', async () => {
  for (const namespace of ['http://schemas.openxmlformats.org/presentationml/2006/main', 'http://purl.oclc.org/ooxml/presentationml/main']) {
    let source = `<mc:AlternateContent xmlns:mc="urn:mc" xmlns:p="${namespace}"><mc:Choice><p:presentation firstSlideNum="7"/></mc:Choice><mc:Fallback><p:presentation firstSlideNum='1' data="untouched"/></mc:Fallback></mc:AlternateContent>`;
    let bytes = Buffer.from(source), old = '1';
    for (const replacement of ['0', '9999', '2147483647']) {
      bytes = Buffer.from(await replaceXmlNumericAttribute(bytes, { elementIndex: 5, elementNamespace: namespace, elementLocal: 'presentation', attributeNamespace: '', attributeLocal: 'firstSlideNum', expected: old, replacement }, options));
      source = source.replace(`firstSlideNum='${old}'`, `firstSlideNum='${replacement}'`);
      assert.equal(bytes.toString(), source); old = replacement;
    }
  }
});


test('optional numeric attribute edits reject stale presence and preserve other source bytes', async () => {
  const target = { elementIndex: 1, elementNamespace: 'urn:p', elementLocal: 'presentation', attributeNamespace: '', attributeLocal: 'firstSlideNum' };
  const source = `<p:presentation xmlns:p="urn:p" other='firstSlideNum="1"' />`, inserted = source.replace('/>', ' firstSlideNum="1"/>');
  for (const encode of encodings) {
    const input = encode(source), added = await editXmlNumericAttribute(input, { ...target, replacement: '1' }, options);
    assert.deepEqual(Buffer.from(added), encode(inserted));
    const removed = await editXmlNumericAttribute(added, { ...target, expected: '1' }, options);
    assert.deepEqual(Buffer.from(removed), encode(inserted.replace('firstSlideNum="1"/>', '/>')));
    await assert.rejects(editXmlNumericAttribute(added, { ...target, replacement: '2' }, options), e => e.code === 'xml-edit-conflict');
    await assert.rejects(editXmlNumericAttribute(input, { ...target, expected: '1' }, options), e => e.code === 'xml-edit-conflict');
    await assert.rejects(editXmlNumericAttribute(input, { ...target, attributeNamespace: 'urn:x', replacement: '1' }, options), e => e.code === 'xml-edit-unsupported');
    await assert.rejects(editXmlNumericAttribute(input, { ...target, replacement: '01' }, options), e => e.code === 'xml-edit-unsupported');
    await assert.rejects(editXmlNumericAttribute(input, { ...target, replacement: '1' }, { maxOutputBytes: input.length }), e => e.code === 'xml-edit-limit-exceeded');
  }
});

test('text attributes preserve unrelated lexical bytes and round-trip Unicode, quotes and XML whitespace in every encoding', async () => {
  const source = `<root xmlns:p="urn:p"><!-- descr="untouched" --><p:cNvPr id='7' name="kept" descr = 'old &amp; value'><extension value="untouched"/></p:cNvPr><p:cNvPr id="8" descr="neighbor"/></root>`;
  const target = { elementIndex: 2, elementNamespace: 'urn:p', elementLocal: 'cNvPr', attributeNamespace: '', attributeLocal: 'descr', expected: 'old & value' };
  const value = `中😀 & <tag> "double" 'single'\tline\nreturn\r`;
  const escaped = `中😀 &amp; &lt;tag&gt; &quot;double&quot; &apos;single&apos;&#x9;line&#xA;return&#xD;`;
  for (const encode of encodings) {
    const input = encode(source), snapshot = Buffer.from(input);
    const changed = await editXmlTextAttribute(input, { ...target, replacement: value }, options);
    assert.deepEqual(Buffer.from(changed), encode(source.replace('old &amp; value', escaped)));
    assert.deepEqual(input, snapshot);
    assert.deepEqual(Buffer.from(await editXmlTextAttribute(changed, { ...target, expected: value, replacement: value }, options)), Buffer.from(changed));
    const cleared = await editXmlTextAttribute(changed, { ...target, expected: value }, options);
    assert.deepEqual(Buffer.from(cleared), encode(source.replace(`descr = 'old &amp; value'`, '')));
    const inserted = await editXmlTextAttribute(cleared, { ...target, expected: undefined, replacement: '' }, options);
    assert.deepEqual(Buffer.from(inserted), encode(source.replace(`descr = 'old &amp; value'`, ' descr=""')));
  }
});

test('text attribute edits reject invalid Unicode, stale presence, resource overflow and cancellation', async () => {
  const input = Buffer.from('<p:cNvPr xmlns:p="urn:p" descr="value"/>');
  const target = { elementIndex: 1, elementNamespace: 'urn:p', elementLocal: 'cNvPr', attributeNamespace: '', attributeLocal: 'descr', expected: 'value' };
  for (const replacement of ['\0', '\u0001', '\uD800', '\uDC00', '\uFFFE', '\uFFFF']) await assert.rejects(editXmlTextAttribute(input, { ...target, replacement }, options), e => e.code === 'xml-edit-unsupported');
  await assert.rejects(editXmlTextAttribute(input, { ...target, expected: 'stale', replacement: 'ok' }, options), e => e.code === 'xml-edit-conflict');
  await assert.rejects(editXmlTextAttribute(input, { ...target, expected: undefined, replacement: 'ok' }, options), e => e.code === 'xml-edit-conflict');
  await assert.rejects(editXmlTextAttribute(input, { ...target, replacement: '中文😀' }, { maxOutputBytes: input.length }), e => e.code === 'xml-edit-limit-exceeded');
  const controller = new AbortController(); controller.abort();
  await assert.rejects(editXmlTextAttribute(input, { ...target, replacement: 'ok' }, { ...options, signal: controller.signal }), e => e.code === 'operation-aborted');
  await assert.rejects(editXmlTextAttribute(input, { ...target, replacement: '&'.repeat(200) }, { maxOutputBytes: 500 }), e => e.code === 'xml-edit-limit-exceeded');
  const mutable = { ...target, replacement: 'safe' }, source = Buffer.from(input), pending = editXmlTextAttribute(source, mutable, options);
  source.fill(0); mutable.replacement = 'changed';
  assert.equal(Buffer.from(await pending).toString(), '<p:cNvPr xmlns:p="urn:p" descr="safe"/>');
  const later = new AbortController();
  const longEdit = editXmlTextAttribute(input, { ...target, replacement: '&'.repeat(200000) }, { maxOutputBytes: 2_000_000, signal: later.signal });
  setTimeout(() => later.abort(), 0);
  await assert.rejects(longEdit, e => e.code === 'operation-aborted');
});

test('three text attribute edits isolate one source occurrence across inactive sibling branches', async () => {
  let source = `<mc:AlternateContent xmlns:mc="urn:mc" xmlns:p="urn:p"><mc:Choice><p:cNvPr descr="source"/></mc:Choice><mc:Fallback><p:cNvPr descr='old'/></mc:Fallback></mc:AlternateContent>`;
  let bytes = Buffer.from(source), old = 'old', lexical = 'old';
  for (const [value, encoded] of [['A&B', 'A&amp;B'], ['甲😀', '甲😀'], ['tab\tend', 'tab&#x9;end']]) {
    bytes = Buffer.from(await editXmlTextAttribute(bytes, { elementIndex: 5, elementNamespace: 'urn:p', elementLocal: 'cNvPr', attributeNamespace: '', attributeLocal: 'descr', expected: old, replacement: value }, options));
    source = source.replace(`descr='${lexical}'`, `descr='${encoded}'`);
    assert.equal(bytes.toString(), source); old = value; lexical = encoded;
  }
});

test('batch text attributes atomically preserve source bytes and support same-node insertions in every encoding', async () => {
  const source = `<root xmlns:p="urn:p"><p:x id='1' name = 'Old' descr="A &amp; B"/><p:x id="2" name="Other"/><tail keep="yes"/></root>`;
  const batch = [
    { elementIndex: 2, elementNamespace: 'urn:p', elementLocal: 'x', attributeNamespace: '', attributeLocal: 'name', expected: 'Old', replacement: 'New & 中文' },
    { elementIndex: 2, elementNamespace: 'urn:p', elementLocal: 'x', attributeNamespace: '', attributeLocal: 'descr', expected: 'A & B' },
    { elementIndex: 2, elementNamespace: 'urn:p', elementLocal: 'x', attributeNamespace: '', attributeLocal: 'first', replacement: '1' },
    { elementIndex: 2, elementNamespace: 'urn:p', elementLocal: 'x', attributeNamespace: '', attributeLocal: 'second', replacement: '2' },
    { elementIndex: 3, elementNamespace: 'urn:p', elementLocal: 'x', attributeNamespace: '', attributeLocal: 'name', expected: 'Other', replacement: '' },
  ];
  const expected = source.replace("name = 'Old' descr=\"A &amp; B\"/>", "name = 'New &amp; 中文'  first=\"1\" second=\"2\"/>").replace('name="Other"', 'name=""');
  for (const encode of encodings) { const input = encode(source), original = Buffer.from(input), result = await editXmlTextAttributes(input, batch, { maxOutputBytes: 10000, maxEdits: 10, maxWorkUnits: 10000 }); assert.deepEqual(Buffer.from(result.bytes), encode(expected)); assert.deepEqual(input, original); assert.ok(result.workUnits > 0); }
});

test('batch text attributes reject conflicts, invalid XML, aggregate output and work limits without partial output', async () => {
  const source = Buffer.from('<root><item a="old"/><item a="second"/></root>'), base = { elementIndex: 2, elementNamespace: '', elementLocal: 'item', attributeNamespace: '', attributeLocal: 'a', expected: 'old', replacement: 'new' }, options = { maxOutputBytes: 10000, maxEdits: 10, maxWorkUnits: 10000 };
  for (const changes of [[base, base], [base, { ...base, elementIndex: 3, expected: 'wrong' }]]) await assert.rejects(editXmlTextAttributes(source, changes, options), { code: 'xml-edit-conflict' });
  await assert.rejects(editXmlTextAttributes(source, [{ ...base, replacement: '\0' }], options), { code: 'xml-edit-unsupported' });
  await assert.rejects(editXmlTextAttributes(source, [base, { ...base, elementIndex: 3, expected: 'second', replacement: '01234567890123456789' }], { ...options, maxOutputBytes: source.length }), { code: 'xml-edit-limit-exceeded' });
  await assert.rejects(editXmlTextAttributes(source, [base], { ...options, maxWorkUnits: source.length }), { code: 'xml-edit-limit-exceeded' });
  await assert.rejects(editXmlTextAttributes(source, [base, { ...base, elementIndex: 3, expected: 'second' }], { ...options, maxEdits: 1 }), { code: 'xml-edit-limit-exceeded' });
  const ctrl = new AbortController(); ctrl.abort(); await assert.rejects(editXmlTextAttributes(source, [base], { ...options, signal: ctrl.signal }), { code: 'operation-aborted' });
  const mutable = { ...base }, copy = Buffer.from(source), pending = editXmlTextAttributes(copy, [mutable], options); copy.fill(0); mutable.replacement = 'changed'; assert.match(Buffer.from((await pending).bytes).toString(), /a="new"/); assert.equal(source.toString(), '<root><item a="old"/><item a="second"/></root>');
});

test('large attribute batch fits a linear work budget and retains every non-target byte', async () => {
  const n = 1000, input = Buffer.from(`<root>${Array.from({ length: n }, (_, i) => `<item id="${i}" value="old${i}" keep="&amp;"/>`).join('')}</root>`);
  const requests = Array.from({ length: n }, (_, i) => ({ elementIndex: i + 2, elementNamespace: '', elementLocal: 'item', attributeNamespace: '', attributeLocal: 'value', expected: `old${i}`, replacement: `new${i}` }));
  const result = await editXmlTextAttributes(input, requests, { maxOutputBytes: input.length * 2, maxEdits: n, maxWorkUnits: input.length * 10 });
  const expected = `<root>${Array.from({ length: n }, (_, i) => `<item id="${i}" value="new${i}" keep="&amp;"/>`).join('')}</root>`;
  assert.equal(Buffer.from(result.bytes).toString(), expected); assert.ok(result.workUnits < input.length * 10);
});

test('attribute batch inserts only the reserved xml:space qualified attribute', async () => {
  const xml = Buffer.from('<root><t>value</t><t xml:space="default">other</t></root>');
  const base = { elementNamespace: '', elementLocal: 't', attributeNamespace: 'http://www.w3.org/XML/1998/namespace', attributeLocal: 'space', replacement: 'preserve' };
  const result = await editXmlTextAttributes(xml, [{ ...base, elementIndex: 2 }, { ...base, elementIndex: 3, expected: 'default' }], { maxOutputBytes: 1000, maxEdits: 2, maxWorkUnits: 10000 });
  assert.equal(Buffer.from(result.bytes).toString(), '<root><t xml:space="preserve">value</t><t xml:space="preserve">other</t></root>');
  await assert.rejects(editXmlTextAttributes(xml, [{ ...base, elementIndex: 2, attributeLocal: 'lang' }], { maxOutputBytes: 1000, maxEdits: 2, maxWorkUnits: 10000 }), { code: 'xml-edit-unsupported' });
  await assert.rejects(editXmlTextAttributes(xml, [{ ...base, elementIndex: 2, replacement: 'invalid' }], { maxOutputBytes: 1000, maxEdits: 2, maxWorkUnits: 10000 }), { code: 'xml-edit-unsupported' });
});

test('attribute batch cancellation interrupts active large-document work', async () => {
  const count = 5000, source = Buffer.from(`<root>${'<item a="old"/>'.repeat(count)}</root>`), edits = Array.from({ length: count }, (_, i) => ({ elementIndex: i + 2, elementNamespace: '', elementLocal: 'item', attributeNamespace: '', attributeLocal: 'a', expected: 'old', replacement: 'new' }));
  const controller = new AbortController(), pending = editXmlTextAttributes(source, edits, { maxOutputBytes: source.length * 2, maxEdits: count, maxWorkUnits: source.length * 20, signal: controller.signal });
  setTimeout(() => controller.abort(), 0); await assert.rejects(pending, { code: 'operation-aborted' }); assert.equal(source.toString().includes('a="new"'), false);
});

test('seeded batch attribute edits match independent sequential edits across encodings and mixed presence states', async () => {
  let seed = 0x5f3759df; const random = n => { seed ^= seed << 13; seed ^= seed >>> 17; seed ^= seed << 5; return (seed >>> 0) % n; };
  const choices = ['plain', '', '中文😀', `&<"'\t\n`, 'same'];
  for (let sample = 0; sample < 30; sample++) {
    const nodes = [], edits = [];
    for (let i = 0; i < 6; i++) {
      const name = `name${i}`, before = random(2) ? `old${i}` : undefined, quote = random(2) ? '"' : "'";
      nodes.push(`<p:node id="${i}" name=${quote}${name}${quote}${before === undefined ? '' : ` descr=${quote}${before}${quote}`} keep='&amp;'/>`);
      const value = choices[random(choices.length)];
      edits.push({ elementIndex: i + 2, elementNamespace: 'urn:p', elementLocal: 'node', attributeNamespace: '', attributeLocal: 'name', expected: name, replacement: `${value}${i}` });
      edits.push({ elementIndex: i + 2, elementNamespace: 'urn:p', elementLocal: 'node', attributeNamespace: '', attributeLocal: 'descr', ...(before !== undefined ? { expected: before } : {}), ...(random(3) !== 0 ? { replacement: choices[random(choices.length)] } : {}) });
    }
    for (let i = edits.length - 1; i > 0; i--) { const j = random(i + 1); [edits[i], edits[j]] = [edits[j], edits[i]]; }
    const xml = `<root xmlns:p="urn:p"><!-- keep -->${nodes.join('')}<tail/></root>`;
    for (const encode of encodings) {
      const input = encode(xml), options = { maxOutputBytes: 100000, maxEdits: 100, maxWorkUnits: 100000 }, actual = await editXmlTextAttributes(input, edits, options);
      let expected = input; for (const edit of edits) expected = await editXmlTextAttribute(expected, edit, options);
      assert.deepEqual(Buffer.from(actual.bytes), Buffer.from(expected), `sample ${sample}`);
    }
  }
});

test('attribute batch interrupts validation of an oversized replacement before editing', async () => {
  const input = Buffer.from('<root a="old"/>'), original = Buffer.from(input), controller = new AbortController();
  const request = { elementIndex: 1, elementNamespace: '', elementLocal: 'root', attributeNamespace: '', attributeLocal: 'a', expected: 'old', replacement: 'A'.repeat(1_000_000) };
  const pending = editXmlTextAttributes(input, [request], { maxOutputBytes: 2_000_000, maxEdits: 1, maxWorkUnits: 10_000_000, signal: controller.signal });
  setTimeout(() => controller.abort(), 0); await assert.rejects(pending, { code: 'operation-aborted' }); assert.deepEqual(input, original);
});
