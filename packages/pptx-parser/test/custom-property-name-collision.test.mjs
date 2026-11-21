import test from 'node:test';
import assert from 'node:assert/strict';
import { importPptx } from '../dist/index.js';
import { metadataFixture, encodeMetadata, custom } from './metadata-fixture.mjs';

for (const strict of [false, true]) for (const other of ['opaque', 'linked', 'scalar']) test(`decoded names include ${other} sources when detecting collisions (strict=${strict})`, async () => {
  const f = metadataFixture(strict);
  const second = other === 'opaque' ? custom(3, 'vector', '<v:lpstr>opaque</v:lpstr>', 'name="_x005f_x0041_"') : custom(3, 'lpwstr', 'other', `name="_x005f_x0041_"${other === 'linked' ? ' linkTarget="elsewhere"' : ''}`);
  f.add('custom', custom(2, 'lpwstr', 'first', 'name="_x005F_x0041_"') + second + custom(4, 'lpwstr', 'kept', 'name="Control"'));
  const bytes = encodeMetadata(f), session = await importPptx(bytes, { projection: 'content', customPropertySaveProfile: 'append-and-remove-v4' });
  try {
    assert.deepEqual(session.document.metadata.customProperties, [{ name:'Control', type:'string', value:'kept' }]);
    assert.deepEqual(session.report.findings.filter(item => item.feature === 'metadata.custom-property').map(item => item.tier), ['unsupported','unsupported','editable']);
    assert.equal(session.report.customPropertySources.filter(item => item.projectedIndex !== undefined).length, 1);
    const no = await session.prepareSave({mode:'no-op'}); assert.equal(no.status, 'ready');
    assert.deepEqual(Buffer.from((await session.save(no.plan)).bytes), Buffer.from(bytes));
    session.document.setMetadata({...session.document.metadata, customProperties:[...session.document.metadata.customProperties,{name:'_x0041_',type:'string',value:'new'}]});
    assert.equal((await session.prepareSave({mode:'edit',customPropertyAdditions:[1]})).status, 'blocked');
  } finally { session.dispose(); }
});

test('currency opt-in also rejects a collision with an opaque source', async () => {
  const f = metadataFixture(); f.add('custom', custom(2, 'cy', '1.2500', 'name="_x005F_x0041_"') + custom(3, 'vector', '<v:lpstr>x</v:lpstr>', 'name="_x005f_x0041_"'));
  const session = await importPptx(encodeMetadata(f), {projection:'content', customCurrencyProjection:'exact-v1'});
  try {
    assert.deepEqual(session.document.metadata.customProperties, []);
    assert.deepEqual(session.report.findings.filter(item => item.feature === 'metadata.custom-property').map(item => item.tier), ['unsupported','unsupported']);
    assert.ok(!session.report.findings.some(item => item.feature === 'metadata.custom-property-currency'));
  } finally { session.dispose(); }
});

for (const strict of [false, true]) test(`unsupported escape names reserve their decoded spelling (strict=${strict})`, async () => {
  const f = metadataFixture(strict);
  f.add('custom', custom(2, 'lpwstr', 'first', 'name="A"') + custom(3, 'vector', '<v:lpstr>opaque</v:lpstr>', 'name="_x0041_"') + custom(4, 'lpwstr', 'kept', 'name="B"'));
  const bytes = encodeMetadata(f), session = await importPptx(bytes, {projection:'content', customPropertySaveProfile:'append-and-remove-v4'});
  try {
    assert.deepEqual(session.document.metadata.customProperties, [{name:'B',type:'string',value:'kept'}]);
    assert.deepEqual(session.report.findings.filter(item => item.feature === 'metadata.custom-property').map(item => item.tier), ['unsupported','unsupported','editable']);
    session.document.setMetadata({...session.document.metadata, customProperties:[...session.document.metadata.customProperties,{name:'A',type:'string',value:'new'}]});
    assert.equal((await session.prepareSave({mode:'edit',customPropertyAdditions:[1]})).status, 'blocked');
  } finally { session.dispose(); }
});

test('escaped literal underscore is decoded only once for collision detection', async () => {
  const f=metadataFixture();
  f.add('custom', custom(2,'lpwstr','one','name="A"') + custom(3,'lpwstr','two','name="_x005F_x0041_"'));
  const session=await importPptx(encodeMetadata(f),{projection:'content'});
  try { assert.deepEqual(session.document.metadata.customProperties.map(p=>p.name), ['A','_x0041_']); } finally {session.dispose();}
});

test('UTF-16 surrogate escapes collide with the equivalent Unicode name', async () => {
  const f=metadataFixture(); f.add('custom', custom(2,'lpwstr','a','name="😀"') + custom(3,'vector','<v:lpstr>x</v:lpstr>','name="_xD83D__xDE00_"'));
  const session=await importPptx(encodeMetadata(f),{projection:'content'});
  try { assert.deepEqual(session.document.metadata.customProperties, []); } finally { session.dispose(); }
});
