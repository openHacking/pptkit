import test from 'node:test';
import assert from 'node:assert/strict';
import { createPresentation, normalizePresentation } from '../dist/index.js';

test('metadata replacement is validated, detached, immutable and atomic', () => {
  const deck = createPresentation({ metadata: { title: 'Old', author: 'A', revision: 7, keywords: ['one'] } }), input = { title: 'New', author: 'A', revision: 7, keywords: ['one', 'two'] };
  deck.setMetadata(input); input.title = 'Mutated'; input.keywords.push('three'); assert.equal(deck.metadata.title, 'New'); assert.deepEqual(deck.metadata.keywords, ['one', 'two']); assert.ok(Object.isFrozen(deck.metadata)); assert.ok(Object.isFrozen(deck.metadata.keywords)); assert.equal(normalizePresentation(deck).metadata.revision, 7);
  for (const bad of [null, [], new Date(), { title: 2 }, { author: false }, { keywords: [1] }, { keywords: new Array(1) }, { revision: -1 }, { revision: 1.2 }, { unknown: 'value' }]) { assert.throws(() => deck.setMetadata(bad), TypeError); assert.equal(deck.metadata.title, 'New'); assert.equal(deck.metadata.revision, 7); }
  deck.setMetadata({ title: '' }); assert.equal(deck.metadata.title, ''); assert.equal(deck.metadata.author, undefined);
});


test("custom properties validate types, identity and atomic snapshots", () => {
  const properties = [
    { name: "Text", type: "string", value: "中文" }, { name: "Flag", type: "boolean", value: false },
    { name: "Count", type: "integer", value: -2147483648 }, { name: "Ratio", type: "real", value: 1.25 },
    { name: "When", type: "date", value: "2020-01-02T03:04:05.000Z" },
  ];
  const deck = createPresentation({ metadata: { customProperties: properties } });
  properties[0].value = "changed";
  assert.equal(deck.metadata.customProperties[0].value, "中文");
  assert.ok(Object.isFrozen(deck.metadata.customProperties[0]));
  assert.deepEqual(normalizePresentation(deck).metadata.customProperties, deck.metadata.customProperties);
  const invalid = [new Array(1), [{ name: "", type: "string", value: "" }], [{ name: "a", type: "string", value: "" }, { name: "a", type: "boolean", value: true }], [{ name: "a", type: "integer", value: 2147483648 }], [{ name: "a", type: "real", value: Infinity }], [{ name: "a", type: "date", value: "2020-02-30T00:00:00.000Z" }], [{ name: "a", type: "date", value: "2020-01-01T00:00:00Z" }], [{ name: "a", type: "string", value: "x", extra: 1 }]];
  for (const customProperties of invalid) {
    assert.throws(() => deck.setMetadata({ customProperties }), TypeError);
    assert.equal(deck.metadata.customProperties[0].value, "中文");
    assert.throws(() => createPresentation({ metadata: { customProperties } }), TypeError);
  }
  deck.setMetadata({ customProperties: [] });
  assert.deepEqual(deck.metadata.customProperties, []);
});


test("custom property snapshots are validated after stateful input access", () => {
  let reads = 0;
  const metadata = { get customProperties() {
    return ++reads <= 2 ? [{ name: "ok", type: "integer", value: 1 }] : [{ name: "bad", type: "integer", value: 1.5 }];
  } };
  assert.throws(() => createPresentation({ metadata }), TypeError);
  const document = createPresentation({ metadata: { title: "unchanged" } });
  reads = 0;
  assert.throws(() => document.setMetadata(metadata), TypeError);
  assert.deepEqual(document.metadata, { title: "unchanged" });
});


test('core date fields are canonical, detached and never inferred',()=>{
 const dates={createdAt:'2020-02-29T23:59:59.123Z',modifiedAt:'2024-01-02T03:04:05.000Z'}, deck=createPresentation({metadata:dates});dates.createdAt='changed';
 assert.equal(deck.metadata.createdAt,'2020-02-29T23:59:59.123Z');assert.equal(normalizePresentation(deck).metadata.modifiedAt,'2024-01-02T03:04:05.000Z');
 for(const value of ['2021-02-29T00:00:00.000Z','2020-01-01T00:00:00Z','2020-01-01T00:00:00.000+00:00','bad',null,123]) {assert.throws(()=>deck.setMetadata({...deck.metadata,createdAt:value}),TypeError);assert.throws(()=>createPresentation({metadata:{createdAt:value}}),TypeError);assert.equal(deck.metadata.createdAt,'2020-02-29T23:59:59.123Z');}
 deck.setMetadata({});assert.equal(normalizePresentation(deck).metadata.createdAt,undefined);assert.equal(normalizePresentation(deck).metadata.modifiedAt,undefined);
});

test('exact decimal custom properties are canonical, bounded and atomic',()=>{
 const values=['0','0.1','-0.1','9007199254740993.0001','-12345678901234567890.123456789'];
 const doc=createPresentation({metadata:{customProperties:values.map((value,i)=>({name:`D${i}`,type:'decimal',value}))}});
 assert.deepEqual(normalizePresentation(doc).metadata.customProperties.map(p=>p.value),values);
 for(const value of ['-0','01','1.0','1.','+1','.1','1e3','-0.0','0\n','0\u2028','0'.repeat(10001),12]){assert.throws(()=>doc.setMetadata({customProperties:[{name:'D',type:'decimal',value}]}),TypeError);assert.deepEqual(doc.metadata.customProperties.map(p=>p.value),values);}
});
