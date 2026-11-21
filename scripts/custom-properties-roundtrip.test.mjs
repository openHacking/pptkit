import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { createPresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { inspectPptx, importPptx } from "../packages/pptx-parser/dist/index.js";
import { readGeneratedZip } from "../tests/support/read-generated-zip.mjs";

test("fresh custom properties survive export and independent parser inspection", async () => {
  const customProperties = [
    { name: "Text\t& <", type: "string", value: "中文\r & 😀" }, { name: "Flag", type: "boolean", value: false },
    { name: "Count", type: "integer", value: -2147483648 }, { name: "Ratio", type: "real", value: 1.25 },
    { name: "When", type: "date", value: "2020-01-02T03:04:05.000Z" }, { name: "Negative zero", type: "real", value: -0 },
  ];
  const doc = createPresentation({ metadata: { customProperties } });
  const output = await generatePptx(doc), inspected = await inspectPptx(output.bytes);
  assert.deepEqual(inspected.metadata.custom.properties.map(p => [p.id, p.name, p.valueType, p.value.value]), [
    [2, "Text\t& <", "lpwstr", "中文\r & 😀"], [3, "Flag", "bool", "false"], [4, "Count", "i4", "-2147483648"],
    [5, "Ratio", "r8", "1.25"], [6, "When", "filetime", "2020-01-02T03:04:05.000Z"], [7, "Negative zero", "r8", "-0"],
  ]);
  const session = await importPptx(output.bytes, { projection: "structure" });
  try {
    assert.deepEqual(session.document.metadata.customProperties, customProperties);
    assert.equal(session.report.findings.filter(f => f.feature === "metadata.custom-property" && f.tier === "editable").length, 6);
    const prepared = await session.prepareSave({ mode: "no-op" });
    assert.equal(prepared.status, "ready");
    assert.deepEqual((await session.save(prepared.plan)).bytes, output.bytes);
    session.document.setMetadata({ ...session.document.metadata, customProperties: [{ name: "changed", type: "string", value: "new" }] });
    const edit = await session.prepareSave({ mode: "edit" });
    assert.equal(edit.status, "blocked");
    assert.equal(edit.plan, undefined);
  } finally { session.dispose(); }
  for (const input of [undefined, []]) {
    const generated = await generatePptx(createPresentation({ metadata: input === undefined ? {} : { customProperties: input } }));
    const metadata = (await inspectPptx(generated.bytes)).metadata;
    if (input === undefined) assert.equal(metadata.custom, undefined);
    else assert.deepEqual(metadata.custom.properties, []);
  }
});


test("custom property authoring updates and removals produce exactly one coherent package relationship", async () => {
  const document = createPresentation({ metadata: { title: "Keep title", customProperties: [
    { name: "Alpha", type: "string", value: "first" }, { name: "Beta", type: "boolean", value: true },
  ] } });
  const states = [
    [{ name: "Alpha", type: "string", value: "first" }, { name: "Beta", type: "boolean", value: true }],
    [{ name: "Beta", type: "boolean", value: false }, { name: "Gamma", type: "integer", value: 17 }],
    [{ name: "Gamma", type: "integer", value: -1 }],
    [],
    undefined,
  ];
  for (const properties of states) {
    const next = { ...document.metadata };
    if (properties === undefined) delete next.customProperties; else next.customProperties = properties;
    document.setMetadata(next);
    for (let round = 0; round < 3; round++) {
      const output = await generatePptx(document), parts = readGeneratedZip(output.bytes);
      const relationships = parts.get("_rels/.rels").toString(), types = parts.get("[Content_Types].xml").toString();
      const exists = properties !== undefined;
      assert.equal(parts.has("docProps/custom.xml"), exists);
      assert.equal((relationships.match(/relationships\/custom-properties/g) ?? []).length, exists ? 1 : 0);
      assert.equal((types.match(/PartName="\/docProps\/custom.xml"/g) ?? []).length, exists ? 1 : 0);
      const inspected = await inspectPptx(output.bytes);
      assert.equal(inspected.metadata.core.properties.find(p => p.name === "title").value, "Keep title");
      if (!exists) assert.equal(inspected.metadata.custom, undefined);
      else {
        assert.deepEqual(inspected.metadata.custom.properties.map(p => p.name), properties.map(p => p.name));
        assert.deepEqual(inspected.metadata.custom.properties.map(p => p.id), properties.map((_, index) => index + 2));
        const imported = await importPptx(output.bytes, { projection: "structure" });
        try { assert.deepEqual(imported.document.metadata.customProperties, properties); } finally { imported.dispose(); }
      }
    }
  }
});


test("three fresh export and structural reimport rounds preserve modeled custom values", async () => {
  const values = [{ name: "text", type: "string", value: "x\r\n中文" }, { name: "flag", type: "boolean", value: false }, { name: "zero", type: "real", value: -0 }, { name: "date", type: "date", value: "2020-01-02T03:04:05.000Z" }];
  let bytes = (await generatePptx(createPresentation({ metadata: { customProperties: values } }))).bytes;
  for (let round = 0; round < 3; round++) {
    const session = await importPptx(bytes, { projection: "structure" });
    try {
      assert.deepEqual(session.document.metadata.customProperties, values);
      bytes = (await generatePptx(session.document)).bytes;
    } finally { session.dispose(); }
    const inspected = await inspectPptx(bytes);
    assert.equal(inspected.metadata.custom.properties.length, values.length);
    assert.deepEqual(inspected.metadata.custom.properties.map(p => p.name), values.map(v => v.name));
  }
});


test("literal escape-like names and strings round-trip without being interpreted", async () => {
  for (const literal of ["_x0041_", "_x0000_", "prefix _x00aF_ suffix", "_x0041__x0042_"]) {
    const property = { name: `name ${literal}`, type: 'string', value: literal };
    const bytes = (await generatePptx(createPresentation({ metadata: { customProperties: [property] } }))).bytes;
    const session = await importPptx(bytes, { projection: 'structure' });
    try { assert.deepEqual(session.document.metadata.customProperties, [property]); } finally { session.dispose(); }
  }
});


test("overlapping literal escapes fail closed when application resaves cannot preserve them", async () => {
  for (const property of [{ name: '_x005F_x0041_', type: 'string', value: 'safe' }, { name: 'safe', type: 'string', value: '_x005F_x0041_' }]) {
    await assert.rejects(generatePptx(createPresentation({ metadata: { customProperties: [property] } })), { name: 'RangeError', message: 'Overlapping OOXML escape-like sequences are not supported for custom properties.' });
  }
});

test('fresh authoring evidence matches public API outputs and retained application inventories', async()=>{
 const evidence=JSON.parse(await readFile(new URL('../docs/testing/fresh-properties-libreoffice-evidence.json',import.meta.url),'utf8')),hash=bytes=>createHash('sha256').update(bytes).digest('hex');
 const properties=[{name:'Empty',type:'string',value:''},{name:'Flag',type:'boolean',value:false},{name:'Count',type:'integer',value:-2147483648},{name:'Real',type:'real',value:1.25},{name:'Date',type:'date',value:'2020-01-02T03:04:05.000Z'},{name:'Decimal',type:'decimal',value:'9007199254740993.0001'},{name:'Control',type:'string',value:'before'}];
 const doc=createPresentation({metadata:{customProperties:properties}});doc.addSlide({id:'fresh-properties'});assert.equal(hash((await generatePptx(doc)).bytes),evidence.records.find(r=>r.kind==='control'&&r.round===0).sha256);
 const updated=[{name:'Flag',type:'boolean',value:true},{name:'Count',type:'integer',value:2147483647},{name:'Real',type:'real',value:.125},{name:'Date',type:'date',value:'2021-02-03T04:05:06.000Z'},{name:'Decimal',type:'decimal',value:'9007199254740994.0001'},{name:'Control',type:'string',value:'after'}];doc.setMetadata({...doc.metadata,customProperties:updated});assert.equal(hash((await generatePptx(doc)).bytes),evidence.records.find(r=>r.kind==='edited'&&r.round===0).sha256);
 doc.setMetadata({...doc.metadata,customProperties:[]});assert.equal(hash((await generatePptx(doc)).bytes),evidence.auxiliaries.find(r=>r.kind==='empty').sha256);const metadata={...doc.metadata};delete metadata.customProperties;doc.setMetadata(metadata);assert.equal(hash((await generatePptx(doc)).bytes),evidence.auxiliaries.find(r=>r.kind==='absent').sha256);
 for(const kind of ['control','edited']){const records=evidence.records.filter(r=>r.kind===kind);assert.deepEqual(records.map(r=>r.round),[0,1,2,3]);for(const r of records){const names=r.properties.map(p=>p.name);assert.equal(new Set(names).size,names.length);}assert.ok(records.slice(1).every(r=>!r.properties.some(p=>p.name==='Decimal')));assert.ok(records.slice(1).every(r=>r.properties.find(p=>p.name==='Control')?.value===(kind==='control'?'before':'after')));}
});
