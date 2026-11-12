import test from 'node:test';
import assert from 'node:assert/strict';
import { appendPropertyDeclarations as append } from '../dist/preservation/append-property-declarations.js';
const relNs='http://schemas.openxmlformats.org/package/2006/relationships', typesNs='http://schemas.openxmlformats.org/package/2006/content-types', core='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', mime='application/vnd.openxmlformats-package.core-properties+xml';
const req={kind:'core',part:'/pptkit-core.xml',relationshipId:'core1',dialect:'transitional'}, options={maxOutputBytes:8192};
const rel=`<r:Relationships xmlns:r="${relNs}"><!--stay--></r:Relationships><!--tail-->`, types=`<t:Types xmlns:t="${typesNs}"><t:Default Extension="xml" ContentType="application/xml"/></t:Types>`;
const encoders=[s=>Buffer.from(s),s=>Buffer.concat([Buffer.from([0xef,0xbb,0xbf]),Buffer.from(s)]),s=>Buffer.concat([Buffer.from([0xff,0xfe]),Buffer.from(s,'utf16le')]),s=>{const b=Buffer.from(s,'utf16le');for(let i=0;i<b.length;i+=2)[b[i],b[i+1]]=[b[i+1],b[i]];return Buffer.concat([Buffer.from([0xfe,0xff]),b]);}];
const code=x=>e=>e.code===x;
test('core OPC declarations retain source bytes and encoding across dialects',async()=>{
 for(const dialect of ['transitional','strict'])for(const encode of encoders){const a=encode(rel),b=encode(types),out=await append(a,b,{...req,dialect},options);assert.deepEqual(Buffer.from(out.relationships),encode(rel.replace('</r:Relationships>',`<r:Relationship Id="core1" Type="${core}" Target="pptkit-core.xml"/></r:Relationships>`)));assert.deepEqual(Buffer.from(out.contentTypes),encode(types.replace('</t:Types>',`<t:Override PartName="/pptkit-core.xml" ContentType="${mime}"/></t:Types>`)));}
});
test('core declarations reject raw relationship, ID, MIME and part collisions',async()=>{
 const wrapRel=x=>Buffer.from(`<Relationships xmlns="${relNs}">${x}</Relationships>`),wrapTypes=x=>Buffer.from(`<Types xmlns="${typesNs}">${x}</Types>`);
 for(const child of [`<Relationship Id="other" Type="${core}" Target="other"/>`,`<Relationship Id="core1" Type="urn:other" Target="other"/>`,`<unknown><Relationship Id="other" Type="${core}" Target="other"/></unknown>`])await assert.rejects(append(wrapRel(child),wrapTypes(''),req,options),code('xml-edit-conflict'));
 for(const child of [`<Override PartName="/PPTKIT-CORE.XML" ContentType="application/xml"/>`,`<Override PartName="/other.xml" ContentType="${mime.toUpperCase()}"/>`,`<Default Extension="xml" ContentType="${mime}"/>`])await assert.rejects(append(wrapRel(''),wrapTypes(child),req,options),code('xml-edit-conflict'));
});
test('core declarations reject unsafe names, MCE, exhausted budgets and cancellation',async()=>{
 const a=Buffer.from(`<Relationships xmlns="${relNs}"/>`),b=Buffer.from(`<Types xmlns="${typesNs}"/>`);
 for(const part of ['/a/../b.xml','/a%20b.xml','/a&b.xml'])await assert.rejects(append(a,b,{...req,part},options),code('xml-edit-unsupported'));
 await assert.rejects(append(Buffer.from(`<Relationships xmlns="${relNs}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>`),b,req,options),code('xml-edit-unsupported'));
 await assert.rejects(append(a,b,req,{maxOutputBytes:2}),code('xml-edit-limit-exceeded'));const controller=new AbortController();controller.abort();await assert.rejects(append(a,b,req,{...options,signal:controller.signal}),code('operation-aborted'));
});


test('core declarations coexist with existing custom-property declarations',async()=>{
 const custom='http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
 const a=Buffer.from(`<Relationships xmlns="${relNs}"><Relationship Id="custom1" Type="${custom}" Target="custom.xml"/></Relationships>`),b=Buffer.from(`<Types xmlns="${typesNs}"><Override PartName="/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>`);
 const out=await append(a,b,req,options);assert.match(Buffer.from(out.relationships).toString(),/Id="custom1"/);assert.match(Buffer.from(out.relationships).toString(),/Id="core1"/);assert.match(Buffer.from(out.contentTypes).toString(),/custom-properties\+xml/);assert.match(Buffer.from(out.contentTypes).toString(),/core-properties\+xml/);
});
