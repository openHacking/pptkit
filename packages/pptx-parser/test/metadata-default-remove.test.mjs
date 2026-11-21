import test from 'node:test';
import assert from 'node:assert/strict';
import {importPptx} from '../dist/index.js';
import {metadataFixture,encodeMetadata} from './metadata-fixture.mjs';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';
const config={projection:'content',metadataSaveProfile:'metadata-default-remove-v17'},fields=['title','author','language','revision'];
const defaults={title:'PPTKit Presentation',author:'PPTKit',language:'en-US',revision:1};

test('explicitly removes and reinserts defaulted metadata across three reopens in both dialects',async()=>{
 for(const strict of [false,true])for(const initial of [defaults,{title:'Title',author:'Author',language:'fr-FR',revision:7}]){const f=metadataFixture(strict);f.add('core',`<d:title>${initial.title}</d:title><d:creator>${initial.author}</d:creator><!--sentinel--><d:language>${initial.language}</d:language><c:revision>${initial.revision}</c:revision><d:subject>S</d:subject>`);let bytes=encodeMetadata(f);
  for(let round=0;round<3;round++){const before=readGeneratedZip(bytes),s=await importPptx(bytes,config);try{if(round!==1)for(const field of fields)assert.equal(s.report.coreMetadataSources.find(x=>x.field===field).removable,true);s.document.setMetadata({...s.document.metadata,...(round===1?defaults:{title:undefined,author:undefined,language:undefined,revision:undefined})});const p=await s.prepareSave(round===1?{mode:'edit',coreMetadataInsertions:fields}:{mode:'edit',coreMetadataRemovals:fields});assert.equal(p.status,'ready',JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;const after=readGeneratedZip(bytes);for(const [name,payload] of before)if(name!=='inventory/core.data')assert.deepEqual(after.get(name),payload,name);assert.match(after.get('inventory/core.data').toString(),/<!--sentinel-->/);}finally{s.dispose();}const r=await importPptx(bytes,config);try{for(const field of fields)assert.equal(r.document.metadata[field],round===1?defaults[field]:undefined);assert.equal(r.document.metadata.subject,'S');}finally{r.dispose();}}
 }
});

test('explicit removal combines with ordinary insertions and blocks conflicting candidates',async()=>{
 const f=metadataFixture();f.add('core','<d:title>Title</d:title><d:creator>Author</d:creator><d:language>fr-FR</d:language><c:revision>7</c:revision><d:subject>Old</d:subject>');const s=await importPptx(encodeMetadata(f),config);try{s.document.setMetadata({...s.document.metadata,title:'Contradiction',author:undefined,language:undefined,revision:undefined});let p=await s.prepareSave({mode:'edit',coreMetadataRemovals:fields});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);s.document.setMetadata({...s.document.metadata,title:undefined,subject:undefined,createdAt:'2024-01-01T00:00:00.000Z'});p=await s.prepareSave({mode:'edit',coreMetadataRemovals:fields});assert.equal(p.status,'ready',JSON.stringify(p.report));const r=await importPptx((await s.save(p.plan)).bytes,config);try{for(const field of fields)assert.equal(r.document.metadata[field],undefined);assert.equal(r.document.metadata.subject,undefined);assert.equal(r.document.metadata.createdAt,'2024-01-01T00:00:00.000Z');}finally{r.dispose();}}finally{s.dispose();}
});

test('old profile, unbound targets, overlapping requests and invalid modes are rejected',async()=>{
 const f=metadataFixture();f.add('core','<d:title>Title</d:title>');const bytes=encodeMetadata(f),old=await importPptx(bytes,{projection:'content',metadataSaveProfile:'metadata-revision-remove-v16'});try{await assert.rejects(old.prepareSave({mode:'edit',coreMetadataRemovals:['title']}),TypeError);}finally{old.dispose();}
 const s=await importPptx(bytes,config);try{await assert.rejects(s.prepareSave({mode:'no-op',coreMetadataRemovals:['title']}),TypeError);await assert.rejects(s.prepareSave({mode:'edit',coreMetadataRemovals:['title','title']}),TypeError);await assert.rejects(s.prepareSave({mode:'edit',coreMetadataRemovals:['title'],coreMetadataInsertions:['title']}),TypeError);s.document.setMetadata({...s.document.metadata,title:undefined});const p=await s.prepareSave({mode:'edit',coreMetadataRemovals:['title','author']});assert.equal(p.status,'blocked');assert.equal(p.plan,undefined);}finally{s.dispose();}
});

test('removal respects abort, stale plans and work limits',async()=>{
 const f=metadataFixture();f.add('core','<d:title>Title</d:title>');const bytes=encodeMetadata(f),s=await importPptx(bytes,config);try{s.document.setMetadata({...s.document.metadata,title:undefined});const controller=new AbortController();controller.abort();await assert.rejects(s.prepareSave({mode:'edit',coreMetadataRemovals:['title'],signal:controller.signal}),{code:'operation-aborted'});const p=await s.prepareSave({mode:'edit',coreMetadataRemovals:['title']});assert.equal(p.status,'ready');s.document.setMetadata({...s.document.metadata,title:'Changed'});await assert.rejects(s.save(p.plan),{code:'save-snapshot-stale'});}finally{s.dispose();}const limited=await importPptx(bytes,{...config,sessionLimits:{maxSaveAnalysisWork:1}});try{limited.document.setMetadata({...limited.document.metadata,title:undefined});await assert.rejects(limited.prepareSave({mode:'edit',coreMetadataRemovals:['title']}),{code:'save-limit-exceeded'});}finally{limited.dispose();}
});

test('default removal and custom-property creation remain atomic over three reopens', async () => {
  for (const strict of [false, true]) {
    const f = metadataFixture(strict);
    f.add('core', '<d:title>Original</d:title><d:creator>Author</d:creator><d:language>fr-FR</d:language><c:revision>7</c:revision><!--sentinel--><d:subject>Control</d:subject>');
    f.parts['opaque.bin'] = new Uint8Array([0, 1, 2, 255]); f.metadataTypes['opaque.bin'] = 'application/octet-stream';
    let bytes = encodeMetadata(f);
    const options = { ...config, customPropertySaveProfile: 'create-and-append-v5' };
    for (let round = 0; round < 3; round++) {
      const before = readGeneratedZip(bytes), session = await importPptx(bytes, options);
      try {
        const property = { name: 'Count', type: 'integer', value: 40 + round };
        session.document.setMetadata({ ...session.document.metadata, ...(round === 1 ? defaults : { title: undefined, author: undefined, language: undefined, revision: undefined }), customProperties: [property] });
        const edit = round === 0 ? { coreMetadataRemovals: fields, customPropertyAdditions: [0] } : round === 1 ? { coreMetadataInsertions: fields } : { coreMetadataRemovals: fields };
        if (round === 0) {
          const blocked = await session.prepareSave({ mode: 'edit', ...edit, customPropertyAdditions: [] });
          assert.equal(blocked.status, 'blocked'); assert.equal(blocked.plan, undefined);
        }
        const plan = await session.prepareSave({ mode: 'edit', ...edit }); assert.equal(plan.status, 'ready', JSON.stringify(plan.report));
        bytes = (await session.save(plan.plan)).bytes;
        const after = readGeneratedZip(bytes);
        const changed = round === 0 ? ['inventory/core.data', '_rels/.rels', '[Content_Types].xml'] : ['inventory/core.data', 'pptkit-custom-properties-1.xml'];
        for (const [name, payload] of before) if (!changed.includes(name)) assert.deepEqual(after.get(name), payload, name);
        assert.match(after.get('inventory/core.data').toString(), /<!--sentinel-->/);
        assert.ok(after.has('pptkit-custom-properties-1.xml'));
        assert.match(after.get('pptkit-custom-properties-1.xml').toString(), new RegExp(`>${40 + round}</`));
      } finally { session.dispose(); }
      const reopened = await importPptx(bytes, options);
      try {
        for (const field of fields) assert.equal(reopened.document.metadata[field], round === 1 ? defaults[field] : undefined);
        assert.equal(reopened.document.metadata.subject, 'Control');
        assert.deepEqual(reopened.document.metadata.customProperties, [{ name: 'Count', type: 'integer', value: 40 + round }]);
      } finally { reopened.dispose(); }
    }
  }
});
