import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { chromium } from '@playwright/test';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';
import { assessPptxCurrencyText, assessPptxCurrencyNumber, importPptx, inspectPptxOpaqueCustomProperties } from '../packages/pptx-parser/dist/index.js';
import { metadataFixture, encodeMetadata, custom } from '../packages/pptx-parser/test/metadata-fixture.mjs';

async function exercise(api) {
  const customProperties = [
    { name: 'text\tname', type: 'string', value: '中文\r & 😀' },
    { name: 'flag', type: 'boolean', value: false },
    { name: 'count', type: 'integer', value: -2147483648 },
    { name: 'real', type: 'real', value: -0 },
    { name: 'date', type: 'date', value: '2020-01-02T03:04:05.000Z' },
  ];
  const doc = api.createPresentation({ metadata: { title: 'before', customProperties } });
  const generated = await api.generatePptx(doc);
  const session = await api.importPptx(generated.bytes, { projection: 'content', customPropertySaveProfile: 'append-and-remove-v4', metadataSaveProfile: 'core-strings-v1' });
  try {
    const prepare = await session.prepareSave({ mode: 'no-op' });
    if (prepare.status !== 'ready') throw new Error('no-op save blocked');
    const saved = await session.save(prepare.plan);
    const metadata = session.document.metadata;
    const removed = session.report.customPropertySources.find(item => item.projectedIndex === 2).source.sourceObjectId;
    const remaining = metadata.customProperties.filter((_, index) => index !== 2).map(p => p.name === 'text\tname' ? { ...p, name: 'renamed & _x0041_', value: 'changed & _x0041_' } : p);
    session.document.setMetadata({ ...metadata, title: 'after', customProperties: [...remaining, { name: 'added', type: 'integer', value: 42 }] });
    const edit = await session.prepareSave({ mode: 'edit', customPropertyRemovals: [removed], customPropertyAdditions: [remaining.length] });
    if (edit.status !== 'ready') throw new Error('custom edit blocked: ' + JSON.stringify(edit.report));
    const changed = await session.save(edit.plan);
    return { bytes: [...generated.bytes], saved: [...saved.bytes], changed: [...changed.bytes], metadata, sources: session.report.customPropertySources, findings: session.report.findings.filter(f => f.feature === 'metadata.custom-property') };
  } finally { session.dispose(); }
}

async function exerciseCreation(api) {
  const doc = api.createPresentation({ metadata: { title: 'before' } });
  const generated = await api.generatePptx(doc); let bytes = generated.bytes;
  const rounds = [];
  for (let round = 0; round < 3; round++) {
    const session = await api.importPptx(bytes, { projection: 'content', customPropertySaveProfile: 'create-and-append-v5' });
    try {
      const properties = session.document.metadata.customProperties ?? [];
      session.document.setMetadata({ ...session.document.metadata, customProperties: [...properties, { name: `property ${round}`, type: 'string', value: `中文 & ${round}` }] });
      const prepared = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [round] });
      if (prepared.status !== 'ready') throw new Error('custom creation blocked: ' + JSON.stringify(prepared.report));
      bytes = (await session.save(prepared.plan)).bytes;
      rounds.push({ bytes: [...bytes], insertion: session.report.customPropertyInsertion });
    } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' });
  try { return { original: [...generated.bytes], rounds, properties: reopened.document.metadata.customProperties }; }
  finally { reopened.dispose(); }
}

async function exerciseRevision(api) {
  const generated = await api.generatePptx(api.createPresentation({ metadata: { title: 'before', revision: 7 } }));
  let bytes = generated.bytes; const rounds = [];
  for (const revision of [8, 42, Number.MAX_SAFE_INTEGER]) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'core-values-v2' });
    try {
      session.document.setMetadata({ ...session.document.metadata, revision });
      const prepared = await session.prepareSave({ mode: 'edit' });
      if (prepared.status !== 'ready') throw new Error('revision edit blocked: ' + JSON.stringify(prepared.report));
      bytes = (await session.save(prepared.plan)).bytes;
      rounds.push({ bytes: [...bytes], binding: session.report.coreMetadataSources.find(item => item.field === 'revision') });
    } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' });
  try { return { rounds, revision: reopened.document.metadata.revision }; }
  finally { reopened.dispose(); }
}

async function exerciseKeywords(api) {
  const generated = await api.generatePptx(api.createPresentation({ metadata: { keywords: ['first, second'], revision: 7 } }));
  let bytes = generated.bytes; const rounds = [];
  for (const value of ['中文 & text', '', 'last']) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'core-keywords-v3' });
    try {
      session.document.setMetadata({ ...session.document.metadata, keywords: [value] });
      const prepared = await session.prepareSave({ mode: 'edit' });
      if (prepared.status !== 'ready') throw new Error('keyword edit blocked: ' + JSON.stringify(prepared.report));
      bytes = (await session.save(prepared.plan)).bytes;
      rounds.push({ bytes: [...bytes], binding: session.report.coreMetadataSources.find(item => item.field === 'keywords') });
    } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' });
  try { return { rounds, keywords: reopened.document.metadata.keywords }; }
  finally { reopened.dispose(); }
}

async function exerciseCompany(api) {
  const generated = await api.generatePptx(api.createPresentation({ metadata: { company: 'Old Company' } }));
  let bytes = generated.bytes; const rounds = [];
  for (const value of ['中文 & company', '', 'Final']) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'extended-company-v4' });
    try {
      session.document.setMetadata({ ...session.document.metadata, company: value });
      const prepared = await session.prepareSave({ mode: 'edit' });
      if (prepared.status !== 'ready') throw new Error('company edit blocked: ' + JSON.stringify(prepared.report));
      bytes = (await session.save(prepared.plan)).bytes;
      rounds.push({ bytes: [...bytes], binding: session.report.coreMetadataSources.find(item => item.field === 'company') });
    } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' });
  try { return { rounds, company: reopened.document.metadata.company }; }
  finally { reopened.dispose(); }
}

async function exerciseCoreInsertion(api, input) {
  let bytes = new Uint8Array(input); const rounds = [];
  for (let round = 0; round < 3; round++) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'core-insert-v5' });
    try {
      if (round > 0) session.document.setMetadata({ ...session.document.metadata, title: `Round ${round} 中文`, subject: `Subject ${round}` });
      const prepared = await session.prepareSave(round === 0 ? { mode: 'edit', coreMetadataInsertions: ['title', 'author', 'language'] } : { mode: 'edit' });
      if (prepared.status !== 'ready') throw new Error('core insertion blocked: ' + JSON.stringify(prepared.report));
      bytes = (await session.save(prepared.plan)).bytes;
      rounds.push({ bytes: [...bytes], changes: prepared.plan.report.changes, sources: session.report.coreMetadataSources });
    } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' });
  try { return { rounds, metadata: reopened.document.metadata }; } finally { reopened.dispose(); }
}

async function exerciseSameValueReplacement(api, input) {
  const session = await api.importPptx(new Uint8Array(input), { projection: 'content', customPropertySaveProfile: 'append-and-remove-v4' });
  try {
    const sourceObjectId = session.report.customPropertySources.find(item => item.projectedIndex === 0).source.sourceObjectId;
    const prepared = await session.prepareSave({ mode: 'edit', customPropertyRemovals: [sourceObjectId], customPropertyAdditions: [0] });
    if (prepared.status !== 'ready') throw new Error('same-value replacement blocked: ' + JSON.stringify(prepared.report));
    const bytes = (await session.save(prepared.plan)).bytes;
    const reopened = await api.importPptx(bytes, { projection: 'content' });
    try { return { bytes: [...bytes], changes: prepared.plan.report.changes, properties: reopened.document.metadata.customProperties, ids: reopened.report.inspection.metadata.custom.properties.map(item => item.id) }; }
    finally { reopened.dispose(); }
  } finally { session.dispose(); }
}

async function exerciseCoreDates(api, input) {
  const authored = api.createPresentation({ metadata: { createdAt: '2020-02-29T12:34:56.123Z', modifiedAt: '2021-01-02T03:04:05.000Z' } });
  const fresh = await api.generatePptx(authored); let bytes = new Uint8Array(input), rounds = [];
  for (const createdAt of ['2022-01-01T00:00:00.000Z', '2022-01-02T00:00:00.000Z', '2022-01-03T00:00:00.000Z']) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'core-date-values-v6' });
    try { session.document.setMetadata({ ...session.document.metadata, createdAt }); const p = await session.prepareSave({ mode: 'edit' }); if (p.status !== 'ready') throw new Error(JSON.stringify(p.report)); bytes = (await session.save(p.plan)).bytes; rounds.push({ bytes: [...bytes], changes: p.plan.report.changes }); }
    finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' }); const freshImport = await api.importPptx(fresh.bytes, { projection: 'content' });
  try { return { fresh: [...fresh.bytes], authored: freshImport.document.metadata, rounds, metadata: reopened.document.metadata }; } finally { reopened.dispose(); freshImport.dispose(); }
}

async function exerciseCoreDateInsertion(api, input) {
  let bytes = new Uint8Array(input); const rounds = [];
  for (let round = 0; round < 3; round++) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'core-date-insert-v7' });
    try { session.document.setMetadata({ ...session.document.metadata, createdAt: `2022-01-0${round + 1}T00:00:00.000Z`, modifiedAt: '2023-01-01T00:00:00.000Z' }); const plan = await session.prepareSave({ mode: 'edit' }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report)); bytes = (await session.save(plan.plan)).bytes; rounds.push({ bytes: [...bytes], sources: session.report.coreMetadataSources, changes: plan.plan.report.changes }); }
    finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' }); try { return { rounds, metadata: reopened.document.metadata }; } finally { reopened.dispose(); }
}

async function exerciseCoreCreation(api, input) {
  let bytes = new Uint8Array(input); const rounds = [];
  for (let round = 0; round < 3; round++) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'core-keywords-insert-v10', customPropertySaveProfile: 'create-and-append-v5' });
    try {
      session.document.setMetadata({ ...session.document.metadata, title: `Title ${round}`, createdAt: `2020-01-0${round + 1}T00:00:00.000Z`, revision: round + 1, keywords: [round === 0 ? '' : `keyword ${round}`], ...(round === 0 ? { customProperties: [{ name: 'A', type: 'integer', value: 3 }] } : {}) });
      const plan = await session.prepareSave(round === 0 ? { mode: 'edit', customPropertyAdditions: [0], coreMetadataInsertions: ['revision'] } : { mode: 'edit' });
      if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report)); bytes = (await session.save(plan.plan)).bytes; rounds.push({ bytes: [...bytes], changes: plan.plan.report.changes });
    } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' }); try { return { rounds, metadata: reopened.document.metadata }; } finally { reopened.dispose(); }
}

async function exerciseCompanyInsertion(api, input) {
  let bytes = new Uint8Array(input); const rounds = [];
  for (const company of ['', '中文 & Co', 'Final']) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'extended-company-insert-v11' });
    try { session.document.setMetadata({ ...session.document.metadata, company }); const p = await session.prepareSave({ mode: 'edit' }); if (p.status !== 'ready') throw new Error(JSON.stringify(p.report)); bytes = (await session.save(p.plan)).bytes; rounds.push({ bytes: [...bytes], sources: session.report.coreMetadataSources, changes: p.plan.report.changes }); } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' }); try { return { rounds, company: reopened.document.metadata.company }; } finally { reopened.dispose(); }
}

async function exerciseExtendedCreation(api, input) {
  const session = await api.importPptx(new Uint8Array(input), { projection: 'content', metadataSaveProfile: 'extended-create-v12', customPropertySaveProfile: 'create-and-append-v5' });
  try {
    session.document.setMetadata({ ...session.document.metadata, title: 'Title', company: '中文 & Co', customProperties: [{ name: 'A', type: 'integer', value: 3 }] });
    const plan = await session.prepareSave({ mode: 'edit', customPropertyAdditions: [0] }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
    const bytes = (await session.save(plan.plan)).bytes, reopened = await api.importPptx(bytes, { projection: 'content' });
    try { return { bytes: [...bytes], metadata: reopened.document.metadata }; } finally { reopened.dispose(); }
  } finally { session.dispose(); }
}

async function exerciseMetadataRemoval(api, input) {
  let bytes = new Uint8Array(input); const rounds = [];
  for (let round=0;round<3;round++) {
    const session=await api.importPptx(bytes,{projection:'content',metadataSaveProfile:'metadata-remove-v13'});
    try { session.document.setMetadata({...session.document.metadata,subject:round===1?'New':undefined,company:round===1?'New company':undefined}); const plan=await session.prepareSave({mode:'edit'}); if(plan.status!=='ready')throw new Error(JSON.stringify(plan.report)); bytes=(await session.save(plan.plan)).bytes; rounds.push({bytes:[...bytes],sources:session.report.coreMetadataSources}); } finally { session.dispose(); }
  }
  const reopened=await api.importPptx(bytes,{projection:'content'});try{return {rounds,metadata:reopened.document.metadata};}finally{reopened.dispose();}
}

async function exerciseDateRemoval(api, input) {
  let bytes=new Uint8Array(input);const rounds=[];
  for(let round=0;round<3;round++){const s=await api.importPptx(bytes,{projection:'content',metadataSaveProfile:'metadata-date-remove-v14'});try{s.document.setMetadata({...s.document.metadata,createdAt:round===1?'2024-01-01T00:00:00.123Z':undefined});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw new Error(JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;rounds.push({bytes:[...bytes],sources:s.report.coreMetadataSources});}finally{s.dispose();}}
  const r=await api.importPptx(bytes,{projection:'content'});try{return {rounds,metadata:r.document.metadata};}finally{r.dispose();}
}

async function exerciseKeywordRemoval(api, input) {
  let bytes=new Uint8Array(input);const rounds=[];
  for(let round=0;round<3;round++){const s=await api.importPptx(bytes,{projection:'content',metadataSaveProfile:'metadata-keywords-remove-v15'});try{s.document.setMetadata({...s.document.metadata,keywords:round===1?['中文 & key']:[]});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw new Error(JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;rounds.push({bytes:[...bytes],sources:s.report.coreMetadataSources});}finally{s.dispose();}}
  const r=await api.importPptx(bytes,{projection:'content'});try{return {rounds,metadata:r.document.metadata};}finally{r.dispose();}
}

async function exerciseRevisionRemoval(api, input) {
  let bytes=new Uint8Array(input);const rounds=[];
  for(let round=0;round<3;round++){const s=await api.importPptx(bytes,{projection:'content',metadataSaveProfile:'metadata-revision-remove-v16'});try{s.document.setMetadata({...s.document.metadata,revision:round===1?1:undefined});const p=await s.prepareSave(round===1?{mode:'edit',coreMetadataInsertions:['revision']}:{mode:'edit',coreMetadataRemovals:['revision']});if(p.status!=='ready')throw new Error(JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;rounds.push({bytes:[...bytes],sources:s.report.coreMetadataSources});}finally{s.dispose();}}
  const r=await api.importPptx(bytes,{projection:'content'});try{return {rounds,metadata:r.document.metadata};}finally{r.dispose();}
}

async function exerciseDefaultRemoval(api, input) {
  let bytes=new Uint8Array(input);const rounds=[],fields=['title','author','language','revision'];
  for(let round=0;round<3;round++){const s=await api.importPptx(bytes,{projection:'content',metadataSaveProfile:'metadata-default-remove-v17'});try{s.document.setMetadata({...s.document.metadata,...(round===1?{title:'PPTKit Presentation',author:'PPTKit',language:'en-US',revision:1}:{title:undefined,author:undefined,language:undefined,revision:undefined})});const p=await s.prepareSave(round===1?{mode:'edit',coreMetadataInsertions:fields}:{mode:'edit',coreMetadataRemovals:fields});if(p.status!=='ready')throw new Error(JSON.stringify(p.report));bytes=(await s.save(p.plan)).bytes;rounds.push({bytes:[...bytes],sources:s.report.coreMetadataSources});}finally{s.dispose();}}
  const r=await api.importPptx(bytes,{projection:'content'});try{return {rounds,metadata:r.document.metadata};}finally{r.dispose();}
}

async function exerciseMixedDefaultAndCustom(api, input) {
  let bytes = new Uint8Array(input); const rounds = [], fields = ['title', 'author', 'language', 'revision'];
  const defaults = { title: 'PPTKit Presentation', author: 'PPTKit', language: 'en-US', revision: 1 };
  for (let round = 0; round < 3; round++) {
    const session = await api.importPptx(bytes, { projection: 'content', metadataSaveProfile: 'metadata-default-remove-v17', customPropertySaveProfile: 'create-and-append-v5' });
    try {
      session.document.setMetadata({ ...session.document.metadata, ...(round === 1 ? defaults : { title: undefined, author: undefined, language: undefined, revision: undefined }), customProperties: [{ name: 'Count', type: 'integer', value: 40 + round }] });
      const edit = round === 0 ? { coreMetadataRemovals: fields, customPropertyAdditions: [0] } : round === 1 ? { coreMetadataInsertions: fields } : { coreMetadataRemovals: fields };
      const plan = await session.prepareSave({ mode: 'edit', ...edit }); if (plan.status !== 'ready') throw new Error(JSON.stringify(plan.report));
      bytes = (await session.save(plan.plan)).bytes; rounds.push([...bytes]);
    } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' });
  try { return { rounds, metadata: reopened.document.metadata }; } finally { reopened.dispose(); }
}

async function exerciseIntegerVariants(api, input) {
  let bytes = new Uint8Array(input); const rounds = [];
  for (const [small, unsigned, single, basic, date] of [[127, 65535, 2.5, '中文 & _x0041_', '2021-02-03T04:05:06.000Z'], [-128, 0, 1.25, '', '2022-02-03T04:05:06.123Z'], [0, 1, -0, 'final', '2023-02-03T04:05:06.000Z']]) {
    const session = await api.importPptx(bytes, { projection: 'content', customPropertySaveProfile: 'scalar-v1' });
    try {
      const properties = session.document.metadata.customProperties.map(item => ({ ...item, value: item.name === 'Small' ? small : item.name === 'Unsigned' ? unsigned : item.name === 'Single' ? single : item.name === 'Decimal' ? (single === -0 ? 0 : single) : item.name === 'Date' ? date : basic }));
      session.document.setMetadata({ ...session.document.metadata, customProperties: properties });
      const prepared = await session.prepareSave({ mode: 'edit' });
      if (prepared.status !== 'ready') throw new Error('integer variant edit blocked: ' + JSON.stringify(prepared.report));
      bytes = (await session.save(prepared.plan)).bytes; rounds.push([...bytes]);
    } finally { session.dispose(); }
  }
  const reopened = await api.importPptx(bytes, { projection: 'content' });
  try { return { rounds, properties: reopened.document.metadata.customProperties, types: reopened.report.inspection.metadata.custom.properties.map(item => item.valueType) }; }
  finally { reopened.dispose(); }
}

test('custom property authoring and import agree across Node, browser ESM and global bundles', async t => {
  const expected = await exercise({ createPresentation, generatePptx, importPptx });
  const expectedCreation = await exerciseCreation({ createPresentation, generatePptx, importPptx });
  const expectedRevision = await exerciseRevision({ createPresentation, generatePptx, importPptx });
  const expectedKeywords = await exerciseKeywords({ createPresentation, generatePptx, importPptx });
  const expectedCompany = await exerciseCompany({ createPresentation, generatePptx, importPptx });
  const insertion = metadataFixture(); insertion.add('core', '<c:revision>7</c:revision>');
  const insertionInput = [...encodeMetadata(insertion)], expectedInsertion = await exerciseCoreInsertion({ importPptx }, insertionInput);
  const replace = metadataFixture(); replace.add('custom', custom(2, 'lpwstr', 'same', 'name="A"'));
  const replaceInput = [...encodeMetadata(replace)], expectedReplace = await exerciseSameValueReplacement({ importPptx }, replaceInput);
  const dates = metadataFixture(); dates.add('core', '<t:created xsi:type="t:W3CDTF">2020-02-29T10:00:00.1200+08:00</t:created><t:modified xsi:type="t:W3CDTF">2021-01-02T03:04:05Z</t:modified>');
  const datesInput = [...encodeMetadata(dates)], expectedDates = await exerciseCoreDates({ createPresentation, generatePptx, importPptx }, datesInput);
  const dateInsertion = metadataFixture(); dateInsertion.add('core', '<c:revision>7</c:revision>');
  const dateInsertionInput = [...encodeMetadata(dateInsertion)], expectedDateInsertion = await exerciseCoreDateInsertion({ importPptx }, dateInsertionInput);
  const coreCreationInput = [...encodeMetadata(metadataFixture())], expectedCoreCreation = await exerciseCoreCreation({ importPptx }, coreCreationInput);
  const expectedExtendedCreation = await exerciseExtendedCreation({ importPptx }, coreCreationInput);
  const removeFixture=metadataFixture();removeFixture.add('core','<d:subject>Old</d:subject>');removeFixture.add('extended','<Company>Old company</Company>');const removeInput=[...encodeMetadata(removeFixture)],expectedRemoval=await exerciseMetadataRemoval({importPptx},removeInput);
  const expectedDateRemoval=await exerciseDateRemoval({importPptx},datesInput);
  const keywordRemovalFixture=metadataFixture();keywordRemovalFixture.add('core','<c:keywords>one</c:keywords>');const keywordRemovalInput=[...encodeMetadata(keywordRemovalFixture)],expectedKeywordRemoval=await exerciseKeywordRemoval({importPptx},keywordRemovalInput);
  const revisionRemovalFixture=metadataFixture();revisionRemovalFixture.add('core','<c:revision>1</c:revision>');const revisionRemovalInput=[...encodeMetadata(revisionRemovalFixture)],expectedRevisionRemoval=await exerciseRevisionRemoval({importPptx},revisionRemovalInput);
  const defaultRemovalFixture=metadataFixture();defaultRemovalFixture.add('core','<d:title>Title</d:title><d:creator>Author</d:creator><d:language>fr-FR</d:language><c:revision>7</c:revision>');const defaultRemovalInput=[...encodeMetadata(defaultRemovalFixture)],expectedDefaultRemoval=await exerciseDefaultRemoval({importPptx},defaultRemovalInput);
  const expectedMixedDefaultAndCustom = await exerciseMixedDefaultAndCustom({ importPptx }, defaultRemovalInput);
  const companyInsertion = metadataFixture(); companyInsertion.add('extended', '<Manager>Sentinel</Manager>');
  const companyInsertionInput = [...encodeMetadata(companyInsertion)], expectedCompanyInsertion = await exerciseCompanyInsertion({ importPptx }, companyInsertionInput);
  const variants = metadataFixture(); variants.add('custom', custom(2, 'i1', '-128', 'name="Small"') + custom(3, 'ui2', '0', 'name="Unsigned"') + custom(4, 'r4', '1.25', 'name="Single"') + custom(5, 'bstr', 'original', 'name="Basic"') + custom(6, 'lpstr', 'original', 'name="Narrow"') + custom(7, 'date', '2020-02-03T04:05:06.120000+08:00', 'name="Date"') + custom(8, 'decimal', '1.25', 'name="Decimal"'));
  const variantInput = [...encodeMetadata(variants)], expectedVariants = await exerciseIntegerVariants({ importPptx }, variantInput);
  const collision = metadataFixture(); collision.add('custom', custom(2, 'lpwstr', 'first', 'name="_x005F_x0041_"') + custom(3, 'vector', '<v:lpstr>x</v:lpstr>', 'name="_x005f_x0041_"'));
  const collisionInput = [...encodeMetadata(collision)];
  const collisionSession = await importPptx(new Uint8Array(collisionInput), { projection: 'content' });
  const expectedCollision = { properties: collisionSession.document.metadata.customProperties, tiers: collisionSession.report.findings.filter(item => item.feature === 'metadata.custom-property').map(item => item.tier) };
  collisionSession.dispose();
  const currencyInput = [...await readFile('tests/fixtures/custom-currency-source.pptx')];
  const expectedOpaque = await inspectPptxOpaqueCustomProperties(new Uint8Array(currencyInput));
  const nodeCurrencySession = await importPptx(new Uint8Array(currencyInput), { projection: 'content', customCurrencyProjection: 'exact-v1' });
  const expectedCurrencyProjection = nodeCurrencySession.document.metadata.customProperties;
  nodeCurrencySession.dispose();
  const currencyTexts = ['1.2500', '.1000', '-12.5000', '9007199254740993.9999', '1.0000\u2028'];
  const expectedCurrency = currencyTexts.map(assessPptxCurrencyText);
  const expectedCurrencyNumbers = currencyTexts.map(assessPptxCurrencyNumber);
  const browser = await chromium.launch(); t.after(() => browser.close());
  const page = await browser.newPage(), requests = [];
  await page.route('**/*', route => { requests.push(route.request().url()); return route.abort(); });
  const names = ['core', 'layout', 'pptx-exporter', 'pptx-parser'], imports = {}, globals = [];
  for (const name of names) {
    imports[`@pptkit/${name}`] = 'data:text/javascript;base64,' + Buffer.from(await readFile(`packages/${name}/dist/browser/index.js`, 'utf8')).toString('base64');
    globals.push(await readFile(`packages/${name}/dist/browser/global.js`, 'utf8'));
  }
  await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
  for (const format of ['esm', 'global']) {
    if (format === 'global') for (const content of globals) await page.addScriptTag({ content });
    const actualCollision = await page.evaluate(async ({ format, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      const session = await api.importPptx(new Uint8Array(input), { projection: 'content' });
      try { return { properties: session.document.metadata.customProperties, tiers: session.report.findings.filter(item => item.feature === 'metadata.custom-property').map(item => item.tier) }; } finally { session.dispose(); }
    }, { format, input: collisionInput });
    assert.deepEqual(actualCollision, expectedCollision);
    const actualCurrencyProjection = await page.evaluate(async ({ format, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      const session = await api.importPptx(new Uint8Array(input), { projection: 'content', customCurrencyProjection: 'exact-v1' });
      try { return session.document.metadata.customProperties; } finally { session.dispose(); }
    }, { format, input: currencyInput });
    assert.deepEqual(actualCurrencyProjection, expectedCurrencyProjection);
    const actualOpaque = await page.evaluate(async ({ format, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return api.inspectPptxOpaqueCustomProperties(new Uint8Array(input));
    }, { format, input: currencyInput });
    assert.deepEqual(actualOpaque, expectedOpaque);
    const actualCurrency = await page.evaluate(async ({ format, texts }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return texts.map(api.assessPptxCurrencyText);
    }, { format, texts: currencyTexts });
    assert.deepEqual(actualCurrency, expectedCurrency);
    const actualCurrencyNumbers = await page.evaluate(async ({ format, texts }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return texts.map(api.assessPptxCurrencyNumber);
    }, { format, texts: currencyTexts });
    assert.deepEqual(actualCurrencyNumbers, expectedCurrencyNumbers);
    const actual = await page.evaluate(async ({ format, exercise }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'pptx-exporter', 'pptx-parser'].map(n => import('@pptkit/' + n))));
      return Function('api', `return (${exercise})(api)`)(api);
    }, { format, exercise: exercise.toString() });
    assert.deepEqual(actual, expected);
    const actualCreation = await page.evaluate(async ({ format, exercise }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'pptx-exporter', 'pptx-parser'].map(n => import('@pptkit/' + n))));
      return Function('api', `return (${exercise})(api)`)(api);
    }, { format, exercise: exerciseCreation.toString() });
    assert.deepEqual(actualCreation, expectedCreation);
    const actualRevision = await page.evaluate(async ({ format, exercise }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'pptx-exporter', 'pptx-parser'].map(n => import('@pptkit/' + n))));
      return Function('api', `return (${exercise})(api)`)(api);
    }, { format, exercise: exerciseRevision.toString() });
    assert.deepEqual(actualRevision, expectedRevision);
    const actualKeywords = await page.evaluate(async ({ format, exercise }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'pptx-exporter', 'pptx-parser'].map(n => import('@pptkit/' + n))));
      return Function('api', `return (${exercise})(api)`)(api);
    }, { format, exercise: exerciseKeywords.toString() });
    assert.deepEqual(actualKeywords, expectedKeywords);
    const actualCompany = await page.evaluate(async ({ format, exercise }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'pptx-exporter', 'pptx-parser'].map(n => import('@pptkit/' + n))));
      return Function('api', `return (${exercise})(api)`)(api);
    }, { format, exercise: exerciseCompany.toString() });
    assert.deepEqual(actualCompany, expectedCompany);
    const actualInsertion = await page.evaluate(async ({ format, exercise, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return Function('api', 'input', `return (${exercise})(api, input)`)(api, input);
    }, { format, exercise: exerciseCoreInsertion.toString(), input: insertionInput });
    assert.deepEqual(actualInsertion, expectedInsertion);
    const actualReplace = await page.evaluate(async ({ format, exercise, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return Function('api', 'input', `return (${exercise})(api, input)`)(api, input);
    }, { format, exercise: exerciseSameValueReplacement.toString(), input: replaceInput });
    assert.deepEqual(actualReplace, expectedReplace);
    const actualDates = await page.evaluate(async ({ format, exercise, input }) => {
      const api = format === 'global' ? PPTKit : Object.assign({}, ...await Promise.all(['core', 'pptx-exporter', 'pptx-parser'].map(n => import('@pptkit/' + n))));
      return Function('api', 'input', `return (${exercise})(api, input)`)(api, input);
    }, { format, exercise: exerciseCoreDates.toString(), input: datesInput });
    assert.deepEqual(actualDates, expectedDates);
    const actualDateInsertion = await page.evaluate(async ({ format, exercise, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return Function('api', 'input', `return (${exercise})(api, input)`)(api, input);
    }, { format, exercise: exerciseCoreDateInsertion.toString(), input: dateInsertionInput });
    assert.deepEqual(actualDateInsertion, expectedDateInsertion);
    const actualCoreCreation = await page.evaluate(async ({ format, exercise, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return Function('api', 'input', `return (${exercise})(api, input)`)(api, input);
    }, { format, exercise: exerciseCoreCreation.toString(), input: coreCreationInput });
    assert.deepEqual(actualCoreCreation, expectedCoreCreation);
    const actualExtendedCreation = await page.evaluate(async ({ format, exercise, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return Function('api', 'input', `return (${exercise})(api, input)`)(api, input);
    }, { format, exercise: exerciseExtendedCreation.toString(), input: coreCreationInput });
    assert.deepEqual(actualExtendedCreation, expectedExtendedCreation);
    const actualRemoval=await page.evaluate(async ({format,exercise,input})=>{const api=format==='global'?PPTKit:await import('@pptkit/pptx-parser');return Function('api','input',`return (${exercise})(api, input)`)(api,input);},{format,exercise:exerciseMetadataRemoval.toString(),input:removeInput});
    assert.deepEqual(actualRemoval,expectedRemoval);
    const actualDateRemoval=await page.evaluate(async ({format,exercise,input})=>{const api=format==='global'?PPTKit:await import('@pptkit/pptx-parser');return Function('api','input',`return (${exercise})(api,input)`)(api,input);},{format,exercise:exerciseDateRemoval.toString(),input:datesInput});
    assert.deepEqual(actualDateRemoval,expectedDateRemoval);
    const actualKeywordRemoval=await page.evaluate(async ({format,exercise,input})=>{const api=format==='global'?PPTKit:await import('@pptkit/pptx-parser');return Function('api','input',`return (${exercise})(api,input)`)(api,input);},{format,exercise:exerciseKeywordRemoval.toString(),input:keywordRemovalInput});
    assert.deepEqual(actualKeywordRemoval,expectedKeywordRemoval);
    const actualRevisionRemoval=await page.evaluate(async ({format,exercise,input})=>{const api=format==='global'?PPTKit:await import('@pptkit/pptx-parser');return Function('api','input',`return (${exercise})(api,input)`)(api,input);},{format,exercise:exerciseRevisionRemoval.toString(),input:revisionRemovalInput});
    assert.deepEqual(actualRevisionRemoval,expectedRevisionRemoval);
    const actualDefaultRemoval=await page.evaluate(async ({format,exercise,input})=>{const api=format==='global'?PPTKit:await import('@pptkit/pptx-parser');return Function('api','input',`return (${exercise})(api,input)`)(api,input);},{format,exercise:exerciseDefaultRemoval.toString(),input:defaultRemovalInput});
    const actualMixedDefaultAndCustom = await page.evaluate(async ({format, exercise, input}) => { const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser'); return Function('api', 'input', `return (${exercise})(api,input)`)(api, input); }, { format, exercise: exerciseMixedDefaultAndCustom.toString(), input: defaultRemovalInput });
    assert.deepEqual(actualMixedDefaultAndCustom, expectedMixedDefaultAndCustom);
    assert.deepEqual(actualDefaultRemoval,expectedDefaultRemoval);
    const actualCompanyInsertion = await page.evaluate(async ({ format, exercise, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return Function('api', 'input', `return (${exercise})(api, input)`)(api, input);
    }, { format, exercise: exerciseCompanyInsertion.toString(), input: companyInsertionInput });
    assert.deepEqual(actualCompanyInsertion, expectedCompanyInsertion);
    const actualVariants = await page.evaluate(async ({ format, exercise, input }) => {
      const api = format === 'global' ? PPTKit : await import('@pptkit/pptx-parser');
      return Function('api', 'input', `return (${exercise})(api, input)`)(api, input);
    }, { format, exercise: exerciseIntegerVariants.toString(), input: variantInput });
    assert.deepEqual(actualVariants, expectedVariants);
  }
  assert.deepEqual(expected.saved, expected.bytes);
  assert.deepEqual(requests, []);
});
