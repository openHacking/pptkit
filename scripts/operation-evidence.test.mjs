import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, symlink, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { renderOperationEvidence, validateOperationEvidence, verifyLocalOperationFixtures } from './operation-evidence.mjs';

const digest = bytes => createHash('sha256').update(bytes).digest('hex');
async function fixture(t) {
  const root = await mkdtemp(path.join(os.tmpdir(), 'pptkit-operation-'));
  t.after(() => rm(root, { recursive: true, force: true }));
  const source = Buffer.from('retained fixture bytes'), artifact = digest(source), bytes = Buffer.from(JSON.stringify({ fixture: artifact, result: { ok: true, count: 0 } }));
  await writeFile(path.join(root, 'source.pptx'), source);
  await writeFile(path.join(root, 'evidence.json'), bytes);
  const row = { id:'case', feature:'shape', subject:{object:'text',owner:'shape',family:'direct',variant:'run'}, operation:'edit', producer:{name:'producer',version:'1'}, application:{name:'app',version:'2'}, profile:'v1', fixture:{kind:'recorded-sha256',sha256:artifact,evidencePath:'evidence.json',pointer:'/fixture',localPath:'source.pptx'}, capability:'editable', status:'pass', method:'automated', procedure:'node example.mjs', result:'named case', limitations:'other cases untested', evidence:[{path:'evidence.json',sha256:digest(bytes),assertions:[{pointer:'/result/ok',equals:true},{pointer:'/result/count',equals:0}]}] };
  return {root, data:{schemaVersion:1,records:[row]}, row};
}
test('valid scoped evidence is deterministic and does not infer unrecorded combinations', async t => {
  const {root,data,row}=await fixture(t);await validateOperationEvidence(data,root);
  const second=structuredClone(row);second.id='other';second.application.version='3';second.status='untested';second.method='not-run';second.evidence=[];second.fixture.evidencePath='';second.fixture.pointer='';data.records.unshift(second);
  await validateOperationEvidence(data,root);const rendered=renderOperationEvidence(data);assert.match(rendered,/absent combination is untested/);assert.match(rendered,/\| untested \|/);const reversed=structuredClone(data);reversed.records.reverse();assert.equal(rendered,renderOperationEvidence(reversed));assert.ok(rendered.indexOf('`case`') < rendered.indexOf('`other`'));
});
test('invalid statuses, methods, schema, duplicate ids and duplicate combinations fail closed', async t => {
  const {root,data,row}=await fixture(t);
  const attempts=[d=>{d.schemaVersion=2},d=>{d.records[0].status='passed'},d=>{d.records[0].capability='supported'},d=>{d.records[0].status='untested'},d=>{d.records[0].method='not-run'},d=>{d.records[0].extra=true},d=>{delete d.records[0].subject.variant},d=>{const x=structuredClone(row);x.id='second';d.records.push(x)},d=>{const x=structuredClone(row);x.operation='different';d.records.push(x)}];
  for(const mutate of attempts){const copy=structuredClone(data);mutate(copy);await assert.rejects(validateOperationEvidence(copy,root));}
});
test('changed, absent, malformed, escaping or falsely asserted evidence cannot pass', async t => {
  const {root,data}=await fixture(t);
  const attempts=[d=>{d.records[0].evidence[0].sha256='0'.repeat(64)},d=>{d.records[0].fixture.sha256='b'.repeat(64)},d=>{d.records[0].fixture.localPath='../source.pptx'},d=>{d.records[0].evidence[0].assertions[0].equals=false},d=>{d.records[0].evidence[0].assertions[0].equals={}},d=>{d.records[0].evidence[0].assertions[0].pointer='/missing'},d=>{d.records[0].evidence[0].assertions[0].pointer='/bad~2'},d=>{d.records[0].evidence[0].assertions=[]},d=>{d.records[0].evidence[0].path='../evidence.json'},d=>{d.records[0].evidence[0].path='missing.json'}];
  for(const mutate of attempts){const copy=structuredClone(data);mutate(copy);await assert.rejects(validateOperationEvidence(copy,root));}
  const external=await mkdtemp(path.join(os.tmpdir(),'pptkit-outside-'));t.after(()=>rm(external,{recursive:true,force:true}));await writeFile(path.join(external,'outside.json'),await readFile(path.join(root,'evidence.json')));await symlink(path.join(external,'outside.json'),path.join(root,'link.json'));const copy=structuredClone(data);copy.records[0].evidence[0].path='link.json';copy.records[0].fixture.evidencePath='link.json';await assert.rejects(validateOperationEvidence(copy,root),/escapes repository/);
});
test('records bind recorded fixture hashes to a manifest without claiming bytes were rechecked', async t => {
  const {root,data}=await fixture(t);await validateOperationEvidence(data,root);assert.match(renderOperationEvidence(data),/fixture availability varies, and this table does not revalidate fixture bytes/);
  const copy=structuredClone(data);copy.records[0].fixture.pointer='/result/ok';await assert.rejects(validateOperationEvidence(copy,root),/fixture digest/);
});
test('manifest changes and incorrect JSON pointer escaping are detected', async t => {
  const {root,data}=await fixture(t);const bytes=Buffer.from(JSON.stringify({fixture:data.records[0].fixture.sha256,'a/b~c':false}));await writeFile(path.join(root,'evidence.json'),bytes);data.records[0].evidence[0].sha256=digest(bytes);data.records[0].evidence[0].assertions=[{pointer:'/a~1b~0c',equals:false}];await validateOperationEvidence(data,root);
  await writeFile(path.join(root,'evidence.json'),Buffer.from(bytes.toString()+' '));await assert.rejects(validateOperationEvidence(data,root),/evidence digest mismatch/);
});
test('markdown rendering keeps delimiters inside values and sorts by identity', async t => {
  const {root,data}=await fixture(t);data.records[0].feature='a | b';data.records[0].profile='has ` backtick';await validateOperationEvidence(data,root);const rendered=renderOperationEvidence(data);assert.match(rendered,/a \\\| b/);assert.match(rendered,/`` has ` backtick ``/);
});
test('optional local fixture verification rejects changed, missing and escaping source bytes', async t => {
  const {root,data}=await fixture(t);await validateOperationEvidence(data,root);assert.equal(await verifyLocalOperationFixtures(data,root),1);
  await writeFile(path.join(root,'source.pptx'),'changed');await assert.rejects(verifyLocalOperationFixtures(data,root),/fixture digest mismatch/);
  await rm(path.join(root,'source.pptx'));await assert.rejects(verifyLocalOperationFixtures(data,root),/ENOENT/);
  const outside=await mkdtemp(path.join(os.tmpdir(),'pptkit-source-outside-'));t.after(()=>rm(outside,{recursive:true,force:true}));await writeFile(path.join(outside,'source.pptx'),'retained fixture bytes');await symlink(path.join(outside,'source.pptx'),path.join(root,'source.pptx'));await assert.rejects(verifyLocalOperationFixtures(data,root),/escapes repository/);
});
test('checked-in catalog and generated table validate together', async () => {
  const root=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');const data=JSON.parse(await readFile(path.join(root,'docs/testing/operation-evidence.json'),'utf8'));await validateOperationEvidence(data,root);assert.equal(await readFile(path.join(root,'docs/testing/operation-evidence.md'),'utf8'),renderOperationEvidence(data));
});
test('connector observation manifests match retained source and output bytes', async () => {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const manifest = JSON.parse(await readFile(path.join(root, 'docs/testing/native-connector-interactions-evidence.json'), 'utf8'));
  for (const [key, source, output] of [
    ['straightMove', 'native-connector-move-before.pptx', 'native-connector-move-after.pptx'],
    ['elbowMove', 'native-elbow-move-before.pptx', 'native-elbow-move-after.pptx'],
    ['straightRotation', 'native-connector-move-before.pptx', 'native-connector-rotate-after.pptx'],
  ]) {
    assert.equal(digest(await readFile(path.join(root, 'tests/fixtures', source))), manifest[key].sourceSha256);
    assert.equal(digest(await readFile(path.join(root, 'tests/fixtures', output))), manifest[key].outputSha256);
  }
});

test('mixed text attribution record binds its retained fixture and registered regression', async () => {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const data = JSON.parse(await readFile(path.join(root, 'docs/testing/operation-evidence.json'), 'utf8'));
  const row = data.records.find(item => item.id === 'text-v1-mixed-three-round-attribution');
  assert.ok(row); assert.equal(row.method, 'automated'); assert.equal(row.status, 'pass');
  assert.equal(digest(await readFile(path.join(root, row.fixture.localPath))), row.fixture.sha256);
  const manifest = JSON.parse(await readFile(path.join(root, row.fixture.evidencePath), 'utf8'));
  assert.equal(manifest.sourceSha256, row.fixture.sha256); assert.equal(manifest.procedure, row.procedure);
  assert.equal(manifest.nativeApplication, 'untested');
  const scripts = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8')).scripts;
  assert.match(scripts.test, /pnpm test:import(?:\s|$)/);
  assert.match(scripts['test:import'], /(?:^|\s)scripts\/text-save-parity\.test\.mjs(?:\s|$)/);
});

test('native notes and later PPTKit records use distinct retained baselines', async () => {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const data = JSON.parse(await readFile(path.join(root, 'docs/testing/operation-evidence.json'), 'utf8'));
  const native = data.records.find(row => row.id === 'notes-native-powerpoint-corpus-edit');
  const pptkit = data.records.find(row => row.id === 'notes-v2-after-powerpoint-three-round');
  assert.ok(native && pptkit); assert.equal(native.method, 'manual'); assert.equal(pptkit.method, 'automated');
  for (const row of [native, pptkit]) assert.equal(digest(await readFile(path.join(root, row.fixture.localPath))), row.fixture.sha256);
  const nativeEvidence = JSON.parse(await readFile(path.join(root, native.fixture.evidencePath), 'utf8'));
  const attribution = JSON.parse(await readFile(path.join(root, pptkit.fixture.evidencePath), 'utf8'));
  assert.equal(nativeEvidence.output.sha256, pptkit.fixture.sha256);
  assert.equal(attribution.nativeSourceSha256, native.fixture.sha256); assert.equal(attribution.procedure, pptkit.procedure);
  const scripts = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8')).scripts;
  assert.match(scripts['test:import'], /(?:^|\s)scripts\/text-owner-save-parity\.test\.mjs(?:\s|$)/);
});

test('caption loss and later no-op use separately verified producer baselines', async () => {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const data = JSON.parse(await readFile(path.join(root, 'docs/testing/operation-evidence.json'), 'utf8'));
  const native = data.records.find(row => row.id === 'captions-libreoffice-headless-loss');
  const pptkit = data.records.find(row => row.id === 'captions-after-libreoffice-pptkit-noop');
  assert.ok(native && pptkit);
  assert.deepEqual([native.capability, native.status, pptkit.capability, pptkit.status], ['unsupported', 'fail', 'preserved', 'pass']);
  for (const row of [native, pptkit]) assert.equal(digest(await readFile(path.join(root, row.fixture.localPath))), row.fixture.sha256);
  const evidence = JSON.parse(await readFile(path.join(root, native.fixture.evidencePath), 'utf8'));
  assert.equal(native.fixture.sha256, evidence.source_sha256);
  assert.equal(pptkit.fixture.sha256, evidence.retained_first_round_sha256);
  assert.notEqual(native.fixture.sha256, pptkit.fixture.sha256);
  assert.equal(pptkit.fixture.localPath, evidence.retained_first_round_path);
  const parser = JSON.parse(await readFile(path.join(root, 'packages/pptx-parser/package.json'), 'utf8'));
  assert.match(parser.scripts.test, /node --test test\/\*\.test\.mjs/);
});

test('mixed SmartArt and timing evidence is bound to distributed fixtures and regression', async () => {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const data = JSON.parse(await readFile(path.join(root, 'docs/testing/operation-evidence.json'), 'utf8'));
  const row = data.records.find(item => item.id === 'text-v1-smartart-timing-three-round'); assert.ok(row);
  const evidence = JSON.parse(await readFile(path.join(root, row.fixture.evidencePath), 'utf8'));
  assert.equal(row.method, 'automated'); assert.equal(row.status, 'pass'); assert.equal(evidence.nativeApplication, 'untested');
  assert.equal(digest(await readFile(path.join(root, row.fixture.localPath))), row.fixture.sha256);
  assert.equal(digest(await readFile(path.join(root, evidence.basePath))), evidence.baseSha256);
  assert.equal(evidence.procedure, row.procedure);
  const scripts = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8')).scripts;
  assert.match(scripts['test:import'], /(?:^|\s)scripts\/text-save-parity\.test\.mjs(?:\s|$)/);
});

test('mixed media and comments evidence is bound to distributed fixtures and regression', async () => {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const data = JSON.parse(await readFile(path.join(root, 'docs/testing/operation-evidence.json'), 'utf8'));
  const row = data.records.find(item => item.id === 'text-v1-media-comments-three-round'); assert.ok(row);
  const evidence = JSON.parse(await readFile(path.join(root, row.fixture.evidencePath), 'utf8'));
  assert.equal(row.method, 'automated'); assert.equal(row.status, 'pass'); assert.equal(evidence.nativeApplication, 'untested');
  assert.equal(digest(await readFile(path.join(root, row.fixture.localPath))), row.fixture.sha256);
  assert.equal(digest(await readFile(path.join(root, evidence.basePath))), evidence.baseSha256);
  assert.equal(evidence.procedure, row.procedure);
  const scripts = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8')).scripts;
  assert.match(scripts['test:import'], /(?:^|\s)scripts\/text-save-parity\.test\.mjs(?:\s|$)/);
});

test('mixed picture evidence binds the distributed source and registered regression', async () => {
  const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
  const data = JSON.parse(await readFile(path.join(root, 'docs/testing/operation-evidence.json'), 'utf8'));
  const row = data.records.find(item => item.id === 'picture-v1-advanced-mixed-three-round'); assert.ok(row);
  const evidence = JSON.parse(await readFile(path.join(root, row.fixture.evidencePath), 'utf8'));
  assert.equal(row.method, 'automated'); assert.equal(row.status, 'pass'); assert.equal(evidence.nativeApplication, 'untested');
  assert.equal(digest(await readFile(path.join(root, row.fixture.localPath))), row.fixture.sha256);
  assert.equal(evidence.procedure, row.procedure);
  const scripts = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8')).scripts;
  assert.match(scripts['test:import'], /(?:^|\s)scripts\/picture-save-parity\.test\.mjs(?:\s|$)/);
});
