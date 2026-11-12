import { createHash } from 'node:crypto';
import { readFile, realpath } from 'node:fs/promises';
import path from 'node:path';

const states = new Set(['planned', 'untested', 'pass', 'fail', 'partial']);
const capabilities = new Set(['editable', 'preserved', 'fallback', 'unsupported', 'unknown']);
const methods = new Set(['automated', 'manual', 'not-run']);
const sha256 = /^[a-f0-9]{64}$/;
const isRecord = value => value !== null && typeof value === 'object' && !Array.isArray(value);
const text = value => typeof value === 'string' && value.trim().length > 0;
const fail = message => { throw new Error(message); };
const keys = (value, allowed, location) => {
  if (!isRecord(value)) fail(`${location}: expected object`);
  for (const key of Object.keys(value)) if (!allowed.includes(key)) fail(`${location}: unknown property ${key}`);
  for (const key of allowed) if (!Object.hasOwn(value, key)) fail(`${location}: missing ${key}`);
};
const string = (value, location) => { if (!text(value)) fail(`${location}: expected nonempty string`); };
const markdown = value => value.replaceAll('\\', '\\\\').replaceAll('|', '\\|').replaceAll('\n', ' ');
const code = value => {
  const content = value.replaceAll('\n', ' ');
  const longest = Math.max(0, ...[...content.matchAll(/`+/g)].map(match => match[0].length));
  const fence = '`'.repeat(longest + 1);
  return `${fence}${longest ? ' ' : ''}${content}${longest ? ' ' : ''}${fence}`;
};

function validatePath(name) {
  if (!text(name) || !/^[A-Za-z0-9._/-]+$/.test(name) || path.posix.isAbsolute(name) || name.split('/').some(x => x === '' || x === '.' || x === '..')) fail(`invalid repository path: ${name}`);
}
async function localFile(root, name) {
  validatePath(name);
  const resolvedRoot = await realpath(root);
  const resolved = await realpath(path.resolve(resolvedRoot, name));
  if (!resolved.startsWith(resolvedRoot + path.sep)) fail(`path escapes repository: ${name}`);
  return readFile(resolved);
}
function pointer(object, ptr) {
  if (typeof ptr !== 'string' || !ptr.startsWith('/') || /~(?![01])/.test(ptr)) fail(`invalid JSON pointer: ${ptr}`);
  let value = object;
  for (const raw of ptr.slice(1).split('/')) {
    const key = raw.replaceAll('~1', '/').replaceAll('~0', '~');
    if ((!isRecord(value) && !Array.isArray(value)) || !Object.hasOwn(value, key)) fail(`missing evidence pointer: ${ptr}`);
    value = value[key];
  }
  return value;
}
const equal = (a, b) => Object.is(a, b);

/** Check explicit evidence claims without running applications or interpreting unlisted combinations. */
export async function validateOperationEvidence(data, root) {
  keys(data, ['schemaVersion', 'records'], 'catalog');
  if (data.schemaVersion !== 1) fail('catalog: unsupported schemaVersion');
  if (!Array.isArray(data.records) || data.records.length === 0) fail('catalog: records must be nonempty');
  const seen = new Set(), ids = new Set(), cache = new Map();
  for (const [i, r] of data.records.entries()) {
    const at = `records[${i}]`;
    keys(r, ['id', 'feature', 'subject', 'operation', 'producer', 'application', 'profile', 'fixture', 'capability', 'status', 'method', 'procedure', 'result', 'limitations', 'evidence'], at);
    for (const key of ['id', 'feature', 'operation', 'profile', 'procedure', 'result', 'limitations']) string(r[key], `${at}.${key}`);
    keys(r.subject, ['object', 'owner', 'family', 'variant'], `${at}.subject`);
    for (const key of ['object', 'owner', 'family', 'variant']) string(r.subject[key], `${at}.subject.${key}`);
    for (const key of ['producer', 'application']) { keys(r[key], ['name', 'version'], `${at}.${key}`); string(r[key].name, `${at}.${key}.name`); string(r[key].version, `${at}.${key}.version`); }
    if (!capabilities.has(r.capability)) fail(`${at}: unknown capability`);
    if (!states.has(r.status)) fail(`${at}: unknown status`);
    if (!methods.has(r.method)) fail(`${at}: unknown method`);
    if (!Array.isArray(r.evidence)) fail(`${at}: evidence must be an array`);
    const pending = r.status === 'untested' || r.status === 'planned';
    if (pending && (r.method !== 'not-run' || r.evidence.length !== 0)) fail(`${at}: pending records cannot claim executed evidence`);
    if (!pending && (r.method === 'not-run' || r.evidence.length === 0)) fail(`${at}: executed records require evidence and a method`);
    keys(r.fixture, ['kind', 'sha256', 'evidencePath', 'pointer', 'localPath'], `${at}.fixture`);
    if (r.fixture.kind !== 'recorded-sha256' || !sha256.test(r.fixture.sha256)) fail(`${at}: invalid fixture hash or kind`);
    validatePath(r.fixture.localPath);
    const identity = JSON.stringify([r.subject.object, r.subject.owner, r.subject.family, r.subject.variant, r.operation, r.producer.name, r.producer.version, r.application.name, r.application.version, r.profile, r.fixture.sha256]);
    if (seen.has(identity)) fail(`${at}: duplicate combination`); seen.add(identity);
    if (ids.has(r.id)) fail(`${at}: duplicate id`); ids.add(r.id);
    for (const e of r.evidence) {
      keys(e, ['path', 'sha256', 'assertions'], `${at}.evidence`);
      if (!sha256.test(e.sha256) || !Array.isArray(e.assertions) || !e.assertions.length) fail(`${at}: invalid evidence digest or empty assertions`);
      let item = cache.get(e.path);
      if (!item) { const bytes = await localFile(root, e.path); item = { digest: createHash('sha256').update(bytes).digest('hex'), json: JSON.parse(bytes.toString('utf8')) }; cache.set(e.path, item); }
      if (item.digest !== e.sha256) fail(`${at}: evidence digest mismatch: ${e.path}`);
      for (const assertion of e.assertions) {
        keys(assertion, ['pointer', 'equals'], `${at}.assertion`);
        if (assertion.equals !== null && !['boolean', 'number', 'string'].includes(typeof assertion.equals)) fail(`${at}: assertion value must be scalar`);
        if (!equal(pointer(item.json, assertion.pointer), assertion.equals)) fail(`${at}: evidence assertion mismatch: ${assertion.pointer}`);
      }
    }
    if (!pending) {
      const source = r.evidence.find(e => e.path === r.fixture.evidencePath);
      if (!source) fail(`${at}: fixture source must be linked evidence`);
      if (pointer(cache.get(source.path).json, r.fixture.pointer) !== r.fixture.sha256) fail(`${at}: fixture digest does not match evidence`);
    } else if (r.fixture.evidencePath !== '' || r.fixture.pointer !== '') fail(`${at}: pending fixture must not claim evidence`);
  }
  return data;
}

/** Explicitly opt in to reading retained local fixture bytes; absence is an error. */
export async function verifyLocalOperationFixtures(data, root) {
  const checked = new Map();
  for (const r of data.records) {
    const { localPath, sha256: expected } = r.fixture;
    let actual = checked.get(localPath);
    if (!actual) { actual = createHash('sha256').update(await localFile(root, localPath)).digest('hex'); checked.set(localPath, actual); }
    if (actual !== expected) fail(`fixture digest mismatch: ${localPath}`);
  }
  return checked.size;
}

export function renderOperationEvidence(data) {
  const lines = ['# Scoped operation evidence', '', '<!-- Generated by scripts/operation-evidence-cli.mjs; edit the catalog instead. -->', '', 'Each row is a single observed combination, not a general compatibility claim. Fixture hashes are recorded in the linked evidence manifests; fixture availability varies, and this table does not revalidate fixture bytes. An absent combination is untested. Capability and verification status are separate.', '', '| Feature | Operation | Producer | Application | Profile | Capability | Status | Evidence | Result and limits |', '| --- | --- | --- | --- | --- | --- | --- | --- | --- |'];
  const sorted = [...data.records].sort((a,b) => a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
  for (const r of sorted) {
    const links = r.evidence.map((e,i) => `[${i+1}](../../${e.path})`).join(', ') || '—';
    lines.push(`| ${markdown(r.feature)} | ${markdown(r.operation)} | ${markdown(r.producer.name)} ${markdown(r.producer.version)} | ${markdown(r.application.name)} ${markdown(r.application.version)} | ${code(r.profile)} | ${r.capability} | ${r.status} | ${links} | ${markdown(r.result)} ${markdown(r.limitations)} |`);
  }
  lines.push('', '## Record details', '');
  for (const r of sorted) {
    lines.push(`- ${code(r.id)} — subject: ${[r.subject.object, r.subject.owner, r.subject.family, r.subject.variant].map(code).join(' / ')}; fixture SHA-256 ${code(r.fixture.sha256)}; retained path: ${code(r.fixture.localPath)}; method: ${r.method}; procedure: ${code(r.procedure)}.`);
  }
  return lines.join('\n') + '\n';
}
