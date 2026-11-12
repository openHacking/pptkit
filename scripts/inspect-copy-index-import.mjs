// Optional inspection of the actual captured producer files. No native editing occurs here.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import { importPptx } from '../packages/pptx-parser/dist/index.js';

const lo = JSON.parse(await readFile('docs/testing/copy-placeholder-index-libreoffice-evidence.json', 'utf8'));
const powerpoint = JSON.parse(await readFile('docs/testing/copy-placeholder-index-powerpoint-evidence.json', 'utf8'));
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const producers = [
  { name: 'LibreOfficeDev', files: lo.fileHashes },
  { name: 'PowerPoint', files: powerpoint.independent.rounds },
];
const results = [];
for (const producer of producers) {
  for (const [round, entry] of producer.files.entries()) {
    const bytes = await readFile(entry.file); assert.equal(hash(bytes), entry.sha256, 'Captured producer file changed');
    const session = await importPptx(bytes, { projection: 'content' });
    try {
      const plan = await session.prepareSave({ mode: 'no-op' }); assert.equal(plan.status, 'ready');
      assert.deepEqual((await session.save(plan.plan)).bytes, new Uint8Array(bytes));
      const declarations = session.report.placeholderInheritance.placeholders.filter(item => item.level === 'slide').map(item => ({
        part: item.source.part, type: item.type, index: item.index, explicitType: item.explicitType, explicitIndex: item.explicitIndex,
        parent: { rule: item.parent.rule, status: item.parent.status, candidates: item.parent.candidateCount },
        geometryStatus: item.geometryStatus,
      }));
      assert.equal(declarations.length, 4);
      if (round === 0) assert.deepEqual(declarations.map(x => [x.index, x.explicitIndex, x.parent.status]), [[0, true, 'matched'], [4294967295, true, 'matched'], [0, true, 'matched'], [4294967295, true, 'matched']]);
      else if (producer.name === 'PowerPoint') assert.deepEqual(declarations.map(x => [x.index, x.explicitIndex, x.parent.status]), [[0, false, 'matched'], [4294967295, true, 'matched'], [0, false, 'matched'], [4294967295, true, 'matched']]);
      else assert.ok(declarations.every(x => x.index === 0 && !x.explicitIndex && x.parent.status === 'ambiguous' && x.parent.candidates === 2));
      const findings = session.report.findings.filter(f => f.feature === 'drawing.placeholder-match').map(f => ({ tier: f.tier, part: f.source.part, reason: f.reason }));
      assert.equal(findings.length, 4);
      assert.ok(findings.every(f => f.tier === (producer.name === 'LibreOfficeDev' && round > 0 ? 'fallback' : 'editable')));
      results.push({ producer: producer.name, round, file: entry.file, sha256: entry.sha256, noOpExact: true, declarations, findings, diagnostics: session.report.diagnostics.map(x => x.code) });
    } finally { session.dispose(); }
  }
}
await writeFile('docs/testing/copy-placeholder-index-import-evidence.json', JSON.stringify({ schemaVersion: 1, scope: 'Parser content projection and no-op save of the two-page actual producer fixtures; findings describe the named files only, not arbitrary native placeholders or edit-aware saves.', inputs: { libreoffice: 'copy-placeholder-index-libreoffice-evidence.json', powerpoint: 'copy-placeholder-index-powerpoint-evidence.json' }, results }, null, 2) + '\n');
console.log(JSON.stringify(results.map(({ producer, round, noOpExact, declarations, findings }) => ({ producer, round, noOpExact, statuses: declarations.map(x => x.parent.status), tiers: findings.map(x => x.tier) })), null, 2));
