import { readFile, writeFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { renderOperationEvidence, validateOperationEvidence, verifyLocalOperationFixtures } from './operation-evidence.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const catalog = path.join(root, 'docs/testing/operation-evidence.json');
const output = path.join(root, 'docs/testing/operation-evidence.md');
const args = process.argv.slice(2);
if (args.length > 1 || (args.length === 1 && !['--write', '--verify-local-fixtures'].includes(args[0]))) throw new Error('Usage: node scripts/operation-evidence-cli.mjs [--write|--verify-local-fixtures]');
const data = JSON.parse(await readFile(catalog, 'utf8'));
await validateOperationEvidence(data, root);
const rendered = renderOperationEvidence(data);
if (args[0] === '--write') await writeFile(output, rendered);
else if (await readFile(output, 'utf8') !== rendered) throw new Error('Generated operation evidence table is stale; run node scripts/operation-evidence-cli.mjs --write');
if (args[0] === '--verify-local-fixtures') console.log(`Verified ${await verifyLocalOperationFixtures(data, root)} retained fixture files`);
console.log(`Validated ${data.records.length} scoped operation evidence records`);
