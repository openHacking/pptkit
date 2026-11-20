import { readFile, writeFile } from 'node:fs/promises';
import { importPptx } from '../../packages/pptx-parser/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

if (!process.argv[2]) throw new Error('Usage: node scripts/research/generate-mp4-export.mjs <output.pptx>');
const bytes = await readFile(new URL('../../tests/fixtures/h264-mp4-powerpoint.pptx', import.meta.url));
const session = await importPptx(bytes, { projection: 'content', videoProjection: 'iso-bmff-v1' });
try { await writeFile(process.argv[2], (await generatePptx(session.document, { videoMode: 'iso-bmff-v1' })).bytes); }
finally { session.dispose(); }
