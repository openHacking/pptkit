import { readFile, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../../tests/support/read-generated-zip.mjs';

const output = process.argv[2];
if (!output || process.argv.length !== 3) throw new TypeError('Provide an output PPTX path.');
const input = await readFile(new URL('../../tests/fixtures/pcm-wave-powerpoint.pptx', import.meta.url));
const wave = readGeneratedZip(input).get('ppt/media/media1.wav');
const p = createPresentation();
p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: wave } });
p.addSlide({ elements: [{ type: 'audio', id: 'speaker', audioAssetId: 'sound', box: { x: 200, y: 100, width: 64, height: 64 } }] });
await writeFile(output, (await generatePptx(p, { audioMode: 'wave-pcm-v1' })).bytes);
