import { readFile, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';

// Generate the pre-PowerPoint base; MP4 must be provided explicitly.
const [input, output] = process.argv.slice(2);
if (!input || !output) throw new TypeError('Usage: node scripts/research/generate-caption-playback.mjs input.mp4 output.pptx');
const p = createPresentation();
p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: await readFile(input) } });
p.addSlide({ elements: [{ type: 'video', videoAssetId: 'movie', playbackStart: 'object-click', trimStartMs: 2000, trimEndMs: 2000, box: { x: 100, y: 100, width: 600, height: 338 } }] });
await writeFile(output, (await generatePptx(p, { videoMode: 'iso-bmff-v1' })).bytes);
