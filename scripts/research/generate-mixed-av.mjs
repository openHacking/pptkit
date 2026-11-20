import { readFile, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../../tests/support/read-generated-zip.mjs';

if (!process.argv[2]) throw new Error('Usage: node scripts/research/generate-mixed-av.mjs <output.pptx>');
const video = readGeneratedZip(await readFile(new URL('../../tests/fixtures/h264-mp4-powerpoint.pptx', import.meta.url)));
const audio = readGeneratedZip(await readFile(new URL('../../tests/fixtures/pcm-wave-powerpoint.pptx', import.meta.url)));
const p = createPresentation();
p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: video.get('ppt/media/media1.mp4') } });
p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/x-wav', source: { type: 'bytes', value: audio.get('ppt/media/media1.wav') } });
p.registerAsset({ id: 'poster', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: video.get('ppt/media/image1.png') } });
p.addSlide({ elements: [
  ...[0, 1].map(i => ({ id: `video-${i}`, type: 'video', videoAssetId: 'movie', posterAssetId: 'poster', box: { x: 50 + i * 300, y: 80, width: 240, height: 135 } })),
  ...[0, 1].map(i => ({ id: `audio-${i}`, type: 'audio', audioAssetId: 'sound', box: { x: 100 + i * 200, y: 270, width: 60, height: 60 } })),
] });
await writeFile(process.argv[2], (await generatePptx(p, { videoMode: 'iso-bmff-v1', audioMode: 'wave-pcm-v1' })).bytes);
