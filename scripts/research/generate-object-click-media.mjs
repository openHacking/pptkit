import { readFile, writeFile } from 'node:fs/promises';
import { createPresentation } from '../../packages/core/dist/index.js';
import { generatePptx } from '../../packages/pptx-exporter/dist/index.js';
import { readGeneratedZip } from '../../tests/support/read-generated-zip.mjs';

if (!process.argv[2]) throw new Error('Usage: node scripts/research/generate-object-click-media.mjs <output.pptx>');
const source = readGeneratedZip(await readFile(new URL('../../tests/fixtures/pptkit-mixed-av.pptx', import.meta.url)));
const p = createPresentation();
if (process.argv[3] === '--single') {
  p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: source.get('ppt/media/sound.wav') } });
  p.addSlide({ elements: [{ type: 'audio', audioAssetId: 'sound', playbackStart: 'object-click', box: { x: 200, y: 100, width: 64, height: 64 } }] });
  await writeFile(process.argv[2], (await generatePptx(p, { audioMode: 'wave-pcm-v1' })).bytes);
  process.exit(0);
}
p.registerAsset({ id: 'movie', kind: 'video', mimeType: 'video/mp4', source: { type: 'bytes', value: source.get('ppt/media/movie.mp4') } });
p.registerAsset({ id: 'sound', kind: 'audio', mimeType: 'audio/wav', source: { type: 'bytes', value: source.get('ppt/media/sound.wav') } });
p.registerAsset({ id: 'poster', kind: 'image', mimeType: 'image/png', source: { type: 'bytes', value: source.get('ppt/media/poster.png') } });
p.addSlide({ elements: [
  ...[0, 1].map(i => ({ id: `video-${i}`, type: 'video', videoAssetId: 'movie', posterAssetId: 'poster', box: { x: 50 + i * 300, y: 80, width: 240, height: 135 } })),
  ...[0, 1].map(i => ({ id: `audio-${i}`, type: 'audio', audioAssetId: 'sound', ...(i === 1 ? { playbackStart: 'object-click' } : {}), box: { x: 100 + i * 200, y: 270, width: 60, height: 60 } })),
] });
await writeFile(process.argv[2], (await generatePptx(p, { videoMode: 'iso-bmff-v1', audioMode: 'wave-pcm-v1' })).bytes);
