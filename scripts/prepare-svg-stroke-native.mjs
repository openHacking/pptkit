import { mkdir, writeFile } from 'node:fs/promises';
import { importSimpleSvg } from '../packages/svg-parser/dist/index.js';
import { createPresentation } from '../packages/core/dist/index.js';
import { generatePptx } from '../packages/pptx-exporter/dist/index.js';

const source = '<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" fill="none" stroke="#f00" stroke-width="16"><circle cx="180" cy="200" r="100"/><g stroke="#00f" stroke-width="12" transform="translate(280,0)"><ellipse cx="180" cy="200" rx="120" ry="80"/></g><ellipse cx="760" cy="200" rx="100" ry="140" stroke="#0a0" stroke-width="8"/></svg>';
const result = importSimpleSvg(source);
if (result.status !== 'editable') throw new Error(JSON.stringify(result.issues));
const deck = createPresentation({ size: { width: result.width, height: result.height } });
deck.addSlide({ elements: [...result.elements] });
await mkdir('.tmp/native-svg-stroke', { recursive: true });
await writeFile('.tmp/native-svg-stroke/source.svg', source, { flag: 'wx' });
await writeFile('.tmp/native-svg-stroke/source.pptx', (await generatePptx(deck)).bytes, { flag: 'wx' });
