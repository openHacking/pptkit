import { readFile, writeFile } from "node:fs/promises";
import { createPresentation } from "../../packages/core/dist/index.js";
import { generatePptx } from "../../packages/pptx-exporter/dist/index.js";
import { readGeneratedZip } from "../../tests/support/read-generated-zip.mjs";

if (!process.argv[2]) throw new Error("Usage: node scripts/research/generate-audio-volume.mjs <output.pptx>");
const source = readGeneratedZip(await readFile(new URL("../../tests/fixtures/pptkit-mixed-av.pptx", import.meta.url)));
const p = createPresentation();
p.registerAsset({ id: "sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: source.get("ppt/media/sound.wav") } });
p.addSlide({ elements: ["low", "medium", "high", "mute"].map((playbackVolume, i) => ({ id: `audio-${playbackVolume}`, type: "audio", audioAssetId: "sound", playbackVolume, playbackStart: i % 2 ? "object-click" : "click-sequence", box: { x: 80 + i * 150, y: 140, width: 64, height: 64 } })) });
await writeFile(process.argv[2], (await generatePptx(p, { audioMode: "wave-pcm-v1" })).bytes);
