import { performance } from "node:perf_hooks";
import { createHash } from "node:crypto";
import { createPresentation, normalizePresentation } from "../../packages/core/dist/index.js";
import { resolveNormalizedLayout } from "../../packages/layout/dist/index.js";
import { renderLayoutToSvg } from "../../packages/svg-renderer/dist/index.js";
import { generatePptx } from "../../packages/pptx-exporter/dist/index.js";
import { createBenchmarkDeck, summarize } from "./workload.mjs";

const [slideCount, iterations, warmup] = process.argv.slice(2).map(Number);
if (![slideCount, iterations, warmup].every(Number.isSafeInteger) || slideCount < 1 || iterations < 2 || warmup < 0) throw new Error("Invalid worker arguments");
const document = createBenchmarkDeck(createPresentation, slideCount);
const samples = { normalizeMs: [], layoutMs: [], previewMs: [], exportPipelineMs: [], heapUsedBytes: [], rssBytes: [] };
let expectedHash;
let outputBytes;
let previewBytes;
for (let index = -warmup; index < iterations; index++) {
  globalThis.gc?.();
  const start = performance.now();
  const normalized = normalizePresentation(document);
  const normalizedAt = performance.now();
  const layout = resolveNormalizedLayout(normalized);
  const layoutAt = performance.now();
  const preview = await renderLayoutToSvg(layout);
  const previewAt = performance.now();
  const exported = await generatePptx(document);
  const exportedAt = performance.now();
  if (exported.warnings.length || preview.warnings.length) throw new Error("Benchmark fixture must not produce warnings");
  const hash = createHash("sha256").update(exported.bytes).digest("hex");
  if (expectedHash !== undefined && expectedHash !== hash) throw new Error("Non-deterministic export during benchmark");
  expectedHash = hash;
  outputBytes = exported.byteLength;
  previewBytes = preview.slides.reduce((sum, slide) => sum + Buffer.byteLength(slide.svg), 0);
  if (index >= 0) {
    samples.normalizeMs.push(normalizedAt - start);
    samples.layoutMs.push(layoutAt - normalizedAt);
    samples.previewMs.push(previewAt - layoutAt);
    samples.exportPipelineMs.push(exportedAt - previewAt);
    const memory = process.memoryUsage();
    samples.heapUsedBytes.push(memory.heapUsed);
    samples.rssBytes.push(memory.rss);
  }
}
console.log(JSON.stringify({ slideCount, iterations, warmup, outputBytes, previewBytes, sha256: expectedHash,
  peakProcessRssBytes: process.resourceUsage().maxRSS * 1024,
  metrics: Object.fromEntries(Object.entries(samples).map(([key, values]) => [key, summarize(values)])),
}));
