# Performance and Output Baselines

Run `pnpm benchmark` to build the browser-neutral packages and measure deterministic mixed decks at 10, 50 and 200 slides. Each size runs in a separate Node process with explicit garbage collection between iterations, two warmups, and five measured samples. The default output is `.tmp/benchmarks/baseline.json`.

The [recorded baseline](baseline.json) includes raw samples, mean/median/p95, standard deviation and coefficient of variation; output bytes/hash; sampled heap/RSS and process lifetime RSS peak; runtime/OS/CPU/package metadata; and source revision/dirty state. A dirty baseline records local development evidence and must not be described as measurements of the clean recorded commit.

The [workload](../../scripts/benchmarks/workload.mjs) includes rich text, nested content, notes, a shared repository-authored PNG, and alternating editable tables/charts. No external asset network access is needed. Font family is Arial; current normalization uses deterministic heuristics and preview measures SVG serialization, not browser paint or font loading. Future font-backed layout and raster/PDF adapters need additional workloads.

## What is measured

- `normalizeMs`: authoring document to Canonical IR.
- `layoutMs`: already-normalized IR to detached Layout.
- `previewMs`: resolved Layout to SVG strings, including asset resolution.
- `exportPipelineMs`: public PPTX export including its own normalization, Layout, packaging and ZIP encoding. Do not sum this with normalization/Layout and call the result one export.
- `heapUsedBytes` and `rssBytes`: memory sampled after each measured iteration; they are not stage peaks.
- `peakProcessRssBytes`: OS-reported high-water RSS of the isolated worker including warmups and workload creation. It is not retained memory or a leak measurement.

Use the same runtime, machine, workload version and build settings for comparisons. The current baseline is descriptive; STB-03 will establish budgets after repeated measurements on the supported CI environments. A single noisy p95 is not a release threshold, and preview serialization timings do not establish interactive frame rate.

For a longer run after building:

```sh
node scripts/benchmark.mjs --iterations 20 --warmup 5 --output .tmp/benchmarks/extended.json
```

The runner rejects invalid iteration counts, fails on fixture warnings, and verifies byte-identical PPTX across iterations. [Focused tests](../../scripts/benchmark.test.mjs) check statistical behavior, workload validity, repeatability and input detachment. Baseline artifacts must be deliberately reviewed before replacing the checked-in record.
