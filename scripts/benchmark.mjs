import { spawnSync, execFileSync } from "node:child_process";
import { readFileSync, mkdirSync, writeFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { workloads } from "./benchmarks/workload.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = process.argv.slice(2);
const options = { output: ".tmp/benchmarks/baseline.json", iterations: 5, warmup: 2 };
for (let index = 0; index < args.length; index += 2) {
  const key = args[index]?.replace(/^--/, "");
  const value = args[index + 1];
  if (!["output", "iterations", "warmup"].includes(key) || value === undefined) throw new Error("Usage: benchmark.mjs [--output file] [--iterations n] [--warmup n]");
  options[key] = key === "output" ? value : Number(value);
}
if (!Number.isSafeInteger(options.iterations) || options.iterations < 2 || !Number.isSafeInteger(options.warmup) || options.warmup < 0) throw new Error("iterations must be >= 2; warmup must be >= 0");
const revision = execFileSync("git", ["rev-parse", "HEAD"], { cwd: root, encoding: "utf8" }).trim();
const dirty = execFileSync("git", ["status", "--porcelain"], { cwd: root, encoding: "utf8" }).trim() !== "";
const report = {
  schemaVersion: 1, revision, dirty, capturedAt: new Date().toISOString(),
  environment: { node: process.version, platform: process.platform, arch: process.arch, osRelease: os.release(), cpu: os.cpus()[0]?.model, logicalCpus: os.cpus().length, totalMemoryBytes: os.totalmem() },
  packages: Object.fromEntries(["core", "layout", "pptx-exporter", "svg-renderer"].map((name) => [name, JSON.parse(readFileSync(path.join(root, "packages", name, "package.json"), "utf8")).version])),
  workload: { version: 1, fonts: ["Arial"], measurement: "Core heuristic; no native/browser font measurement", asset: "repository-authored 8x4 PNG data URL; no network", exportIncludesNormalizationAndLayout: true, memory: "isolated process per size; sampled RSS/heap plus process lifetime RSS peak, not per-stage peaks" },
  results: {},
};
for (const [name, count] of Object.entries(workloads)) {
  const child = spawnSync(process.execPath, ["--expose-gc", path.join(root, "scripts/benchmarks/worker.mjs"), String(count), String(options.iterations), String(options.warmup)], { cwd: root, encoding: "utf8", timeout: 300000, maxBuffer: 2 ** 20 });
  if (child.status !== 0) throw new Error(`${name} benchmark failed: ${child.error?.message ?? child.stderr}`);
  report.results[name] = JSON.parse(child.stdout);
  console.log(`${name}: ${count} slides; export median ${report.results[name].metrics.exportPipelineMs.median.toFixed(1)} ms; ${report.results[name].outputBytes} bytes`);
}
const output = path.resolve(root, options.output);
mkdirSync(path.dirname(output), { recursive: true });
writeFileSync(output, `${JSON.stringify(report, null, 2)}\n`);
console.log(`Benchmark report: ${output}`);
