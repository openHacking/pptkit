import assert from "node:assert/strict";
import test from "node:test";
import { createPresentation, normalizePresentation, validatePresentation } from "../packages/core/dist/index.js";
import { generatePptx } from "../packages/pptx-exporter/dist/index.js";
import { createBenchmarkDeck, summarize } from "./benchmarks/workload.mjs";

test("benchmark statistics preserve samples and expose variance instead of hiding it", () => {
  const values = [4, 1, 3, 2];
  const stats = summarize(values);
  assert.equal(stats.mean, 2.5);
  assert.equal(stats.median, 2.5);
  assert.equal(stats.p95, 4);
  assert.equal(stats.standardDeviation, Math.sqrt(1.25));
  assert.equal(summarize([0, 0]).coefficientOfVariation, 0);
  assert.deepEqual(values, [4, 1, 3, 2]);
  assert.throws(() => summarize([]));
  assert.throws(() => summarize([NaN]));
});

test("mixed benchmark workload validates, is detached, and produces repeatable nonempty output", async () => {
  const document = createBenchmarkDeck(createPresentation, 4);
  assert.deepEqual(validatePresentation(document), []);
  const input = normalizePresentation(document);
  assert.deepEqual(input, normalizePresentation(createBenchmarkDeck(createPresentation, 4)));
  const result = await generatePptx(document);
  assert.deepEqual(result.warnings, []);
  assert.equal(result.slideCount, 4);
  assert.ok(result.byteLength > 0);
  assert.ok(Buffer.from(result.bytes).equals(Buffer.from((await generatePptx(document)).bytes)));
  assert.deepEqual(normalizePresentation(document), input);
  assert.throws(() => createBenchmarkDeck(createPresentation, 0));
});
