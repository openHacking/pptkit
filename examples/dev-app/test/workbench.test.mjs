import test from "node:test";
import assert from "node:assert/strict";

const { listExamples, listFeatures } = await import("../dist/example-registry.js");
const { buildExampleReport } = await import("../dist/report-builder.js");

test("every feature has at least one example", () => {
  const features = listFeatures();
  const examples = listExamples();

  for (const feature of features) {
    assert.ok(
      examples.some((example) => example.feature === feature),
      `expected at least one example for feature ${feature}`,
    );
  }
});

test("example ids are unique", () => {
  const ids = listExamples().map((example) => example.id);
  assert.equal(ids.length, new Set(ids).size);
});

test("report builder returns placeholder-friendly preview data", async () => {
  const example = listExamples()[0];
  assert.ok(example);

  const report = await buildExampleReport(example);

  assert.equal(report.normalizedDocument.slideCount > 0, true);
  assert.equal(report.visualPreview.status, "structural-preview");
  assert.equal(typeof report.exportResult.status, "string");
  assert.equal(Array.isArray(report.diagnostics), true);
});
