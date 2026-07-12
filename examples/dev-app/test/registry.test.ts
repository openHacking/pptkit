import { describe, expect, it } from "vitest";
import { listExamples, listFeatures } from "../src/example-registry";
import { buildExampleReport } from "../src/report-builder";

describe("example registry", () => {
  it("ensures every feature has at least one example", () => {
    const features = listFeatures();
    const examples = listExamples();

    for (const feature of features) {
      expect(examples.some((example) => example.feature === feature)).toBe(true);
    }
  });

  it("keeps example ids unique", () => {
    const ids = listExamples().map((example) => example.id);
    expect(ids).toHaveLength(new Set(ids).size);
  });

  it("builds placeholder-friendly report data", async () => {
    const example = listExamples()[0];

    expect(example).toBeDefined();

    const report = await buildExampleReport(example!);

    expect(report.normalizedDocument.slideCount).toBeGreaterThan(0);
    expect(report.visualPreview.status).toBe("structural-preview");
    expect(typeof report.exportResult.status).toBe("string");
    expect(Array.isArray(report.diagnostics)).toBe(true);
  });
});
