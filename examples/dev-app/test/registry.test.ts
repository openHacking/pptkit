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

  it("registers one labeled full-feature export case", () => {
    const examples = listExamples();
    const fullFeatureDeck = examples.find((example) => example.id === "export-full-feature-deck");

    expect(fullFeatureDeck).toBeDefined();
    expect(examples.some((example) => example.id === "export-pitch-deck")).toBe(false);
    expect(examples.some((example) => example.id === "export-mixed-media-slide")).toBe(false);

    const input = fullFeatureDeck!.createInput();
    expect(input.slides).toHaveLength(8);
    expect(input.slides.every((slide) => slide.title.startsWith("Feature:"))).toBe(true);
    expect(input.slides.every((slide) => slide.elements.length >= 1 && slide.elements.length <= 2)).toBe(true);
    expect(fullFeatureDeck!.source.content).toContain('"Feature: image asset + alt text"');
    expect(fullFeatureDeck!.source.content).toContain('"type": "bullet"');
    expect(fullFeatureDeck!.source.content).toContain('"italic": true');
    expect(fullFeatureDeck!.source.content).toContain('"type": "number"');
    expect(fullFeatureDeck!.source.content).toContain('"type": "group"');
    expect(fullFeatureDeck!.source.content).toContain('"type": "table"');
    expect(fullFeatureDeck!.source.content).toContain('"textStylePresets"');
    expect(fullFeatureDeck!.source.content).toContain('"textStylePreset": "shapeLabel"');
    expect(fullFeatureDeck!.source.content).toContain("Feature: editable text inside one rectangle shape");

    const numberedParagraph = input.slides[1]?.elements[1];
    expect(numberedParagraph).toMatchObject({
      type: "text",
      box: { x: 48, y: 152, width: 624 },
      text: expect.stringContaining("\nFirst numbered item\nSecond numbered item"),
    });
    expect((numberedParagraph as { box?: { height?: number } }).box?.height).toBeUndefined();

    const textShape = input.slides[2]?.elements[0];
    expect(textShape).toMatchObject({
      type: "shape",
      shape: "rect",
      text: { textStylePreset: "shapeLabel" },
    });
    expect(input.textStylePresets).toHaveProperty("featureText");
  });

  it("builds placeholder-friendly report data", async () => {
    const example = listExamples()[0];

    expect(example).toBeDefined();

    const report = await buildExampleReport(example!);

    expect(report.normalizedDocument.slideCount).toBeGreaterThan(0);
    expect(report.visualPreview.status).toBe("structural-preview");
    expect(report.exportResult.status).toBe("not-exported");
    expect(Array.isArray(report.diagnostics)).toBe(true);
  });
});
