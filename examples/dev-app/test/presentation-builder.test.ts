import { describe, expect, it } from "vitest";
import { createExamplePresentation } from "../src/presentation-builder";
import { parseExampleSource } from "../src/source-parser";
import { saasHuntSwissStyleExample } from "../src/examples/saas-hunt-swiss-style";
import { generatePptx } from "@pptkit/pptx-exporter";

const mixedSource = JSON.stringify({
  title: "Mixed deck",
  summary: "Text, shape, and image coverage.",
  slides: [
    {
      title: "Overview",
      elements: [
        "Plain text still works",
        {
          type: "shape",
          shape: "ellipse",
          style: { fill: "#FDE68A" },
        },
        {
          type: "image",
          altText: "Accent pixel",
          asset: {
            id: "accent-pixel",
            source: {
              type: "url",
              value:
                "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9WnR8WQAAAAASUVORK5CYII=",
            },
            mimeType: "image/png",
            altText: "Accent pixel",
          },
        },
      ],
    },
  ],
});

describe("example presentation builder", () => {
  it("parses mixed element source payloads", () => {
    const parsed = parseExampleSource(mixedSource);

    expect(parsed.title).toBe("Mixed deck");
    expect(parsed.slides[0]?.elements).toHaveLength(3);
    expect(parsed.slides[0]?.elements[0]).toBe("Plain text still works");
    expect(parsed.slides[0]?.elements[1]).toMatchObject({ type: "shape", shape: "ellipse" });
    expect(parsed.slides[0]?.elements[2]).toMatchObject({
      type: "image",
      altText: "Accent pixel",
      asset: { id: "accent-pixel" },
    });
  });

  it("builds text, shape, and image elements while registering image assets", () => {
    const presentation = createExamplePresentation(parseExampleSource(mixedSource));
    const slide = presentation.slides[0];

    expect(slide).toBeDefined();
    expect(slide?.elements.map((element) => element.type)).toEqual(["text", "shape", "image"]);
    expect(slide?.elements[0]).toMatchObject({
      type: "text",
      text: "Plain text still works",
    });
    expect(slide?.elements[1]).toMatchObject({
      type: "shape",
      shape: "ellipse",
      style: { fill: "#FDE68A" },
    });
    expect(slide?.elements[2]).toMatchObject({
      type: "image",
      assetId: "accent-pixel",
      altText: "Accent pixel",
    });
    expect(presentation.assets).toHaveLength(1);
    expect(presentation.assets[0]).toMatchObject({
      id: "accent-pixel",
      source: { type: "url" },
    });
  });

  it("parses and builds the complete SaaS Hunt Swiss Style reference", async () => {
    const parsed = parseExampleSource(saasHuntSwissStyleExample.source.content);
    const presentation = createExamplePresentation(parsed);

    expect(presentation.size).toEqual({ width: 960, height: 540, unit: "pt" });
    expect(presentation.slides).toHaveLength(7);
    expect(presentation.slides.map((slide) => slide.elements.length)).toEqual([12, 11, 18, 20, 13, 13, 11]);
    expect(presentation.slides.every((slide) => slide.background === "#F7F5EF")).toBe(true);
    expect(presentation.slides.flatMap((slide) => slide.elements).filter((element) => element.type === "text")).toHaveLength(66);
    expect(presentation.slides.flatMap((slide) => slide.elements).filter((element) => element.type === "shape")).toHaveLength(32);

    const result = await generatePptx(presentation);
    expect(result.slideCount).toBe(7);
    expect(result.status).toBe("generated");
    expect(result.warnings).toEqual([]);
  });
});
