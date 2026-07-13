import test from "node:test";
import assert from "node:assert/strict";

import { createPresentation, normalizePresentation } from "../dist/index.js";

test("authoring collections are method-managed and stable", () => {
  const presentation = createPresentation({
    metadata: { title: "Roadmap", author: "PPTKit Team" },
    theme: { fonts: { heading: "Arial" } },
  });
  const first = presentation.addSlide({ id: "first" });
  const second = presentation.insertSlide(0, { id: "second" });
  const title = first.addElement({
    type: "text",
    content: "Launch plan",
    box: { x: 48, y: 48, width: 400, height: 32 },
  });
  const duplicate = first.duplicateElement(title.id);

  assert.deepEqual(presentation.slides.map((slide) => slide.id), ["second", "first"]);
  assert.notEqual(duplicate.id, title.id);
  first.moveElement(duplicate.id, 0);
  assert.equal(first.elements[0].id, duplicate.id);
  first.removeElement(title.id);
  assert.equal(first.elements.length, 1);
  presentation.moveSlide(second.id, 1);
  assert.deepEqual(presentation.slides.map((slide) => slide.id), ["first", "second"]);
  assert.throws(() => presentation.slides.push(first), TypeError);
  assert.throws(() => { duplicate.name = "changed"; }, TypeError);
});

test("duplicateSlide creates fresh nested element identities", () => {
  const presentation = createPresentation();
  const slide = presentation.addSlide({
    elements: [{
      type: "group",
      box: { x: 0, y: 0, width: 200, height: 100 },
      coordinateSize: { width: 200, height: 100 },
      children: [{ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 50, height: 50 } }],
    }],
  });
  const duplicate = presentation.duplicateSlide(slide.id);
  assert.notEqual(duplicate.id, slide.id);
  assert.notEqual(duplicate.elements[0].id, slide.elements[0].id);
  assert.notEqual(duplicate.elements[0].children[0].id, slide.elements[0].children[0].id);
});

test("normalizePresentation materializes IR v1 defaults and rich text", () => {
  const presentation = createPresentation({
    metadata: { title: "Quarterly Update", language: "zh-CN" },
    theme: { colors: { accent1: "112233" }, fonts: { body: "Microsoft YaHei" } },
  });
  presentation.defineSlideLayout({
    id: "title-layout",
    name: "Title",
    background: { type: "solid", color: { theme: "background2" } },
    placeholders: [{
      key: "title",
      kind: "title",
      box: { x: 40, y: 40, width: 600, height: 80 },
      textStyle: { run: { fontFamily: { theme: "heading" }, fontSize: 36, bold: true } },
    }],
  });
  const slide = presentation.addSlide({
    layoutId: "title-layout",
    notes: "Speaker note",
    section: "Opening",
    tags: ["intro"],
    customData: { source: "test" },
  });
  slide.addElement({
    type: "text",
    content: [{
      style: { bullet: { type: "bullet" }, lineSpacing: 1.2 },
      runs: [
        { text: "Hello ", style: { color: { theme: "text1" } } },
        { text: "PPTKit", style: { color: { theme: "accent1" }, italic: true } },
      ],
    }],
    placeholderKey: "title",
  });

  const normalized = normalizePresentation(presentation);
  const text = normalized.slides[0].elements[0];
  assert.equal(normalized.irVersion, 1);
  assert.equal(normalized.metadata.title, "Quarterly Update");
  assert.equal(normalized.theme.colors.accent1, "112233");
  assert.equal(normalized.slides[0].backgroundSource, "layout");
  assert.deepEqual(text.box, { x: 40, y: 40, width: 600, height: 80 });
  assert.equal(text.plainText, "Hello PPTKit");
  assert.equal(text.content[0].runs[0].style.fontSize, 36);
  assert.equal(text.content[0].runs[1].style.italic, true);
  assert.equal(normalized.slides[0].notes[0].runs[0].text, "Speaker note");
});

test("normalization detaches nested groups, tables, assets, and custom data", () => {
  const presentation = createPresentation();
  const asset = presentation.registerAsset({
    kind: "image",
    source: { type: "url", value: "https://example.com/hero.png" },
    width: 800,
    height: 400,
  });
  presentation.addSlide({
    customData: { nested: { value: 1 } },
    elements: [
      { type: "image", assetId: asset.id, fit: "cover", box: { x: 0, y: 0, width: 200, height: 200 } },
      {
        type: "table",
        box: { x: 0, y: 220, width: 300, height: 100 },
        columns: [150, 150],
        rows: [{ cells: [{ content: "A", colSpan: 2, style: { fill: { type: "solid", color: "FF0000" } } }] }],
      },
    ],
  });
  const normalized = normalizePresentation(presentation);
  assert.equal(normalized.assets[0].accessibility.decorative, false);
  assert.equal(normalized.slides[0].elements[0].fit, "cover");
  assert.equal(normalized.slides[0].elements[1].rows[0].cells[0].colSpan, 2);
  assert.equal(normalized.slides[0].elements[1].rows[0].cells[0].style.fill.opacity, 1);
  assert.notEqual(normalized.slides[0].customData, presentation.slides[0].customData);
});
