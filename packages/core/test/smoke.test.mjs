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

test("auto-sizes text height from width, paragraphs, and styles", () => {
  const presentation = createPresentation();
  const slide = presentation.addSlide({
    elements: [
      {
        type: "text",
        content: "Short line",
        box: { x: 40, y: 40, width: 200 },
      },
      {
        type: "text",
        content: "A deliberately long line that wraps within the fixed text width",
        box: { x: 40, y: 90, width: 200 },
        frame: { margin: { top: 10, bottom: 10 } },
      },
      {
        type: "text",
        content: [{
          style: { bullet: { type: "bullet" }, lineSpacing: 1.5, spaceAfter: 8 },
          runs: [{ text: "First" }],
        }, {
          style: { bullet: { type: "bullet" }, lineSpacing: 1.5, spaceAfter: 8 },
          runs: [{ text: "Second" }],
        }],
        box: { x: 40, y: 180, width: 200 },
      },
      {
        type: "text",
        content: "Explicit height remains authoritative",
        box: { x: 40, y: 280, width: 200, height: 42 },
      },
    ],
  });

  const normalized = normalizePresentation(presentation);
  const [shortText, wrappedText, bulletText, explicitText] = normalized.slides[0].elements;

  assert.equal(shortText.type, "text");
  assert.equal(wrappedText.type, "text");
  assert.equal(bulletText.type, "text");
  assert.equal(explicitText.type, "text");
  assert.ok(shortText.box.height >= 28);
  assert.ok(wrappedText.box.height > shortText.box.height);
  assert.ok(bulletText.box.height > shortText.box.height);
  assert.equal(explicitText.box.height, 42);
  assert.equal(slide.elements[0].box.height, undefined);
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

test("shape text and document text style presets normalize as one editable object", () => {
  const presentation = createPresentation({
    textStylePresets: {
      title: {
        frame: { margin: 0, verticalAlign: "middle" },
        paragraph: { align: "center" },
        run: { fontSize: 30, bold: true, color: "123456" },
      },
      cell: { frame: { margin: 0, verticalAlign: "middle" }, paragraph: { align: "right" }, run: { fontSize: 12 } },
    },
  });
  assert.throws(() => { presentation.textStylePresets.title.run.fontSize = 40; }, TypeError);

  const slide = presentation.addSlide({ elements: [
    {
      type: "shape",
      id: "title-card",
      shape: "roundRect",
      box: { x: 20, y: 20, width: 300, height: 80 },
      text: { content: [{ style: { align: "left" }, runs: [{ text: "Hello" }] }], textStylePreset: "title" },
    },
    {
      type: "table",
      box: { x: 20, y: 120, width: 300, height: 40 },
      columns: [300],
      rows: [{ cells: [{ content: "Cell", textStylePreset: "cell" }] }],
    },
  ] });

  const duplicate = slide.duplicateElement("title-card");
  assert.notEqual(duplicate.id, "title-card");
  const normalized = normalizePresentation(presentation);
  const shape = normalized.slides[0].elements[0];
  assert.equal(shape.type, "shape");
  assert.equal(shape.text.plainText, "Hello");
  assert.equal(shape.text.frame.verticalAlign, "middle");
  assert.equal(shape.text.content[0].style.align, "left");
  assert.equal(shape.text.content[0].runs[0].style.fontSize, 30);
  assert.equal(shape.text.content[0].runs[0].style.bold, true);
  assert.equal("textStylePreset" in shape, false);
  assert.equal(normalized.slides[0].elements[2].rows[0].cells[0].content[0].style.align, "right");
  assert.equal(normalized.slides[0].elements[2].rows[0].cells[0].style.margin.left, 0);
  assert.equal(normalized.slides[0].elements[2].rows[0].cells[0].style.verticalAlign, "middle");
});
