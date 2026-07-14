import { test, expect } from "@playwright/test";
import { readFileSync } from "node:fs";

import { createPresentation } from "@pptkit/core";
import { renderPresentationToSvg } from "../dist/index.js";

const fixtureFont = readFileSync(new URL("./fixtures/roboto-mono-latin-400-normal.woff2", import.meta.url)).toString("base64");

function visualFixture() {
  const presentation = createPresentation({
    metadata: { title: "SVG visual regression" },
    theme: { colors: { accent1: "2457D6", accent2: "D94A66" }, fonts: { heading: "PPTKit Fixture", body: "PPTKit Fixture" } },
  });
  const asset = presentation.registerAsset({
    id: "visual-image",
    kind: "image",
    source: { type: "url", value: "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='400' height='200'%3E%3Crect width='400' height='200' fill='%232457D6'/%3E%3Ccircle cx='200' cy='100' r='70' fill='%23F8FAFC'/%3E%3C/svg%3E" },
    mimeType: "image/svg+xml",
    width: 400,
    height: 200,
  });
  const slide = presentation.addSlide({ id: "visual", background: { type: "solid", color: "F8FAFC" } });
  slide.addElement({
    type: "text", id: "visual-text", content: [{ style: { bullet: { type: "bullet" }, lineSpacing: 1.2 }, runs: [
      { text: "Rich text ", style: { fontFamily: "PPTKit Fixture", fontSize: 24, bold: true, color: { theme: "accent1" } } },
      { text: "preview", style: { fontFamily: "PPTKit Fixture", fontSize: 24, italic: true, underline: true } },
    ] }], box: { x: 40, y: 35, width: 400, height: 90 }, frame: { verticalAlign: "middle" },
  });
  slide.addElement({ type: "shape", id: "visual-shape", shape: "roundRect", box: { x: 40, y: 145, width: 220, height: 120 }, transform: { rotation: -4 }, style: { fill: { type: "solid", color: "DBEAFE" }, stroke: { paint: { type: "solid", color: { theme: "accent1" } }, width: 2 } }, text: { content: "Editable shape", frame: { verticalAlign: "middle" } } });
  slide.addElement({ type: "connector", id: "visual-connector", start: { x: 275, y: 205 }, route: [{ x: 325, y: 170 }], end: { x: 390, y: 205 }, style: { paint: { type: "solid", color: { theme: "accent2" } }, width: 3, beginArrow: "oval", endArrow: "triangle" } });
  slide.addElement({ type: "image", id: "visual-crop", assetId: asset.id, fit: "crop", crop: { left: 0.2, right: 0.2, top: 0, bottom: 0 }, box: { x: 430, y: 35, width: 250, height: 150 } });
  slide.addElement({ type: "group", id: "visual-group", box: { x: 710, y: 35, width: 200, height: 150 }, coordinateSize: { width: 200, height: 150 }, transform: { rotation: 3 }, children: [
    { type: "shape", id: "visual-group-diamond", shape: "diamond", box: { x: 10, y: 20, width: 100, height: 100 }, style: { fill: { type: "solid", color: "FBCFE8" } } },
    { type: "text", id: "visual-group-label", content: "Group", box: { x: 105, y: 55, width: 85, height: 40 } },
  ] });
  slide.addElement({ type: "table", id: "visual-table", box: { x: 40, y: 310, width: 870, height: 170 }, columns: [290, 290, 290], rows: [
    { height: 55, cells: [{ content: "Merged heading", colSpan: 3, style: { fill: { type: "solid", color: "E0E7FF" } } }] },
    { height: 55, cells: [{ content: "Alpha" }, { content: "Beta" }, { content: "Gamma" }] },
    { height: 55, cells: [{ content: "Row span", rowSpan: 2 }, { content: "Delta" }, { content: "Epsilon" }] },
  ] });
  return presentation;
}

test("renders hybrid SVG fixtures in a real browser", async ({ page }) => {
  const result = await renderPresentationToSvg(visualFixture());
  expect(result.warnings).toEqual([]);
  await page.setContent(`<style>@font-face{font-family:'PPTKit Fixture';src:url(data:font/woff2;base64,${fixtureFont}) format('woff2');font-style:normal;font-weight:100 900;font-display:block}html,body{margin:0;background:#d1d5db}#frame{width:960px;height:540px;background:white}svg{display:block;width:960px;height:540px}</style><div id="frame">${result.slides[0].svg}</div>`);
  await page.evaluate(() => document.fonts.ready);

  await expect(page.locator("svg[data-pptkit-slide-id='visual']")).toHaveCount(1);
  await expect(page.locator("foreignObject[data-pptkit-element-id='visual-text']")).toHaveCount(1);
  await expect(page.locator("[data-pptkit-element-id='visual-table'] table")).toHaveCount(1);
  await expect(page.locator("#frame")).toHaveScreenshot("composite.png");

  const clips = [
    ["rich-text.png", { x: 25, y: 20, width: 420, height: 115 }],
    ["shape-connector.png", { x: 20, y: 125, width: 390, height: 165 }],
    ["image-crop.png", { x: 415, y: 20, width: 280, height: 180 }],
    ["nested-group.png", { x: 695, y: 20, width: 230, height: 180 }],
    ["table-spans.png", { x: 25, y: 295, width: 900, height: 200 }],
  ];
  for (const [name, clip] of clips) {
    expect(await page.screenshot({ clip, animations: "disabled" })).toMatchSnapshot(name);
  }
});

test("keeps long text inside its authored SVG bounds", async ({ page }) => {
  const presentation = createPresentation({
    theme: { fonts: { heading: "Arial", body: "Arial" } },
  });
  const slide = presentation.addSlide({ id: "text-units" });
  slide.addElement({
    type: "text",
    id: "long-text",
    content: [{
      style: { lineSpacing: 1.1 },
      runs: [{
        text: "Feature: bold, italic, underline, strike, color, size, and auto-fit",
        style: { fontFamily: "Arial", fontSize: 24 },
      }],
    }],
    box: { x: 48, y: 72, width: 624, height: 48 },
    frame: { margin: 0, autoFit: { mode: "shrink", fontScale: 0.8 } },
  });

  const result = await renderPresentationToSvg(presentation);
  await page.setContent(result.slides[0].svg);
  const textLines = page.locator("g[data-pptkit-element-id='long-text'] text[data-pptkit-text-line]");
  await expect(textLines).toHaveCount(1);
  await expect(textLines.locator("tspan")).toHaveAttribute("font-size", "19.2");
});

test("uses PowerPoint-oriented line wrapping for compact text boxes", async ({ page }) => {
  const presentation = createPresentation({
    theme: { fonts: { heading: "Helvetica Neue", body: "Helvetica Neue" } },
  });
  const slide = presentation.addSlide({ id: "font-overflow" });
  slide.addElement({
    type: "text",
    id: "discovery",
    content: [{
      style: { lineSpacing: 1 },
      runs: [{
        text: "Discovery becomes easier.",
        style: { fontFamily: "Helvetica Neue", fontSize: 21.75, bold: true },
      }],
    }],
    box: { x: 504, y: 243, width: 255, height: 39 },
    frame: { autoFit: { mode: "shrink", fontScale: 0.92595 } },
  });

  const result = await renderPresentationToSvg(presentation);
  await page.setContent(result.slides[0].svg);
  const lines = await page.locator("g[data-pptkit-element-id='discovery'] text[data-pptkit-text-line]").allTextContents();
  expect(lines).toEqual(["Discovery becomes", "easier."]);
  await expect(page.locator("foreignObject[data-pptkit-element-id='discovery']")).toHaveCount(0);
});
