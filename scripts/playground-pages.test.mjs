import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { createServer } from "node:http";
import test from "node:test";
import { chromium, expect } from "@playwright/test";
import { importPptx, inspectPptx } from "../packages/pptx-parser/dist/index.js";

const nativeSample = new URL("../tests/fixtures/corpus-text-powerpoint-interactive.pptx", import.meta.url);
const dist = new URL("../examples/dev-app/dist/playground/", import.meta.url);
const types = { html: "text/html", js: "text/javascript", css: "text/css", pptx: "application/octet-stream" };

async function downloadBytes(page, button, filename) {
  const [download] = await Promise.all([page.waitForEvent("download"), page.getByRole("button", { name: button, exact: true }).click()]);
  assert.equal(download.suggestedFilename(), filename);
  return readFile(await download.path());
}

test("the Pages artifact works at /pptkit/ without runtime APIs", { timeout: 60_000 }, async (t) => {
  const requests = [];
  const server = createServer(async (request, response) => {
    requests.push(request.url);
    const asset = request.url === "/pptkit/" ? "index.html" : /^\/pptkit\/(assets\/[\w.-]+)$/.exec(request.url)?.[1];
    if (asset === undefined) return response.writeHead(404).end();
    try {
      const bytes = await readFile(new URL(asset, dist));
      response.writeHead(200, { "content-type": types[asset.split(".").at(-1)] }).end(bytes);
    } catch {
      response.writeHead(404).end();
    }
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => new Promise((resolve) => server.close(resolve)));
  const origin = `http://127.0.0.1:${server.address().port}`;
  const browser = await chromium.launch();
  t.after(() => browser.close());
  const page = await browser.newPage();
  page.setDefaultTimeout(5_000);
  const pageErrors = [];
  const failedAssets = [];
  const external = [];
  page.on("pageerror", (error) => pageErrors.push(error.message));
  page.on("response", (response) => { if (!response.ok()) failedAssets.push(response.url()); });
  page.on("request", (request) => {
    if (!request.url().startsWith(origin) && !request.url().startsWith("data:")) external.push(request.url());
  });

  await page.goto(`${origin}/pptkit/`);
  const frame = page.frameLocator('iframe[title="Slide preview"]');
  await expect(frame.locator("svg").first()).toBeVisible();
  assert.equal(await page.locator("iframe").getAttribute("sandbox"), "");
  const bounds = await page.locator("iframe").boundingBox();
  assert.ok(Math.abs(bounds.width / bounds.height - 16 / 9) < 0.01);
  await expect(page.getByText("Slide 1 of 11")).toBeVisible();
  await page.getByRole("button", { name: "Next slide" }).click();
  await expect(page.getByText("Slide 2 of 11")).toBeVisible();
  await page.getByRole("button", { name: "Previous slide" }).click();
  await page.getByText("Edit example JSON").click();
  const editor = page.getByRole("textbox", { name: "Example JSON" });
  const initial = await editor.inputValue();
  await editor.fill("{");
  await expect(page.getByRole("button", { name: "Download PPTX", exact: true })).toBeDisabled();
  await page.getByRole("button", { name: "Apply changes" }).click();
  await expect(page.getByRole("alert")).toBeVisible();
  const config = JSON.parse(initial);
  config.slides[0].elements[0] = "Edited on GitHub Pages";
  await editor.fill(JSON.stringify(config));
  await page.getByRole("button", { name: "Apply changes" }).click();
  await expect(frame.getByText("Edited on GitHub Pages")).toBeVisible();
  assert.match((await downloadBytes(page, "SVG", "export-full-feature-deck-slide-1.svg")).toString(), /Edited on GitHub Pages/);
  const deck = await downloadBytes(page, "Download PPTX", "export-full-feature-deck.pptx");
  const authored = await importPptx(deck, { projection: "content" });
  try {
    assert.equal(authored.document.slides.length, 11);
    assert.ok(authored.document.slides[0].elements.some(({ content }) =>
      Array.isArray(content) && content.some(({ runs }) => runs.some(({ text }) => text === "Edited on GitHub Pages"))));
  } finally {
    authored.dispose();
  }

  for (const title of ["Rich text & notes", "Themes & layouts", "Shapes & connectors", "Images & cropping", "Merged table cells", "Bar, line & pie charts", "Missing asset diagnostics", "Advanced charts", "Fade transition"]) {
    await page.getByRole("button", { name: title, exact: true }).click();
    await expect(page.getByRole("heading", { name: title, exact: true })).toBeVisible();
    await expect(frame.locator("svg").first()).toBeVisible();
  }
  await page.goto(`${origin}/pptkit/#advanced-charts`);
  await page.reload();
  await expect(page.getByRole("heading", { name: "Advanced charts" })).toBeVisible();
  assert.equal((await inspectPptx(await downloadBytes(page, "Download PPTX", "advanced-charts.pptx"))).slides.length, 5);

  await page.getByRole("button", { name: "PowerPoint import" }).click();
  await expect(page.getByRole("heading", { name: "PowerPoint import" })).toBeVisible();
  assert.deepEqual(await downloadBytes(page, "Download preserved PPTX", "powerpoint-sample-preserved.pptx"), await readFile(nativeSample));
  await page.getByText("Edit a preserved text run").click();
  await page.getByLabel("Text from the first supported text box").fill("Preserved edit");
  const edited = await downloadBytes(page, "Save edited PPTX", "powerpoint-sample-edited.pptx");
  const session = await importPptx(edited, { projection: "content" });
  try {
    assert.ok(session.document.slides[0].elements.some(({ content }) =>
      Array.isArray(content) && content.some(({ runs }) => runs.some(({ text }) => text === "Preserved edit"))));
  } finally {
    session.dispose();
  }
  const input = page.locator("input[type=file]");
  await input.setInputFiles({ name: "invalid.pptx", mimeType: types.pptx, buffer: Buffer.from("bad file") });
  await expect(page.getByRole("alert")).toBeVisible();
  await input.setInputFiles({ name: "own.pptx", mimeType: types.pptx, buffer: deck });
  await expect(page.getByText(/own\.pptx/)).toBeVisible();
  assert.deepEqual(await downloadBytes(page, "Download preserved PPTX", "own-preserved.pptx"), deck);

  await page.setViewportSize({ width: 375, height: 800 });
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth), 375);
  assert.deepEqual(requests.filter((url) => !url.startsWith("/pptkit/")), []);
  assert.deepEqual(failedAssets, []);
  assert.deepEqual(external, []);
  assert.deepEqual(pageErrors, []);
});
