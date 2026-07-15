import assert from "node:assert/strict";
import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { chromium } from "@playwright/test";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function mimeType(file) {
  if (file.endsWith(".html")) return "text/html";
  if (file.endsWith(".js")) return "text/javascript";
  if (file.endsWith(".css")) return "text/css";
  return "application/octet-stream";
}

async function serve() {
  const server = createServer(async (request, response) => {
    const requestPath = request.url === "/" ? "/index.html" : request.url.split("?")[0];
    const file = path.join(root, "dist", requestPath);
    try {
      const info = await stat(file);
      if (!info.isFile()) throw new Error("not a file");
      response.writeHead(200, { "content-type": mimeType(file) });
      response.end(await readFile(file));
    } catch {
      response.writeHead(404);
      response.end("Not found");
    }
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const address = server.address();
  return { server, url: `http://127.0.0.1:${address.port}` };
}

function fixture(revision = 1) {
  const now = new Date().toISOString();
  return {
    schemaVersion: 1,
    id: "browser-review",
    revision,
    createdAt: now,
    updatedAt: now,
    deck: {
      brief: { title: "Browser Review", audience: "QA", purpose: "Review before download", language: "en-US", slideCountRange: [3, 3], themeId: "clean-business", imagePolicy: "Local", constraints: [] },
      slides: [
        { id: "cover", role: "cover", title: "Browser Review", subtitle: "SVG first" },
        { id: "process", role: "process", title: revision === 1 ? "Review loop" : "Updated review loop", steps: ["Import", "Preview", "Revise", "Download"] },
        { id: "closing", role: "closing", title: "Approved", message: "Download only on request." },
      ],
    },
    sources: [],
    assets: [],
  };
}

test("imports, persists, revises, previews, and exports a browser deck session", async (t) => {
  const { server, url } = await serve();
  t.after(() => server.close());
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage({ acceptDownloads: true });
  await page.goto(url);
  await page.locator("#session-input").fill(JSON.stringify(fixture()));
  await page.getByRole("button", { name: "Import and preview" }).click();
  await page.waitForFunction(() => document.querySelectorAll("#thumbnails button").length === 3);
  assert.equal(await page.locator("#stage svg").count(), 1);
  assert.match(await page.locator("#status").innerText(), /Saved in this browser/);
  assert.equal(await page.getByRole("button", { name: "Generate & download PPTX" }).isEnabled(), true);
  assert.equal(await page.evaluate(() => globalThis.__pptkitPreviewState.exportStatus), "not-run");

  await page.getByRole("button", { name: "Next" }).click();
  assert.equal(await page.locator("#page-status").innerText(), "2 / 3");
  await page.getByRole("button", { name: "Import deck session" }).click();
  await page.locator("#session-input").fill(JSON.stringify(fixture(2)));
  await page.getByRole("button", { name: "Import and preview" }).click();
  await page.waitForFunction(() => document.querySelector("#status")?.textContent?.includes("Changed slides: process"));
  assert.equal(await page.locator("#page-status").innerText(), "2 / 3");
  assert.match(await page.locator("#status").innerText(), /Changed slides: process/);

  await page.reload();
  await page.waitForFunction(() => document.querySelectorAll("#thumbnails button").length === 3);
  assert.match(await page.locator("#deck-meta").innerText(), /revision 2/);

  const downloads = [];
  page.on("download", (download) => downloads.push(download));
  await page.getByRole("button", { name: "Generate & download PPTX" }).click();
  await page.waitForFunction(() => document.querySelector("#status")?.textContent?.includes("passed package inspection"));
  assert.equal(await page.evaluate(() => globalThis.__pptkitPreviewState.exportStatus), "generated");
  await new Promise((resolve) => setTimeout(resolve, 100));
  assert.deepEqual(downloads.map((download) => download.suggestedFilename()).sort(), ["browser-review.pptx", "build-report.json"]);
  const pptx = downloads.find((download) => download.suggestedFilename().endsWith(".pptx"));
  const pptxPath = await pptx.path();
  const bytes = await readFile(pptxPath);
  assert.deepEqual(Array.from(bytes.subarray(0, 2)), [0x50, 0x4b]);
});
