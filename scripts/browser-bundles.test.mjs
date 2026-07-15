import assert from "node:assert/strict";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { chromium } from "@playwright/test";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const browserPackages = ["core", "layout", "pptx-exporter", "svg-renderer"];
function bundlePath(packageName, format) {
  return path.join(root, "packages", packageName, "dist", "browser", format === "esm" ? "index.js" : "global.js");
}

function javascriptDataUrl(source) {
  return `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;
}

test("published browser artifacts are single-file and browser-safe", async () => {
  for (const packageName of browserPackages) {
    for (const format of ["esm", "global"]) {
      const file = bundlePath(packageName, format);
      assert.ok((await stat(file)).size > 0, `${file} should not be empty`);
      const source = await readFile(file, "utf8");
      assert.doesNotMatch(source, /(?:^|["'])node:/);
      assert.doesNotMatch(source, /\brequire\s*\(/);
      assert.doesNotMatch(source, /\bprocess(?:\.|\[)/);
      assert.doesNotMatch(source, /sourceMappingURL/);
    }

    const packageRoot = path.join(root, "packages", packageName);
    const packed = spawnSync("npm", ["pack", "--dry-run", "--json"], {
      cwd: packageRoot,
      encoding: "utf8",
      env: { ...process.env, npm_config_cache: path.join("/tmp", "pptkit-npm-cache") },
    });
    assert.equal(packed.status, 0, packed.stderr);
    const [{ files }] = JSON.parse(packed.stdout);
    const names = new Set(files.map((file) => file.path));
    assert.ok(names.has("dist/browser/index.js"));
    assert.ok(names.has("dist/browser/global.js"));
  }

  const layoutEsm = await readFile(bundlePath("layout", "esm"), "utf8");
  const exporterEsm = await readFile(bundlePath("pptx-exporter", "esm"), "utf8");
  const rendererEsm = await readFile(bundlePath("svg-renderer", "esm"), "utf8");
  assert.match(layoutEsm, /from["']@pptkit\/core["']/);
  assert.match(exporterEsm, /from["']@pptkit\/core["']/);
  assert.match(exporterEsm, /from["']@pptkit\/layout["']/);
  assert.doesNotMatch(exporterEsm, /from["']fflate["']/);
  assert.match(rendererEsm, /from["']@pptkit\/core["']/);
  assert.match(rendererEsm, /from["']@pptkit\/layout["']/);
});

test("official ESM example previews and downloads a PPTX in Chromium", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  const cdnSources = {
    core: await readFile(bundlePath("core", "esm"), "utf8"),
    layout: await readFile(bundlePath("layout", "esm"), "utf8"),
    "pptx-exporter": await readFile(bundlePath("pptx-exporter", "esm"), "utf8"),
    "svg-renderer": await readFile(bundlePath("svg-renderer", "esm"), "utf8"),
  };
  let html = await readFile(path.join(root, "examples", "browser-esm", "index.html"), "utf8");
  for (const [packageName, source] of Object.entries(cdnSources)) {
    html = html.replace(
      `https://cdn.jsdelivr.net/npm/@pptkit/${packageName}@latest/dist/browser/index.js`,
      javascriptDataUrl(source),
    );
  }
  const appSource = await readFile(path.join(root, "examples", "browser-esm", "app.js"), "utf8");
  html = html.replace("./app.js", javascriptDataUrl(appSource));
  await page.setContent(html, { waitUntil: "load" });
  await page.evaluate(() => {
    const createObjectURL = URL.createObjectURL.bind(URL);
    URL.createObjectURL = (blob) => {
      globalThis.__pptkitBlobType = blob.type;
      return createObjectURL(blob);
    };
  });
  await page.getByRole("button", { name: "Preview SVG" }).click();
  await assert.doesNotReject(() => page.waitForFunction(() => document.querySelectorAll("#preview svg").length === 2));

  const downloadPromise = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download PPTX" }).click();
  const download = await downloadPromise;
  assert.equal(download.suggestedFilename(), "pptkit-browser-example.pptx");
  const downloadPath = await download.path();
  assert.ok(downloadPath !== null);
  const bytes = await readFile(downloadPath);
  assert.deepEqual(Array.from(bytes.subarray(0, 2)), [0x50, 0x4b]);
  assert.equal(await page.evaluate(() => globalThis.__pptkitBlobType), "application/vnd.openxmlformats-officedocument.presentationml.presentation");
  await assert.doesNotReject(() => page.waitForFunction(() => document.querySelector("#status")?.textContent?.includes("Generated 2 slides")));
});

test("global bundles compose into one PPTKit namespace", async (t) => {
  const browser = await chromium.launch({ headless: true });
  t.after(() => browser.close());
  const page = await browser.newPage();
  for (const packageName of browserPackages) {
    await page.addScriptTag({ content: await readFile(bundlePath(packageName, "global"), "utf8") });
  }
  const result = await page.evaluate(async () => {
    const presentation = PPTKit.createPresentation();
    presentation.addSlide({ elements: [{ type: "text", content: "Global bundle", box: { x: 20, y: 20, width: 300, height: 60 } }] });
    presentation.addSlide();
    const [preview, generated] = await Promise.all([
      PPTKit.renderPresentationToSvg(presentation),
      PPTKit.generatePptx(presentation),
    ]);
    return {
      namespace: ["createPresentation", "resolveLayout", "generatePptx", "renderPresentationToSvg"].every((key) => typeof PPTKit[key] === "function"),
      previews: preview.slides.length,
      slideCount: generated.slideCount,
      signature: Array.from(generated.bytes.slice(0, 2)),
    };
  });
  assert.deepEqual(result, {
    namespace: true,
    previews: 2,
    slideCount: 2,
    signature: [0x50, 0x4b],
  });
});
