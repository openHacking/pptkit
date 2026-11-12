import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { chromium } from "@playwright/test";
const directory = process.argv[2];
if (!directory || process.argv.length !== 3) throw new Error("Usage: node scripts/inspect-gif-browser-playback.mjs <directory>");
const svg = await readFile(join(directory, "preview.svg"), "utf8");
const browser = await chromium.launch({ headless: true });
try {
  const page = await browser.newPage({ viewport: { width: 960, height: 540 }, deviceScaleFactor: 1 });
  await page.route("**/*", route => route.abort());
  await page.setContent(`<style>html,body{margin:0}</style>${svg}`);
  const times = [];
  for (let i = 0; i < 150; i++) {
    const start = Date.now(), bytes = await page.screenshot(), end = Date.now();
    const file = `browser-${i}.png`; await writeFile(join(directory, file), bytes); times.push({ file, start, end });
  }
  await writeFile(join(directory, "browser-times.json"), JSON.stringify({ browserVersion: browser.version(), samples: times }, null, 2) + "\n");
  console.log(JSON.stringify({ browserVersion: browser.version(), samples: times.length, elapsedMs: times.at(-1).end - times[0].start }));
} finally { await browser.close(); }
