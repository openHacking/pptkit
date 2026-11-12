import { defineConfig } from '@playwright/test';
import os from 'node:os';
import path from 'node:path';

export default defineConfig({
  testDir: './test',
  testMatch: ['visual.spec.mjs'],
  grep: /document slide links navigate/,
  workers: 1,
  reporter: 'line',
  outputDir: path.join(os.tmpdir(), 'pptkit-document-links-playwright'),
  use: { viewport: { width: 1000, height: 700 }, deviceScaleFactor: 1 },
  projects: [
    { name: 'chromium', use: { browserName: 'chromium' } },
    { name: 'firefox', use: { browserName: 'firefox' } },
    { name: 'webkit', use: { browserName: 'webkit' } },
  ],
});
