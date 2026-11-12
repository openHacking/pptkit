import { defineConfig } from '@playwright/test';
import os from 'node:os';
import path from 'node:path';

export default defineConfig({
  testDir: './test',
  testMatch: ['visual.spec.mjs'],
  grep: /browser raster|multi-slide raster|strict raster|covered raster|coverage callback abort|coverage requests follow|static image attestation|static image policy|JPEG opacity inspection/,
  workers: 1,
  fullyParallel: false,
  reporter: 'line',
  outputDir: path.join(os.tmpdir(), 'pptkit-raster-cross-playwright'),
  use: { viewport: { width: 1000, height: 700 }, deviceScaleFactor: 1, colorScheme: 'light', reducedMotion: 'reduce' },
  projects: [{ name: 'chromium', use: { browserName: 'chromium' } }, { name: 'firefox', use: { browserName: 'firefox' } }, { name: 'webkit', use: { browserName: 'webkit' } }],
});
