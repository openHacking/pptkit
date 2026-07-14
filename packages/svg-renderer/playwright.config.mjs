import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: "./test",
  testMatch: "visual.spec.mjs",
  fullyParallel: false,
  workers: 1,
  reporter: "line",
  outputDir: "/tmp/pptkit-svg-renderer-playwright",
  snapshotPathTemplate: "{testDir}/snapshots/{arg}{ext}",
  expect: {
    toHaveScreenshot: { animations: "disabled", maxDiffPixelRatio: 0.04, threshold: 0.25 },
    toMatchSnapshot: { maxDiffPixelRatio: 0.04, threshold: 0.25 },
  },
  use: {
    ...devices["Desktop Chrome"],
    viewport: { width: 1000, height: 700 },
    deviceScaleFactor: 1,
    colorScheme: "light",
    reducedMotion: "reduce",
    ...(process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH === undefined ? {} : {
      launchOptions: { executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE_PATH },
    }),
  },
  projects: [{ name: "chromium", use: { browserName: "chromium" } }],
});
