import { chromium } from '@playwright/test';
import type { PdfBrowser } from '../src/index.js';

// The adapter deliberately accepts a real Playwright Chromium browser without a cast.
const browser: PdfBrowser = await chromium.launch();
await browser.newContext();
