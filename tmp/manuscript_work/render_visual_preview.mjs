import { chromium } from 'file:///C:/Users/DyyYah/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs';
import path from 'node:path';

const browser = await chromium.launch({headless: true, executablePath: 'C:/Program Files/Google/Chrome/Application/chrome.exe'});
const page = await browser.newPage();
const htmlPath = path.resolve('tmp/manuscript_work/final_visual_preview.html');
await page.goto('file:///' + htmlPath.replaceAll('\\', '/'), {waitUntil: 'networkidle'});
await page.pdf({
  path: 'tmp/manuscript_work/final_visual_preview.pdf',
  format: 'A4',
  printBackground: true,
  preferCSSPageSize: true,
});
await browser.close();
