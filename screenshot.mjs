import { createRequire } from 'module';
import { pathToFileURL } from 'url';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);

const url   = process.argv[2] || 'http://localhost:3000';
const label = process.argv[3] || '';

const screenshotsDir = path.join(__dirname, 'temporary screenshots');
if (!fs.existsSync(screenshotsDir)) fs.mkdirSync(screenshotsDir, { recursive: true });

const existing = fs.readdirSync(screenshotsDir).filter(f => f.endsWith('.png'));
const nums = existing.map(f => parseInt(f.match(/screenshot-(\d+)/)?.[1] || '0')).filter(n => !isNaN(n) && n > 0);
const next = nums.length ? Math.max(...nums) + 1 : 1;
const filename = label ? `screenshot-${next}-${label}.png` : `screenshot-${next}.png`;
const outPath = path.join(screenshotsDir, filename);

// Try multiple possible puppeteer locations
const puppeteerPaths = [
  'C:/Users/Jess/AppData/Local/Temp/puppeteer-test/node_modules/puppeteer',
  'C:/Users/nateh/AppData/Local/Temp/puppeteer-test/node_modules/puppeteer',
];

const chromePaths = [
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Users/Jess/.cache/puppeteer/chrome/win64-136.0.7103.92/chrome-win64/chrome.exe',
  'C:/Users/nateh/.cache/puppeteer/chrome/win64-136.0.7103.92/chrome-win64/chrome.exe',
];

let puppeteer;
for (const p of puppeteerPaths) {
  try { puppeteer = require(p); break; } catch(e) {}
}
if (!puppeteer) {
  try { puppeteer = require('puppeteer'); } catch(e) {}
}
if (!puppeteer) throw new Error('Puppeteer not found. Checked: ' + puppeteerPaths.join(', '));

let executablePath;
for (const p of chromePaths) {
  if (fs.existsSync(p)) { executablePath = p; break; }
}

(async () => {
  const launchOpts = {
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
  };
  if (executablePath) launchOpts.executablePath = executablePath;

  const browser = await puppeteer.launch(launchOpts);
  const page = await browser.newPage();
  await page.setViewport({ width: 390, height: 844, deviceScaleFactor: 2 });
  await page.goto(url, { waitUntil: 'networkidle2', timeout: 30000 });

  // Scroll to trigger animations
  await page.evaluate(async () => {
    await new Promise(resolve => {
      let total = 0;
      const dist = 200;
      const t = setInterval(() => {
        window.scrollBy(0, dist);
        total += dist;
        if (total >= document.body.scrollHeight) {
          clearInterval(t);
          window.scrollTo(0, 0);
          resolve();
        }
      }, 80);
    });
  });

  await new Promise(r => setTimeout(r, 400));
  await page.screenshot({ path: outPath, fullPage: true });
  await browser.close();
  console.log(`Screenshot saved: ${outPath}`);
})();
