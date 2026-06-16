import puppeteer from 'puppeteer';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const html = 'file://' + join(__dirname, 'index.html');
const out = join(__dirname, 'MAWANG_NIGHT_기획보고.pdf');

const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--font-render-hinting=none'] });
const page = await browser.newPage();
await page.goto(html, { waitUntil: 'networkidle0' });
await page.evaluate(async () => { await document.fonts.ready; });
await page.pdf({ path: out, printBackground: true, preferCSSPageSize: true });
await browser.close();
console.log('PDF written:', out);
