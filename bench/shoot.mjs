// Screenshot helper: node bench/shoot.mjs <url> <out.png> [actions...]
import { chromium } from 'playwright';
const [url, out, ...actions] = process.argv.slice(2);
const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium', args: ['--no-sandbox'] }).catch(() => chromium.launch());
const page = await browser.newPage({ viewport: { width: 1440, height: 860 }, deviceScaleFactor: 2, colorScheme: process.env.SCHEME || 'light' });
const logs = [];
page.on('console', (m) => logs.push(`[${m.type()}] ${m.text()}`));
page.on('pageerror', (e) => logs.push(`[pageerror] ${e.message}`));
await page.goto(url);
await page.waitForTimeout(500);
for (const a of actions) {
  const [kind, ...rest] = a.split(':');
  const arg = rest.join(':');
  if (kind === 'click') await page.getByText(arg, { exact: false }).first().click();
  else if (kind === 'key') await page.keyboard.press(arg);
  else if (kind === 'wait') await page.waitForTimeout(Number(arg));
  else if (kind === 'move') { const [x, y] = arg.split(',').map(Number); await page.mouse.move(x, y); }
  else if (kind === 'mclick') { const [x, y] = arg.split(',').map(Number); await page.mouse.click(x, y); }
  else if (kind === 'scroll') { const [dx, dy] = arg.split(',').map(Number); await page.mouse.wheel(dx, dy); }
}
await page.waitForTimeout(400);
await page.screenshot({ path: out });
console.log(logs.join('\n'));
await browser.close();
