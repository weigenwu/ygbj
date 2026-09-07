const assert = require('node:assert/strict');
const { chromium } = require('playwright');

(async () => {
  assert(process.argv[2], 'Provide a local pathology image path');
  const browser = await chromium.launch({ channel: 'msedge', headless: true });
  try {
    const context = await browser.newContext({ permissions: ['clipboard-read', 'clipboard-write'], viewport: { width: 1440, height: 1000 } });
    const page = await context.newPage();
    page.setDefaultTimeout(180000);
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(process.env.FUSIONMARK_URL || 'http://127.0.0.1:8765/');
    await page.selectOption('#annotationMode', 'cic');
    await page.setInputFiles('#fileInput', process.argv[2]);
    await page.waitForFunction(() => !document.querySelector('#copyButton').disabled);
    const dimensions = await page.locator('#imageCanvas').evaluate(c => [c.width, c.height]);
    await page.click('#actualSizeButton');
    assert.equal(await page.locator('#zoomOutput').textContent(), '100%');
    const bounds = await page.locator('#viewport').boundingBox();
    const x = bounds.x + bounds.width / 2, y = bounds.y + bounds.height / 2;
    await page.mouse.move(x, y);
    await page.mouse.down();
    await page.mouse.move(x + 24, y, { steps: 5 });
    await page.mouse.up();
    await page.mouse.click(x + 90, y);
    const saved = () => page.evaluate(() => JSON.parse(localStorage.getItem(Object.keys(localStorage).find(k => k.startsWith('fusionmark:v1:')))));
    let markers = await saved();
    assert.equal(markers.length, 2);
    assert(markers.every(m => m.cic && m.type === 'circle' && m.color === '#d00000'));
    assert.equal(markers[0].r, markers[1].r);
    assert.equal(await page.locator('#overlay > circle').count(), 2, 'CIC has no thick black outline');
    if (process.env.FUSIONMARK_SCREENSHOT) await page.screenshot({ path: process.env.FUSIONMARK_SCREENSHOT });
    await page.click('#undoButton');
    assert.match(await page.locator('#currentMeta').textContent(), /CIC 手动圈选 1/);
    await page.click('#redoButton');
    assert.match(await page.locator('#currentMeta').textContent(), /CIC 手动圈选 2/);
    await page.click('#copyButton');
    await page.waitForFunction(() => document.querySelector('#toast').textContent.includes('标注图已复制'));
    const png = await page.evaluate(async () => {
      const items = await navigator.clipboard.read();
      const blob = await items[0].getType('image/png');
      const header = new DataView(await blob.slice(0, 24).arrayBuffer());
      return { width: header.getUint32(16), height: header.getUint32(20), bytes: blob.size };
    });
    assert.deepEqual([png.width, png.height], dimensions);
    await page.selectOption('#annotationMode', 'fluorescence');
    assert.equal((await saved()).filter(m => m.cic).length, 2);
    await page.mouse.click(x - 90, y);
    assert.equal((await saved())[2].type, 'arrow');
    await page.reload();
    await page.setInputFiles('#fileInput', process.argv[2]);
    await page.waitForFunction(() => !document.querySelector('#copyButton').disabled);
    assert.equal((await saved()).length, 3);
    assert.match(await page.locator('#currentMeta').textContent(), /CIC 手动圈选 2/);
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({ result: 'passed', dimensions, clipboard: png, errors }));
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
