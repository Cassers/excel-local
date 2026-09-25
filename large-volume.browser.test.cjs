const { chromium } = require('playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
(async () => {
  const browser = await chromium.launch({ executablePath: '/usr/bin/chromium', headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
    const errors = [], dialogs = [], results = [];
    page.on('pageerror', e => errors.push(e.message)); page.on('dialog', d => { dialogs.push(d.message()); d.accept(); });
    await page.goto('http://127.0.0.1:8768/');
    const cdp = await page.context().newCDPSession(page); await cdp.send('Performance.enable');
    await page.evaluate(() => { window.uiTicks = 0; setInterval(() => window.uiTicks++, 100); });
    for (const filename of ['synthetic-100MB.csv', 'synthetic-500MB.xlsx']) {
      const file = path.join(__dirname, '.test-output/large', filename), start = Date.now();
      const ticks = await page.evaluate(() => window.uiTicks);
      await page.locator('#file-input').setInputFiles(file);
      await page.locator('#large-editor').waitFor({ state: 'visible', timeout: 180000 });
      await page.locator('#large-busy').waitFor({ state: 'hidden', timeout: 180000 });
      const seconds = (Date.now() - start) / 1000;
      assert.ok(await page.evaluate(() => window.uiTicks) - ticks > seconds * 4, 'UI timer must keep running during import');
      assert.ok(await page.locator('#large-grid td').count() <= 2400);
      await page.locator('[data-large-cell="A1"]').click(); await page.locator('#large-value').fill('Edición de volumen verificada'); await page.locator('#large-save').click();
      await page.waitForFunction(() => document.querySelector('[data-large-cell="A1"]')?.textContent === 'Edición de volumen verificada');
      const metrics = await cdp.send('Performance.getMetrics'); const heapMB = metrics.metrics.find(m => m.name === 'JSHeapUsedSize').value / 1024 / 1024;
      const waiting = page.waitForEvent('download', { timeout: 180000 }); await page.locator('#large-csv').click(); const download = await waiting;
      const output = path.join(__dirname, '.test-output/large/browser-export.csv'); await download.saveAs(output);
      const fd = fs.openSync(output, 'r'), prefix = Buffer.alloc(128); fs.readSync(fd, prefix, 0, 128, 0); fs.closeSync(fd);
      assert.ok(prefix.toString('utf8').includes('Edición de volumen verificada')); fs.unlinkSync(output);
      await page.locator('#large-close').click(); await page.locator('#large-editor').waitFor({ state: 'hidden' });
      results.push({ filename, bytes: fs.statSync(file).size, import_seconds: seconds, browser_js_heap_mb: +heapMB.toFixed(1) });
    }
    // Cancellation is tested against a real large file, not an artificial delay.
    await page.locator('#file-input').setInputFiles(path.join(__dirname, '.test-output/large/synthetic-500MB.xlsx'));
    await page.waitForFunction(() => /Leyendo|filas|textos/.test(document.getElementById('large-progress').textContent), { timeout: 30000 });
    await page.locator('#large-cancel').click(); await page.locator('#large-busy').waitFor({ state: 'hidden', timeout: 30000 });
    assert.ok(dialogs.includes('Importación cancelada.'));
    assert.deepEqual(errors, []);
    const result = { results, cancellation: 'passed', browser_errors: errors };
    fs.writeFileSync(path.join(__dirname, '.test-output/large/browser-results.json'), JSON.stringify(result, null, 2));
    console.log(JSON.stringify(result));
  } finally { await browser.close(); }
})().catch(e => { console.error(e); process.exitCode = 1; });
