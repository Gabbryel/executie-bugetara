// Test end-to-end. Rulează: python3 -m http.server 8899 && node tests/smoke.js
// Optional: MACHETA_CUR, MACHETA_PREV, BC39 = cai catre fisiere reale (NU se commit-uiesc).
const { chromium } = require('playwright');
const fs = require('fs'); const os = require('os'); const path = require('path');
const URL = process.env.APP_URL || 'http://localhost:8899/index.html';
const { MACHETA_CUR, MACHETA_PREV, BC39 } = process.env;
let fails = 0;
const ok = (cond, msg) => { console.log((cond ? '  ✓ ' : '  ✗ ') + msg); if (!cond) fails++; };

(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ acceptDownloads: true });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push(e.message));
  page.on('console', m => { if (m.type() === 'error') errs.push(m.text()); });

  await page.goto(URL, { waitUntil: 'networkidle' });
  await page.evaluate(() => localStorage.clear());
  await page.reload({ waitUntil: 'networkidle' });
  await page.waitForTimeout(500);

  console.log('Incarcare');
  ok((await page.$$('#tbl tbody tr')).length === 121, '121 randuri in tabel');
  const checks = async () => page.$$eval('#checks .chk', els => els.map(e => ({ cls: e.className, text: e.innerText })));
  ok((await checks()).some(c => /Integritatea/.test(c.text) && /chk ok/.test(c.cls)), 'sablonul trece verificarea de integritate (I117)');

  const rand8 = async () => page.$eval('tr[data-row="8"]', tr => [...tr.querySelectorAll('td.calc')].slice(0, 6).map(t => t.textContent));

  if (MACHETA_CUR) {
    console.log('Import macheta curenta');
    await page.setInputFiles('#fCur', MACHETA_CUR); await page.waitForTimeout(700);
    const r8 = await rand8();
    console.log('    rand 8:', r8.join(' | '));
    ok(r8[0] !== '0,000', 'randul 8 are valori dupa import');
    ok(r8[0] === r8[2] && r8[0] === r8[3], 'r.8: col.3 = col.5 = col.6');
  }
  if (BC39) {
    console.log('Import BC39');
    await page.setInputFiles('#fBc', BC39); await page.waitForTimeout(500);
    ok((await checks()).some(c => /Corelația/.test(c.text) && /chk ok/.test(c.cls)), 'corelatia BC39 trece');
  }
  if (MACHETA_PREV) {
    console.log('Import luna precedenta');
    await page.setInputFiles('#fPrev', MACHETA_PREV); await page.waitForTimeout(700);
    ok(!(await page.$eval('#modeDelta', e => e.disabled)), 'modul delta s-a activat');
    ok((await checks()).some(c => /monoton/.test(c.text) && /chk ok/.test(c.cls)), 'cumulat monoton fata de luna precedenta');
  }
  if (MACHETA_CUR || MACHETA_PREV || BC39) {
    const bad = (await checks()).filter(c => /chk err/.test(c.cls));
    ok(bad.length === 0, 'nicio verificare in eroare' + (bad.length ? ': ' + bad.map(b => b.text.split('\n')[0]).join('; ') : ''));
  }

  console.log('Export .xlsx');
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 15000 }), page.click('#btnXlsx')]);
  const out = path.join(os.tmpdir(), 'export-test.xlsx'); await dl.saveAs(out);
  const buf = fs.readFileSync(out);
  ok(buf.length > 10000 && buf[0] === 0x50 && buf[1] === 0x4b, 'fisier .xlsx valid (' + buf.length + ' bytes): ' + dl.suggestedFilename());

  await browser.close();
  ok(errs.length === 0, 'fara erori in consola' + (errs.length ? ': ' + errs.join(' | ') : ''));
  console.log(fails ? `\n${fails} verificari esuate` : '\nToate verificarile au trecut');
  process.exit(fails ? 1 : 0);
})();
