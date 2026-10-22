const { chromium } = require('playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const root = path.resolve(process.argv[2]); const P = process.argv[3]; const shots = process.argv[4] === 'shots';
const srv = http.createServer((q, r) => { const f = path.join(root, decodeURIComponent(q.url.split('?')[0]));
  fs.readFile(f, (e, d) => { if (e) { r.writeHead(404); r.end(); return; } r.writeHead(200, {'Content-Type': f.endsWith('.png') ? 'image/png' : 'text/html; charset=utf-8'}); r.end(d); }); });
srv.listen(0, async () => {
  const port = srv.address().port;
  const b = await chromium.launch(); const pg = await b.newPage({ viewport: { width: 1360, height: 900 } });
  const errs = [];
  pg.on('console', m => { if (m.type() === 'error' && !/Failed to load resource|ERR_|fonts\.g/.test(m.text())) errs.push(m.text()); });
  pg.on('pageerror', e => errs.push('PAGEERROR ' + e.message));
  await pg.goto(`http://localhost:${port}/${P}.html`, { waitUntil: 'load' });
  await pg.waitForTimeout(1500);
  const tabs = await pg.$$eval('nav.tabs button[role="tab"]', bs => bs.map(b => b.dataset.p));
  const out = {};
  for (const t of tabs) {
    await pg.evaluate(t => { const b = document.querySelector(`nav.tabs button[data-p="${t}"]`); b.click(); }, t);
    await pg.waitForTimeout(250);
    out[t] = await pg.evaluate(t => { const s = document.getElementById('p-' + t); return s && !s.hidden ? s.innerText.length : -1; }, t);
    if (shots && ['eco', 'tam', 'int', 'hs', 'fic', 'con', 'cob'].includes(t)) await pg.screenshot({ path: `shots/${P}_${t}.png`, fullPage: true });
  }
  console.log(P, JSON.stringify(out));
  console.log('ERRORS', errs.length, errs.slice(0, 10).join('\n'));
  await b.close(); srv.close();
});
