const { chromium } = require('/opt/node22/lib/node_modules/playwright');
(async () => {
  const f = process.argv[2]; const shot = process.argv[3]||'';
  const b = await chromium.launch();
  const p = await b.newPage({viewport:{width:1300,height:900}});
  const errs = [];
  p.on('pageerror', e => errs.push('PAGEERR ' + e.message));
  p.on('console', m => { if (m.type()==='error') errs.push('CONSOLE ' + m.text()); });
  await p.route(/fonts\.(googleapis|gstatic)\.com/, r => r.abort());
  await p.goto('file://' + f);
  await p.waitForTimeout(800);
  const tabs = await p.$$eval('nav.tabs button[role=tab]', bs => bs.map(b => b.dataset.p));
  for (const t of tabs) {
    await p.evaluate(t => { const b=[...document.querySelectorAll('nav.tabs button[role=tab]')].find(x=>x.dataset.p===t); b.click(); }, t);
    await p.waitForTimeout(250);
    const txt = await p.$eval('#p-'+t, s => s.innerText.length);
    if (shot && shot.split(',').includes(t)) await p.screenshot({path: 'shot_'+f.split('/').pop()+'_'+t+'.png', fullPage: true});
    console.log(t, txt);
  }
  console.log('ERRORS', errs.length); errs.forEach(e=>console.log(e));
  await b.close();
})();
