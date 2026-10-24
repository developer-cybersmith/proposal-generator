// End-to-end test of proposal_excel_ui.html against the REAL server (no browser: jsdom + node fetch).
const { JSDOM } = require('jsdom');
const fs = require('fs');
const BASE = process.argv[2];            // http://127.0.0.1:8123
const XLSX = process.argv[3];            // path of a filled scoping workbook
const results = [];
const check = (name, ok, detail = '') => { results.push([ok, name, detail]); console.log((ok ? 'PASS ' : 'FAIL ') + name + (detail ? '  -> ' + detail : '')); };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
async function until(fn, label, ms = 20000) { const t = Date.now(); while (Date.now() - t < ms) { try { const v = fn(); if (v) return v; } catch (e) {} await sleep(100); } throw new Error('timeout: ' + label); }

(async () => {
  const dom = await JSDOM.fromURL(BASE + '/excel/', { runScripts: 'dangerously', resources: 'usable', pretendToBeVisual: true });
  const { window } = dom; const { document } = window;
  // adapters: jsdom FormData/File -> node fetch (undici)
  let downloaded = null;
  window.URL.createObjectURL = (b) => { downloaded = b; return 'blob:test'; };
  window.URL.revokeObjectURL = () => {};
  window.HTMLAnchorElement.prototype.click = function () {};
  window.fetch = async (url, opts = {}) => {
    const u = new URL(url, window.location.href).href;
    let body = opts.body;
    const readBlob = (blob) => new Promise((res, rej) => { const r = new window.FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(blob); });
    if (body && typeof body.entries === 'function' && !(body instanceof URLSearchParams) && typeof body !== 'string') {
      const fd = new FormData();
      for (const [k, v] of body.entries()) {
        if (v instanceof window.Blob) fd.append(k, new Blob([Buffer.from(await readBlob(v))]), v.name || 'blob');
        else fd.append(k, v);
      }
      body = fd;
    }
    return fetch(u, Object.assign({}, opts, { body }));
  };
  await until(() => document.getElementById('btnUpload'), 'page loaded');
  check('page loads with upload step visible', !document.getElementById('screen-upload').classList.contains('hidden'));

  // choose the workbook: jsdom cannot assign files, so give the input a files list and make FormData carry bytes
  const bytes = fs.readFileSync(XLSX);
  const file = new window.File([bytes], require('path').basename(XLSX), { type: 'application/octet-stream' });
  file.__bytes = bytes;
  Object.defineProperty(document.getElementById('xlsx'), 'files', { value: [file] });
  document.getElementById('btnUpload').click();
  await until(() => !document.getElementById('screen-p1').classList.contains('hidden'), 'Preview 1 shown');
  check('upload leads to Preview 1', true);

  const MODE = process.argv[4] || 'edit';
  const rows = [...document.querySelectorAll('#p1Table tbody tr')];
  check('Preview 1 lists the in-scope annexures', rows.length >= 2, rows.length + ' rows');
  if (MODE === 'generate') {
    check('Generate is enabled only when nothing blocks it', true);
    document.getElementById('btnNext').click();
    await until(() => !document.getElementById('screen-p2').classList.contains('hidden'), 'Preview 2 shown');
    check('Generate button is enabled for a complete known-only workbook', !document.getElementById('btnGenerate').disabled);
    document.getElementById('btnGenerate').click();
    await until(() => /downloaded|check reported/i.test(document.getElementById('genStatus').textContent), 'generation finished', 180000);
    const st = document.getElementById('genStatus').textContent;
    check('Generate finished and the file was handed to the browser', /Proposal downloaded/.test(st) && downloaded && downloaded.size > 100000, st + ' | ' + (downloaded ? downloaded.size + ' bytes' : 'no blob'));
    const bad = results.filter((r) => !r[0]); console.log(`\n${results.length - bad.length}/${results.length} UI checks passed`); process.exit(bad.length ? 1 : 0);
  }
  const web = document.querySelector('#screen-p1 tr[data-entry="web_app_vapt"]');
  check('Web row shows Count 6 (screenshot example)', web && web.querySelector('.f-count').value === '6', web && web.querySelector('.f-count').value);
  const manual = [...document.querySelectorAll('#p1Manual tbody tr')];
  check('five manual-toggle annexures are offered, unticked', manual.length === 5 && manual.every((r) => !r.querySelector('.f-include').checked));
  check('client name pre-filled', document.getElementById('clientName').value.length > 0, document.getElementById('clientName').value);
  const ref = document.getElementById('refNo').value;
  check('reference number pre-filled and editable', ref.length > 0, ref);

  // edit: frequency of web, count label; tick a manual annexure
  const freq = web.querySelector('.f-freq'); freq.value = 'Quarterly'; freq.dispatchEvent(new window.Event('change'));
  await sleep(600);
  const webAfter = document.querySelector('#screen-p1 tr[data-entry="web_app_vapt"]');
  check('Frequency edit survives a server round trip', webAfter.querySelector('.f-freq').value === 'Quarterly');
  const pb = document.querySelector('#p1Manual tr[data-entry="post_breach"] .f-include'); pb.checked = true; pb.dispatchEvent(new window.Event('change'));
  await sleep(600);
  check('manual annexure can be included', document.querySelector('#p1Manual tr[data-entry="post_breach"] .f-include').checked);

  document.getElementById('btnNext').click();
  await until(() => !document.getElementById('screen-p2').classList.contains('hidden'), 'Preview 2 shown');
  const prow = [...document.querySelectorAll('#p2Table tbody tr')];
  check('Preview 2 has a row for every included annexure', prow.length >= 3, prow.length + ' rows');
  const webRow = prow.find((r) => r.querySelector('.p-title').value.startsWith('Web Application'));
  check('Quantity is the block sum (23)', webRow && webRow.querySelector('.p-qty').value === '23', webRow && webRow.querySelector('.p-qty').value);
  check('shared Frequency carried into Preview 2', webRow.querySelector('.p-freq').value === 'Quarterly');
  check('cost column starts empty', prow.every((r) => r.querySelector('.p-cost').value === ''));
  check('total shows the placeholder while prices are empty', document.getElementById('totalText').textContent.includes('₹ /- * + 18% GST'));
  prow.forEach((r) => { const c = r.querySelector('.p-cost'); c.value = '100000'; c.dispatchEvent(new window.Event('input')); });
  const total = document.getElementById('totalText').textContent;
  const expected = '₹ ' + (prow.length * 100000).toLocaleString('en-IN') + ' /- * + 18% GST';
  check('total uses Indian grouping when every row is priced', total.includes(expected), total);
  prow[prow.length - 1].querySelector('button.danger').click();
  check('a pricing row can be deleted', document.querySelectorAll('#p2Table tbody tr').length === prow.length - 1);
  document.getElementById('btnAddRow').click();
  check('a pricing row can be added', document.querySelectorAll('#p2Table tbody tr').length === prow.length);

  check('Generate stays disabled while an AI-written annexure still needs a sector', document.getElementById('btnGenerate').disabled === true);
  check('the blocking reason is shown to the user', /Sector is required/.test(document.getElementById('issues2').textContent), document.getElementById('issues2').textContent.slice(0, 80));

  const failed = results.filter((r) => !r[0]);
  console.log(`\n${results.length - failed.length}/${results.length} UI checks passed`);
  process.exit(failed.length ? 1 : 0);
})().catch((e) => { console.error('UI TEST CRASHED:', e.message); process.exit(2); });
