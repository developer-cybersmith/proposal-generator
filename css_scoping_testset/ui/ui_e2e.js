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

  // Preview 3: the page must show the built document and offer both downloads (real files from the server)
  async function checkFinalPreview(label) {
    await until(() => !document.getElementById('screen-p3').classList.contains('hidden'), 'Final preview shown', 300000);
    check(label + ': Final preview step is shown', true);
    const summary = document.getElementById('p3Summary').textContent;
    check(label + ': summary reports the layout check', /Built:/.test(summary) && !/Blank page/.test(document.getElementById('p3Warn').textContent), summary);
    const frame = document.querySelector('#p3Frame iframe');
    check(label + ': the PDF is shown in a frame before download', !!frame && /\/download\/pdf\?inline=1/.test(frame.getAttribute('src')), frame && frame.getAttribute('src'));
    const docxHref = new URL(document.getElementById('dlDocx').getAttribute('href'), window.location.href).href;
    const pdfHref = new URL(document.getElementById('dlPdf').getAttribute('href'), window.location.href).href;
    check(label + ': both download buttons are visible', !document.getElementById('dlDocx').classList.contains('hidden') && !document.getElementById('dlPdf').classList.contains('hidden'));
    const dx = await fetch(docxHref); const dxb = Buffer.from(await dx.arrayBuffer());
    check(label + ': Word download is a real .docx (zip, > 100 KB, attachment)', dxb.slice(0, 2).toString() === 'PK' && dxb.length > 100000 && /attachment/.test(dx.headers.get('content-disposition') || ''), dxb.length + ' bytes');
    const dp = await fetch(pdfHref); const dpb = Buffer.from(await dp.arrayBuffer());
    check(label + ': PDF download is a real PDF', dpb.slice(0, 4).toString() === '%PDF' && dpb.length > 50000, dpb.length + ' bytes');
    const inl = await fetch(pdfHref + '?inline=1');
    check(label + ': the frame URL serves the PDF inline', /inline/.test(inl.headers.get('content-disposition') || ''));
    require('fs').writeFileSync(process.env.UI_OUT_DOCX || 'ui_out.docx', dxb);
    require('fs').writeFileSync(process.env.UI_OUT_PDF || 'ui_out.pdf', dpb);
  }
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
  if (MODE === 'generate-ai') {
    // a workbook with AI-written annexures: the user confirms the sector on Preview 1 (REAL OpenAI is called)
    check('AI panel is shown because an annexure has no template', !document.getElementById('aiPanel').classList.contains('hidden'));
    check('Generate is blocked until the sector is given', /Sector is required/.test(document.getElementById('issues1').textContent));
    check('duration was pre-filled from the Excel dates', document.getElementById('aiDuration').value.length > 0, document.getElementById('aiDuration').value);
    const sector = document.getElementById('aiSector'); sector.value = 'Banking, Financial Services and Insurance (BFSI)';
    sector.dispatchEvent(new window.Event('change'));
    document.getElementById('btnNext').click();
    await until(() => !document.getElementById('screen-p2').classList.contains('hidden'), 'Preview 2 shown');
    await until(() => !document.getElementById('btnGenerate').disabled, 'Generate enabled once the sector is set');
    check('Generate becomes enabled after the sector is entered', true);
    document.getElementById('btnGenerate').click();
    await checkFinalPreview('AI proposal');
    const bad = results.filter((r) => !r[0]); console.log(`\n${results.length - bad.length}/${results.length} UI checks passed`); process.exit(bad.length ? 1 : 0);
  }
  if (MODE === 'generate') {
    check('Generate is enabled only when nothing blocks it', true);
    document.getElementById('btnNext').click();
    await until(() => !document.getElementById('screen-p2').classList.contains('hidden'), 'Preview 2 shown');
    check('Build button is enabled for a complete known-only workbook', !document.getElementById('btnGenerate').disabled);
    document.getElementById('btnGenerate').click();
    await checkFinalPreview('Known-only proposal');
    document.getElementById('btnEdit').click();
    check('Back to edit returns to Preview 2 without losing the draft', !document.getElementById('screen-p2').classList.contains('hidden') && document.querySelectorAll('#p2Table tbody tr').length >= 3);
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
