/* CyberSmithSECURE Proposal Studio: presentation helpers shared by the sign-in page, the manual form and the Excel flow.
 *
 * This file only changes how things LOOK and what the user is TOLD. It never submits, saves, validates or changes
 * any field, and it never calls the server. The page's own script (inline in each page) is untouched and keeps doing
 * all the work; this script only reads the page and adds comfort around it:
 *   - light/dark theme toggle (remembered in this browser only)
 *   - drag-and-drop zones that show the chosen file name (the real <input type=file> is still what the page reads)
 *   - the progress rail: marks finished steps
 *   - the live summary card (reads the form, writes nothing back)
 *   - "what is happening" panel while a document is being built (an honest elapsed-time indicator, no fake progress)
 *   - show/hide password on the sign-in page
 */
(function () {
  'use strict';
  var doc = document;
  function $(id) { return doc.getElementById(id); }
  function qsa(sel, root) { return Array.prototype.slice.call((root || doc).querySelectorAll(sel)); }
  function store(key, value) { try { if (value === undefined) return window.localStorage.getItem(key); window.localStorage.setItem(key, value); } catch (e) { return null; } }

  /* ---------- theme ---------- */
  function currentTheme() { return doc.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light'; }
  function setTheme(t) {
    doc.documentElement.setAttribute('data-theme', t);
    qsa('.theme-toggle').forEach(function (b) { b.setAttribute('aria-pressed', t === 'dark' ? 'true' : 'false'); b.setAttribute('aria-label', t === 'dark' ? 'Switch to light theme' : 'Switch to dark theme'); });
  }
  setTheme(currentTheme());
  qsa('.theme-toggle').forEach(function (b) {
    b.addEventListener('click', function () { var next = currentTheme() === 'dark' ? 'light' : 'dark'; setTheme(next); store('css_theme', next); });
  });

  /* ---------- drop zones ---------- */
  function humanSize(n) { return n < 1048576 ? Math.max(1, Math.round(n / 1024)) + ' KB' : (n / 1048576).toFixed(1) + ' MB'; }
  qsa('.dropzone').forEach(function (zone) {
    var input = zone.querySelector('input[type=file]');
    var label = zone.querySelector('.dz-file');
    if (!input) return;
    function show() {
      var f = input.files && input.files[0];
      zone.classList.toggle('has-file', !!f);
      if (label) label.textContent = f ? f.name + '  ·  ' + humanSize(f.size) : '';
    }
    input.addEventListener('change', show);
    ['dragenter', 'dragover'].forEach(function (ev) { zone.addEventListener(ev, function (e) { e.preventDefault(); zone.classList.add('drag'); }); });
    ['dragleave', 'drop'].forEach(function (ev) { zone.addEventListener(ev, function () { zone.classList.remove('drag'); }); });
    zone.addEventListener('drop', function (e) {
      e.preventDefault();
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
        try { input.files = e.dataTransfer.files; input.dispatchEvent(new Event('change', { bubbles: true })); } catch (err) { /* the browser refused: the click path still works */ }
      }
    });
    show();
  });

  /* ---------- progress rail: finished steps get a tick ---------- */
  var stepEls = [0, 1, 2, 3].map(function (i) { return $('st' + i); }).filter(Boolean);
  function markSteps() {
    var current = -1;
    stepEls.forEach(function (el, i) { if (el.classList.contains('on')) current = i; });
    stepEls.forEach(function (el, i) {
      var done = current > i;
      if (el.classList.contains('done') !== done) el.classList.toggle('done', done);
      if (el.classList.contains('on')) el.setAttribute('aria-current', 'step'); else el.removeAttribute('aria-current');
    });
  }
  if (stepEls.length && window.MutationObserver) {
    var stepObserver = new MutationObserver(markSteps);
    stepEls.forEach(function (el) { stepObserver.observe(el, { attributes: true, attributeFilter: ['class'] }); });
    markSteps();
  }

  /* ---------- live summary (read only) ---------- */
  var summary = $('summaryCard');
  function setDd(id, text, empty) {
    var el = $(id); if (!el) return;
    el.textContent = text; el.classList.toggle('empty', !!empty);
  }
  function updateSummary() {
    if (!summary) return;
    var name = ($('clientName') && $('clientName').value || '').trim();
    var p1 = $('screen-p1');
    var haveDraft = !!(p1 && qsa('#screen-p1 tr[data-entry]').length);
    setDd('sumClient', name || 'Not set yet', !name);
    if (!haveDraft) {
      setDd('sumAnnexures', 'Upload a sheet first', true); setDd('sumAi', '–', true); setDd('sumDate', '–', true); setDd('sumTotal', '–', true);
      return;
    }
    var rows = qsa('#screen-p1 tr[data-entry]').filter(function (tr) { var c = tr.querySelector('.f-include'); return c && c.checked; });
    var ai = rows.filter(function (tr) { return tr.querySelector('.badge.ai'); }).length;
    setDd('sumAnnexures', String(rows.length), false);
    setDd('sumAi', ai ? ai + ' written by AI' : 'None', !ai);
    var d = ($('propDate') && $('propDate').value || '').trim();
    setDd('sumDate', d || 'Not set', !d);
    var t = $('totalText') && $('totalText').textContent.replace(/^Total:\s*/, '').trim();
    var priced = t && t.indexOf('₹ /-') !== 0;
    setDd('sumTotal', priced ? t.replace(/\s*\/-.*$/, '') : 'Not priced yet', !priced);
  }
  if (summary) {
    var scheduled = false;
    var queue = function () { if (scheduled) return; scheduled = true; setTimeout(function () { scheduled = false; updateSummary(); }, 60); };
    ['input', 'change', 'click'].forEach(function (ev) { doc.addEventListener(ev, queue, true); });
    if (window.MutationObserver) {
      var so = new MutationObserver(queue);
      ['screen-p1', 'screen-p2', 'totalText'].forEach(function (id) { if ($(id)) so.observe($(id), { childList: true, subtree: true, characterData: true }); });
    }
    updateSummary();
  }

  /* ---------- "what is happening" while a document is built ---------- */
  var panel = $('buildPanel');
  if (panel) {
    var trigger = $('btnGenerate') || $('generateBtn');
    var statusEl = $('genStatus') || $('status');
    var timer = null, started = 0;
    var elapsed = panel.querySelector('.elapsed');
    function stop() { if (timer) { clearInterval(timer); timer = null; } panel.classList.add('hidden'); }
    function start() {
      if (timer) return;
      started = Date.now(); panel.classList.remove('hidden');
      timer = setInterval(function () { if (elapsed) elapsed.textContent = Math.round((Date.now() - started) / 1000) + ' s'; }, 500);
    }
    function sync() {
      var building = trigger && trigger.disabled && statusEl && /Building|Generating|Step \d\/3/i.test(statusEl.textContent) && !/error|failed|please/i.test(statusEl.className);
      if (building) start(); else stop();
    }
    if (window.MutationObserver && trigger) {
      var bo = new MutationObserver(sync);
      bo.observe(trigger, { attributes: true, attributeFilter: ['disabled'] });
      if (statusEl) bo.observe(statusEl, { childList: true, characterData: true, subtree: true, attributes: true, attributeFilter: ['class'] });
    }
    sync();
  }

  /* ---------- sign-in: show / hide password ---------- */
  var reveal = $('pwToggle'), pw = $('pwField');
  if (reveal && pw) {
    reveal.addEventListener('click', function () {
      var show = pw.type === 'password';
      pw.type = show ? 'text' : 'password';
      reveal.textContent = show ? 'Hide' : 'Show';
      reveal.setAttribute('aria-pressed', show ? 'true' : 'false');
    });
  }
})();
