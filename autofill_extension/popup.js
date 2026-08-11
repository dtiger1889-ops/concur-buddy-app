// Popup: loads Concur Buddy's export into chrome.storage.local, lists the staged expenses,
// and injects the fill function DIRECTLY INTO THE PAGE'S MAIN WORLD on demand
// (chrome.scripting.executeScript, world:'MAIN').
//
// Why main-world injection instead of a content script: content scripts run in Chrome's
// ISOLATED world and aren't injected into tabs that were open before the extension loaded —
// both are silent-failure modes against Concur's React app. Main-world on-demand injection is
// byte-for-byte the mechanism verified live against a real Concur expense form.

const fileInput = document.getElementById('file');
const listDiv = document.getElementById('list');
const statusDiv = document.getElementById('status');

function showStatus(text) { statusDiv.style.display = 'block'; statusDiv.textContent = text; }

// ---- This function is serialized and executed INSIDE the Concur page (main world). ----
// It must be fully self-contained: no closures, no chrome.* APIs; cfg carries the selector map.
// Async: combobox filling has to wait for Concur's dropdowns to render; executeScript awaits it.
async function pageFill(expense, cfg) {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  function setVal(el, val) {
    const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, val);
    el.dispatchEvent(new Event('input',  { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }
  function press(el) {  // React widgets often open on pointer/mouse down, not just click
    for (const t of ['pointerdown', 'mousedown', 'mouseup', 'click'])
      el.dispatchEvent(new MouseEvent(t, { bubbles: true, cancelable: true, view: window }));
  }
  function comboRoot(spec) {
    if (spec.sel) return document.querySelector(spec.sel);
    if (spec.label) {  // fallback: find the visible label, then the nearest combobox around it
      const labs = [...document.querySelectorAll('label,span,div')]
        .filter(e => e.childElementCount <= 1 && e.textContent.trim().replace(/\s*\*\s*$/, '') === spec.label && (e.offsetWidth || e.offsetHeight));
      for (const l of labs) {
        let scope = l.parentElement;
        for (let i = 0; i < 5 && scope; i++) {
          const c = scope.querySelector('[role="combobox"],input[role="combobox"]');
          if (c) return c;
          scope = scope.parentElement;
        }
      }
    }
    return null;
  }
  async function fillCombo(spec, wantText, matchFn) {
    const root = comboRoot(spec);
    if (!root) return 'not found on page';
    const opener = root.matches('input,[role="combobox"]') ? root : (root.querySelector('input,[role="combobox"]') || root);
    press(opener);
    const match = matchFn || (o => o.toLowerCase().includes(String(wantText).toLowerCase()));
    const deadline = Date.now() + 4000;
    let typed = false, sawOptions = false;
    while (Date.now() < deadline) {
      await sleep(200);
      if (!typed) {  // dropdowns open with a search box (sometimes the opener itself is it)
        const search = [...document.querySelectorAll('input')].filter(i =>
          (i.offsetWidth || i.offsetHeight) && !i.readOnly && !i.disabled &&
          (i === opener || /search/i.test(i.placeholder || '') || i.getAttribute('aria-autocomplete')));
        if (search.length) { setVal(search[0], String(wantText)); typed = true; await sleep(400); }
      }
      const opts = [...document.querySelectorAll('[role="option"],[role="listbox"] li')]
        .filter(o => o.offsetWidth || o.offsetHeight);
      if (opts.length) sawOptions = true;
      const hit = opts.find(o => match(o.textContent.trim()));
      if (hit) { press(hit); await sleep(300); return 'set to "' + hit.textContent.trim().slice(0, 50) + '"'; }
    }
    return sawOptions ? 'dropdown opened but no option matched "' + wantText + '"' : 'dropdown did not open';
  }

  if (!document.querySelector(cfg.SENTINEL)) {
    return { ok: false, error: 'No expense form here. Open the expense (or Add Expense + pick a type) — then Fill.' };
  }
  const combos = {};
  // Expense type FIRST — changing it re-renders the whole form, which would wipe later fills.
  if (expense.expense_type_code) {
    const code = expense.expense_type_code;
    const already = (comboRoot(cfg.COMBOS.expenseType) || {}).textContent || '';
    combos.expenseType = already.includes(code)
      ? 'already ' + code
      : await fillCombo(cfg.COMBOS.expenseType, code, t => t.startsWith(code + '-') || t.includes(code));
    await sleep(800);  // let the per-type form settle
  }

  const filled = [], locked = [], missing = [], emptyInExport = [];
  for (const [key, sel] of Object.entries(cfg.TEXT)) {
    const val = (expense.concur_fields || {})[key];
    if (val === undefined || val === '') { emptyInExport.push(key); continue; }
    const el = document.querySelector(sel);
    if (!el) { missing.push(key); continue; }
    if (el.readOnly || el.disabled) { locked.push(key); continue; }  // card-fed fields stay Concur's
    setVal(el, val); filled.push(key);
  }
  for (const [key, sel] of Object.entries(cfg.CHECK)) {
    const want = (expense.concur_checkboxes || {})[key];
    if (want === undefined) continue;
    const el = document.querySelector(sel);
    if (!el) { missing.push(key); continue; }
    if (el.disabled) { locked.push(key); continue; }
    if (el.checked !== !!want) el.click();
    filled.push(key);
  }

  const mf = expense.manual_fields || {};
  // "Is this a Vendor Invoice?" — Concur Buddy stages it as a boolean; Concur wants Yes/No.
  combos.vendorInvoice = await fillCombo(cfg.COMBOS.custom17, mf.is_vendor_invoice ? 'Yes' : 'No',
    t => t.toLowerCase() === (mf.is_vendor_invoice ? 'yes' : 'no'));
  if (mf.city) combos.city = await fillCombo(cfg.COMBOS.location, mf.city);

  let receipt = 'none in export';
  if (expense.receipt && expense.receipt.b64) {
    const inp = document.querySelector(cfg.FILE);
    if (!inp) receipt = 'no upload input found on this page';
    else {
      try {
        const bin = atob(expense.receipt.b64); const bytes = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        const ext = expense.receipt.name.toLowerCase().split('.').pop();
        const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', pdf: 'application/pdf',
                       tif: 'image/tiff', tiff: 'image/tiff' }[ext] || 'application/octet-stream';
        const dt = new DataTransfer();
        dt.items.add(new File([bytes], expense.receipt.name, { type: mime }));
        inp.files = dt.files;
        inp.dispatchEvent(new Event('change', { bubbles: true }));
        receipt = 'upload started (' + expense.receipt.name + ') — confirm it renders in the viewer';
      } catch (e) { receipt = 'attach failed: ' + e.message; }
    }
  }
  return { ok: true, filled, locked, missing, emptyInExport, combos, receipt,
           manual: { payment_type: mf.payment_type, invoice_number: mf.invoice_number, attendees: mf.attendees } };
}

function render(data) {
  listDiv.replaceChildren();
  if (!data || !data.expenses || !data.expenses.length) return;
  for (const exp of data.expenses) {
    const row = document.createElement('div'); row.className = 'exp';
    const who = document.createElement('div'); who.className = 'who';
    const b = document.createElement('b'); b.textContent = exp.vendor || '(no vendor)';
    const small = document.createElement('small');
    small.textContent = `${exp.concur_fields.transactionDate || ''}  $${exp.concur_fields.transactionAmount || ''}  ${exp.expense_type_code || ''} ${exp.expense_type_label || ''}${exp.receipt ? '  📎' : ''}`;
    who.append(b, small);
    const btn = document.createElement('button'); btn.textContent = 'Fill';
    btn.addEventListener('click', () => fill(exp));
    row.append(who, btn);
    listDiv.append(row);
  }
}

async function fill(exp) {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !/^https:\/\/us2\.concursolutions\.com\//.test(tab.url || '')) {
    showStatus('Switch to your Concur tab (us2.concursolutions.com) first.'); return;
  }
  const cfg = { TEXT: CB_TEXT_FIELDS, CHECK: CB_CHECKBOXES, SENTINEL: CB_FORM_SENTINEL, FILE: CB_RECEIPT_INPUT, COMBOS: CB_COMBOS };
  showStatus('Filling…');
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      world: 'MAIN',
      func: pageFill,
      args: [exp, cfg]
    });
    const res = results && results[0] && results[0].result;
    if (!res) { showStatus('The page gave no answer — reload the Concur tab and try again.'); return; }
    if (!res.ok) { showStatus(res.error || 'Could not fill.'); return; }
    let text = `Filled: ${res.filled.join(', ') || '(nothing)'}`;
    for (const [k, v] of Object.entries(res.combos || {})) text += `\n${k}: ${v}`;
    if (res.locked.length) text += `\nLocked by Concur (card data, left alone): ${res.locked.join(', ')}`;
    if (res.emptyInExport && res.emptyInExport.length) text += `\nEmpty in Concur Buddy (nothing to fill): ${res.emptyInExport.join(', ')}`;
    if (res.missing.length) text += `\nNot found on page (selector drift? see remap runbook): ${res.missing.join(', ')}`;
    if (res.receipt && res.receipt !== 'none in export') text += `\nReceipt: ${res.receipt}`;
    const manual = Object.entries(res.manual || {}).filter(([, v]) => v !== '' && v !== false && v !== undefined);
    if (manual.length) text += '\nStill yours to set: ' + manual.map(([k, v]) => `${k}=${v}`).join(', ');
    text += '\nReview the form, then click Save Expense yourself.';
    showStatus(text);
  } catch (e) {
    showStatus('Could not inject into the page (' + e.message + '). Is this the Concur tab? Try reloading it.');
  }
}

fileInput.addEventListener('change', () => {
  const f = fileInput.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      chrome.storage.local.set({ staged: data });
      render(data);
      showStatus(`Loaded ${data.expenses.length} staged expense(s).`);
    } catch (e) { showStatus('Not a valid export file: ' + e.message); }
  };
  reader.readAsText(f);
});

document.getElementById('ver').textContent = 'v' + chrome.runtime.getManifest().version;
chrome.storage.local.get('staged').then(({ staged }) => render(staged));
