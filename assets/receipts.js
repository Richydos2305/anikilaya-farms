// Anikilaya Farms — Receipts tool, ported from the entrepreneur's standalone
// build. Persistence is rewired from window.storage (a host API that only
// existed in whatever tool generated the original file) onto callApi(),
// reusing the same fetch helper and sessionPasscode already set up in
// app.js. Receipt numbers are now assigned by the server (see saveReceipt in
// Code.gs) instead of being counted client-side, so two people saving at
// once can't collide on the same number.

let itemRowId = 0;

function uid() {
  return 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function formatMoney(n) {
  n = Number(n) || 0;
  return '₦' + n.toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatDateDisplay(iso) {
  if (!iso) return '';
  const d = new Date(iso + 'T00:00:00');
  return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

/* ---------------- SETTINGS ---------------- */
let receiptSettings = { phone: '', whatsapp: '', address: '', footer: 'Thank you for your patronage!' };

async function loadReceiptSettings() {
  try {
    const res = await callApi('getReceiptSettings');
    if (res && res.ok && res.settings) {
      receiptSettings = Object.assign(receiptSettings, res.settings);
    }
  } catch (e) {
    // keep defaults — the business-details panel just stays blank/default
  }
  document.getElementById('biz-phone').value = receiptSettings.phone || '';
  document.getElementById('biz-whatsapp').value = receiptSettings.whatsapp || '';
  document.getElementById('biz-address').value = receiptSettings.address || '';
  document.getElementById('biz-footer').value = receiptSettings.footer || '';
  renderPreview();
}

async function saveReceiptSettings() {
  receiptSettings = {
    phone: document.getElementById('biz-phone').value.trim(),
    whatsapp: document.getElementById('biz-whatsapp').value.trim(),
    address: document.getElementById('biz-address').value.trim(),
    footer: document.getElementById('biz-footer').value.trim()
  };
  try {
    await callApi('saveReceiptSettings', Object.assign({ passcode: sessionPasscode }, receiptSettings));
  } catch (e) {
    console.error('Could not save business settings', e);
  }
  renderPreview();
}

['biz-phone', 'biz-whatsapp', 'biz-address', 'biz-footer'].forEach((id) => {
  document.getElementById(id).addEventListener('change', saveReceiptSettings);
});

/* ---------------- RECEIPTS LOG ---------------- */
let receipts = [];

async function loadReceipts() {
  try {
    const res = await callApi('listReceipts', { passcode: sessionPasscode });
    receipts = res && res.ok ? res.receipts : [];
  } catch (e) {
    receipts = [];
  }
}

function previewReceiptNumber_() {
  const year = new Date().getFullYear();
  const countThisYear = receipts.filter((r) => r.number && r.number.includes('-' + year + '-')).length;
  const seq = String(countThisYear + 1).padStart(4, '0');
  return `AF-${year}-${seq} (preview)`;
}

/* ---------------- ITEM ROWS ---------------- */
function addItemRow(desc, qty, price) {
  itemRowId++;
  const container = document.getElementById('items-container');
  const row = document.createElement('div');
  row.className = 'item-row';
  row.dataset.rowId = itemRowId;
  row.innerHTML = `
    <input type="text" class="item-desc" placeholder="e.g. Fresh tomatoes (basket)" value="${escapeHtml(desc || '')}">
    <input type="number" class="item-qty" min="1" step="1" value="${qty || 1}">
    <input type="number" class="item-price" min="0" step="0.01" placeholder="0.00" value="${price || ''}">
    <button type="button" class="remove" title="Remove item" onclick="removeItemRow(${itemRowId})">×</button>
  `;
  container.appendChild(row);
  row.querySelectorAll('input').forEach((inp) => inp.addEventListener('input', onFormChange));
  updateRemoveButtons();
}

function removeItemRow(rowId) {
  const container = document.getElementById('items-container');
  const row = container.querySelector(`[data-row-id="${rowId}"]`);
  if (row) row.remove();
  if (container.children.length === 0) addItemRow();
  updateRemoveButtons();
  onFormChange();
}

function updateRemoveButtons() {
  const rows = document.querySelectorAll('#items-container .item-row');
  rows.forEach((r) => {
    r.querySelector('.remove').disabled = rows.length <= 1;
  });
}

function getItems() {
  const rows = document.querySelectorAll('#items-container .item-row');
  const items = [];
  rows.forEach((r) => {
    const desc = r.querySelector('.item-desc').value.trim();
    const qty = parseFloat(r.querySelector('.item-qty').value) || 0;
    const price = parseFloat(r.querySelector('.item-price').value) || 0;
    if (desc || qty || price) {
      items.push({ desc, qty, price, lineTotal: qty * price });
    }
  });
  return items;
}

/* ---------------- TOTALS & FORM STATE ---------------- */
function computeTotals() {
  const items = getItems();
  const subtotal = items.reduce((s, i) => s + i.lineTotal, 0);
  const discount = parseFloat(document.getElementById('discount').value) || 0;
  const total = Math.max(subtotal - discount, 0);
  return { items, subtotal, discount, total };
}

function onFormChange() {
  const { total } = computeTotals();
  const paidField = document.getElementById('amount-paid');
  if (paidField.dataset.touched !== 'true') {
    paidField.value = total.toFixed(2);
  }
  renderTotalsBox();
  renderPreview();
}

document.getElementById('amount-paid').addEventListener('input', () => {
  document.getElementById('amount-paid').dataset.touched = 'true';
  renderTotalsBox();
  renderPreview();
});
['discount', 'cust-name', 'cust-phone', 'r-date', 'payment-method', 'notes'].forEach((id) => {
  document.getElementById(id).addEventListener('input', () => {
    renderTotalsBox();
    renderPreview();
  });
});

function renderTotalsBox() {
  const { subtotal, discount, total } = computeTotals();
  const paid = parseFloat(document.getElementById('amount-paid').value) || 0;
  const balance = total - paid;
  document.getElementById('totals-box').innerHTML = `
    <div class="trow"><span>Subtotal</span><span>${formatMoney(subtotal)}</span></div>
    <div class="trow"><span>Discount</span><span>-${formatMoney(discount)}</span></div>
    <div class="trow grand"><span>Total</span><span>${formatMoney(total)}</span></div>
    <div class="trow"><span>Amount paid</span><span>${formatMoney(paid)}</span></div>
    <div class="trow" style="${balance > 0.004 ? 'color:#9A3324;font-weight:700;' : 'color:#2F5738;font-weight:700;'}"><span>${balance > 0.004 ? 'Balance due' : 'Status'}</span><span>${balance > 0.004 ? formatMoney(balance) : 'Fully paid'}</span></div>
  `;
}

/* ---------------- PREVIEW ---------------- */
function renderPreview(receiptOverride) {
  const r = receiptOverride || getCurrentFormReceipt();
  const balance = r.total - r.amountPaid;
  const itemRows = r.items.map((i) => `
    <tr>
      <td>${escapeHtml(i.desc || '—')}</td>
      <td class="num">${i.qty}</td>
      <td class="num">${formatMoney(i.price)}</td>
      <td class="num">${formatMoney(i.lineTotal)}</td>
    </tr>`).join('');

  document.getElementById('preview-inner').innerHTML = `
    <div class="r-number">${escapeHtml(r.number)}</div>
    <div class="r-head">
      <div class="fname">Anikilaya Farms</div>
      <div class="faddr">
        ${receiptSettings.address ? escapeHtml(receiptSettings.address) + '<br>' : ''}
        ${receiptSettings.phone ? 'Tel: ' + escapeHtml(receiptSettings.phone) : ''}
      </div>
    </div>
    <div class="r-meta">
      <span>Date: ${escapeHtml(formatDateDisplay(r.date))}</span>
      <span>${escapeHtml(r.paymentMethod)}</span>
    </div>
    <div class="r-cust">Received from: <b>${escapeHtml(r.customerName || '—')}</b></div>
    <table class="r-items">
      <thead><tr><th>Item</th><th class="num">Qty</th><th class="num">Price</th><th class="num">Amount</th></tr></thead>
      <tbody>${itemRows || '<tr><td colspan="4" style="color:#9a9d84;">No items added yet</td></tr>'}</tbody>
    </table>
    <div class="r-totals">
      <div class="trow"><span>Subtotal</span><span>${formatMoney(r.subtotal)}</span></div>
      <div class="trow"><span>Discount</span><span>-${formatMoney(r.discount)}</span></div>
      <div class="trow grand"><span>Total</span><span>${formatMoney(r.total)}</span></div>
      <div class="trow"><span>Amount paid</span><span>${formatMoney(r.amountPaid)}</span></div>
      ${balance > 0.004 ? `<div class="trow balance"><span>Balance due</span><span>${formatMoney(balance)}</span></div>` : ''}
    </div>
    ${r.notes ? `<div class="r-foot" style="font-style:normal;text-align:left;margin-top:10px;">Note: ${escapeHtml(r.notes)}</div>` : ''}
    <div class="r-foot">${escapeHtml(receiptSettings.footer || 'Thank you for your patronage!')}</div>
    ${balance <= 0.004 ? `<div class="stamp"><b>PAID</b><small>ANIKILAYA FARMS</small></div>` : ''}
  `;
}

function getCurrentFormReceipt() {
  const { items, subtotal, discount, total } = computeTotals();
  const amountPaid = parseFloat(document.getElementById('amount-paid').value) || 0;
  return {
    number: previewReceiptNumber_(),
    date: document.getElementById('r-date').value || new Date().toISOString().slice(0, 10),
    customerName: document.getElementById('cust-name').value.trim(),
    customerPhone: document.getElementById('cust-phone').value.trim(),
    items, subtotal, discount, total, amountPaid,
    paymentMethod: document.getElementById('payment-method').value,
    notes: document.getElementById('notes').value.trim()
  };
}

/* ---------------- VALIDATION ---------------- */
function validateForm(r) {
  if (!r.customerName) return 'Add the customer\'s name before saving.';
  if (!r.items.length || r.items.every((i) => !i.desc && !i.qty && !i.price)) return 'Add at least one item.';
  return null;
}

/* ---------------- SAVE + SHARE ---------------- */
async function saveAndShare(shareToWhatsapp) {
  const r = getCurrentFormReceipt();
  const err = validateForm(r);
  const hint = document.getElementById('validation-hint');
  if (err) { hint.textContent = err; hint.style.color = '#9A3324'; return; }
  hint.textContent = '';

  const actionButtons = document.querySelectorAll('#new-receipt-actions button');
  actionButtons.forEach((b) => (b.disabled = true));

  let result;
  try {
    result = await callApi('saveReceipt', {
      passcode: sessionPasscode,
      date: r.date,
      customerName: r.customerName,
      customerPhone: r.customerPhone,
      items: r.items,
      subtotal: r.subtotal,
      discount: r.discount,
      total: r.total,
      amountPaid: r.amountPaid,
      paymentMethod: r.paymentMethod,
      notes: r.notes
    });
  } catch (e) {
    actionButtons.forEach((b) => (b.disabled = false));
    hint.textContent = 'Something went wrong. Check your connection and try again.';
    hint.style.color = '#9A3324';
    return;
  }

  actionButtons.forEach((b) => (b.disabled = false));

  if (!result.ok) {
    hint.textContent = 'Something went wrong. Please try again.';
    hint.style.color = '#9A3324';
    return;
  }

  const saved = result.receipt;
  receipts.unshift(saved);

  renderPreview(saved);
  renderLog();

  if (shareToWhatsapp) {
    shareReceiptOnWhatsapp(saved);
  } else {
    hint.textContent = `Saved as ${saved.number}.`;
    hint.style.color = '#2F5738';
  }

  resetForm();
}

function buildWhatsappMessage(r) {
  const balance = r.total - r.amountPaid;
  const lines = [];
  lines.push(`*Anikilaya Farms — Payment Receipt*`);
  lines.push(`Receipt No: ${r.number}`);
  lines.push(`Date: ${formatDateDisplay(r.date)}`);
  lines.push('');
  lines.push(`Customer: ${r.customerName}`);
  lines.push('');
  lines.push('Items:');
  r.items.forEach((i, idx) => {
    lines.push(`${idx + 1}. ${i.desc || 'Item'} x${i.qty} — ${formatMoney(i.lineTotal)}`);
  });
  lines.push('');
  lines.push(`Subtotal: ${formatMoney(r.subtotal)}`);
  if (r.discount) lines.push(`Discount: ${formatMoney(r.discount)}`);
  lines.push(`Total: ${formatMoney(r.total)}`);
  lines.push(`Amount Paid: ${formatMoney(r.amountPaid)}`);
  if (balance > 0.004) lines.push(`Balance Due: ${formatMoney(balance)}`);
  lines.push(`Payment Method: ${r.paymentMethod}`);
  if (r.notes) lines.push(`Note: ${r.notes}`);
  lines.push('');
  lines.push(receiptSettings.footer || 'Thank you for your patronage!');
  return lines.join('\n');
}

function sanitizePhoneForWa(phone) {
  if (!phone) return '';
  let digits = phone.replace(/[^\d]/g, '');
  if (digits.startsWith('0')) digits = '234' + digits.slice(1);
  return digits;
}

function shareReceiptOnWhatsapp(r) {
  const text = buildWhatsappMessage(r);
  const target = sanitizePhoneForWa(r.customerPhone);
  const url = target
    ? `https://wa.me/${target}?text=${encodeURIComponent(text)}`
    : `https://wa.me/?text=${encodeURIComponent(text)}`;
  window.open(url, '_blank');
}

function printReceipt() {
  const r = getCurrentFormReceipt();
  const err = validateForm(r);
  if (err) {
    document.getElementById('validation-hint').textContent = err;
    document.getElementById('validation-hint').style.color = '#9A3324';
    return;
  }
  window.print();
}

function resetForm() {
  document.getElementById('cust-name').value = '';
  document.getElementById('cust-phone').value = '';
  document.getElementById('notes').value = '';
  document.getElementById('discount').value = '0';
  document.getElementById('amount-paid').dataset.touched = 'false';
  document.getElementById('items-container').innerHTML = '';
  addItemRow();
  document.getElementById('r-date').value = new Date().toISOString().slice(0, 10);
  onFormChange();
}

/* ---------------- LOG VIEW ---------------- */
function renderLog() {
  const search = (document.getElementById('log-search-input').value || '').toLowerCase();
  const filtered = receipts.filter((r) =>
    !search ||
    (r.customerName || '').toLowerCase().includes(search) ||
    (r.number || '').toLowerCase().includes(search)
  );

  const now = new Date();
  const thisMonthKey = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');
  const totalCollected = receipts.reduce((s, r) => s + (r.amountPaid || 0), 0);
  const monthCollected = receipts
    .filter((r) => (r.date || '').startsWith(thisMonthKey))
    .reduce((s, r) => s + (r.amountPaid || 0), 0);

  document.getElementById('stat-cards').innerHTML = `
    <div class="stat-card"><div class="label">Total receipts</div><div class="value">${receipts.length}</div></div>
    <div class="stat-card"><div class="label">Total collected</div><div class="value">${formatMoney(totalCollected)}</div></div>
    <div class="stat-card"><div class="label">This month</div><div class="value">${formatMoney(monthCollected)}</div></div>
  `;

  const listEl = document.getElementById('log-list');
  if (filtered.length === 0) {
    listEl.innerHTML = `<div class="empty-state"><div class="glyph">🌾</div>No receipts found yet. Issue your first receipt from the "New Receipt" tab.</div>`;
    return;
  }

  listEl.innerHTML = filtered.map((r) => {
    const balance = r.total - r.amountPaid;
    const paid = balance <= 0.004;
    return `
    <div class="log-item">
      <div class="log-row" onclick="toggleLogDetail('${r.id}')">
        <div>
          <div class="log-cust">${escapeHtml(r.customerName)}</div>
          <div class="log-num">${escapeHtml(r.number)} · ${escapeHtml(formatDateDisplay(r.date))}</div>
        </div>
        <div style="display:flex;align-items:center;gap:10px;">
          <span class="badge ${paid ? 'paid' : 'partial'}">${paid ? 'Paid' : 'Partial'}</span>
          <span class="log-amt">${formatMoney(r.total)}</span>
        </div>
      </div>
      <div class="log-detail" id="detail-${r.id}">
        ${renderLogDetailContent(r)}
      </div>
    </div>`;
  }).join('');
}

function renderLogDetailContent(r) {
  const balance = r.total - r.amountPaid;
  const itemsList = r.items.map((i) => `<div style="font-size:13px;">• ${escapeHtml(i.desc || 'Item')} x${i.qty} — ${formatMoney(i.lineTotal)}</div>`).join('');
  return `
    <div style="padding-top:10px;font-size:13.5px;">
      ${itemsList}
      <div style="margin-top:8px;font-family:var(--font-ledger);">
        Total: ${formatMoney(r.total)} · Paid: ${formatMoney(r.amountPaid)} ${balance > 0.004 ? '· Balance: ' + formatMoney(balance) : ''}
      </div>
      <div style="margin-top:4px;color:var(--soil);">Payment method: ${escapeHtml(r.paymentMethod)}</div>
      ${r.notes ? `<div style="margin-top:4px;color:var(--soil);">Note: ${escapeHtml(r.notes)}</div>` : ''}
    </div>
    <div class="actions">
      <button class="secondary" onclick='resendWhatsapp(${JSON.stringify(r.id)})'>Share on WhatsApp</button>
      <button class="ghost" onclick='deleteReceipt(${JSON.stringify(r.id)})'>Delete</button>
    </div>
  `;
}

function toggleLogDetail(id) {
  const el = document.getElementById('detail-' + id);
  document.querySelectorAll('.log-detail.open').forEach((d) => { if (d !== el) d.classList.remove('open'); });
  el.classList.toggle('open');
}

function resendWhatsapp(id) {
  const r = receipts.find((x) => x.id === id);
  if (r) shareReceiptOnWhatsapp(r);
}

async function deleteReceipt(id) {
  if (!confirm('Delete this receipt permanently from your records?')) return;
  try {
    await callApi('deleteReceipt', { passcode: sessionPasscode, number: id });
  } catch (e) {
    alert('Could not delete receipt. Check your connection and try again.');
    return;
  }
  receipts = receipts.filter((r) => r.id !== id);
  renderLog();
}

/* ---------------- TABS ---------------- */
function switchTab(tab) {
  document.getElementById('tab-new').classList.toggle('active', tab === 'new');
  document.getElementById('tab-log').classList.toggle('active', tab === 'log');
  document.getElementById('view-new').classList.toggle('active', tab === 'new');
  document.getElementById('view-log').classList.toggle('active', tab === 'log');
  if (tab === 'log') renderLog();
}

/* ---------------- ENTRY POINT ---------------- */
// Called from the chooser's "Receipts" button (see app.js) each time the
// view is opened — re-fetches settings + the receipt log from the server so
// the shared, multi-device data stays fresh, but leaves any in-progress New
// Receipt form fields alone rather than wiping them on every revisit.
async function initReceiptsApp() {
  if (!document.getElementById('r-date').value) {
    document.getElementById('r-date').value = new Date().toISOString().slice(0, 10);
  }
  if (document.getElementById('items-container').children.length === 0) {
    addItemRow();
  }
  await loadReceiptSettings();
  await loadReceipts();
  onFormChange();
  renderLog();
}
window.initReceiptsApp = initReceiptsApp;
