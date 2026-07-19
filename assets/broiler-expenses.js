// Broiler Expenses — admin-only ledger. The chooser button is hidden for
// staff and the backend independently enforces isAdminAuthorized_, so no
// extra role check is needed inside this file.

let broilerExpensesEntries = [];

const broilerExpensesForm = document.getElementById('broiler-expenses-form');
const broilerExpensesStatus = document.getElementById('broiler-expenses-status');
const broilerExpensesError = document.getElementById('broiler-expenses-error');
const broilerExpensesSubmitBtn = broilerExpensesForm.querySelector('button[type="submit"]');
const broilerExpensesLogEl = document.getElementById('broiler-expenses-log');

const beDateInput = document.getElementById('be-date');
const beQtyInput = document.getElementById('be-qty');
const bePriceInput = document.getElementById('be-price');
const beAmountInput = document.getElementById('be-amount');

// Auto-suggests Amount when both Qty and Price are filled, but stays editable —
// real entries (e.g. "Transport", "Glucose") sometimes have an amount with no
// qty/price breakdown at all.
function suggestBeAmount_() {
  if (beQtyInput.value !== '' && bePriceInput.value !== '') {
    beAmountInput.value = (Number(beQtyInput.value) * Number(bePriceInput.value)).toFixed(2);
  }
}
beQtyInput.addEventListener('input', suggestBeAmount_);
bePriceInput.addEventListener('input', suggestBeAmount_);

function formatNaira_(n) {
  return '₦' + Number(n || 0).toLocaleString('en-NG', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function renderBroilerExpensesLog() {
  broilerExpensesLogEl.innerHTML = '';

  if (broilerExpensesEntries.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'loading-placeholder';
    empty.textContent = 'No entries yet.';
    broilerExpensesLogEl.appendChild(empty);
    return;
  }

  const table = document.createElement('table');
  table.className = 'data-log-table';

  const thead = document.createElement('thead');
  thead.innerHTML = '<tr><th>Date</th><th>Particulars</th><th>Qty</th><th>Price</th><th>Amount</th><th>Balance</th></tr>';
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  broilerExpensesEntries.forEach((entry) => {
    const row = document.createElement('tr');
    [
      entry.date,
      entry.particulars,
      entry.qty === '' || entry.qty === undefined || entry.qty === null ? '—' : entry.qty,
      entry.price === '' || entry.price === undefined || entry.price === null ? '—' : formatNaira_(entry.price),
      formatNaira_(entry.amount),
      formatNaira_(entry.cumulativeBalance)
    ].forEach((value) => {
      const cell = document.createElement('td');
      cell.textContent = value;
      row.appendChild(cell);
    });
    tbody.appendChild(row);
  });
  table.appendChild(tbody);

  broilerExpensesLogEl.appendChild(table);
}

async function loadBroilerExpensesLog() {
  broilerExpensesLogEl.innerHTML = '<p class="loading-placeholder">Loading log…</p>';

  let result;
  try {
    result = await callApi('listBroilerExpenses', { passcode: sessionPasscode });
  } catch (err) {
    broilerExpensesLogEl.innerHTML =
      '<p class="field-error">Couldn\'t load the log. Check your connection and reload.</p>';
    return;
  }

  if (!result.ok) {
    if (result.error === 'unauthorized') {
      clearSession();
      showView('login');
      return;
    }
    broilerExpensesLogEl.innerHTML = '<p class="field-error">Couldn\'t load the log.</p>';
    return;
  }

  broilerExpensesEntries = result.entries || [];
  renderBroilerExpensesLog();
}

broilerExpensesForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  broilerExpensesStatus.hidden = true;
  broilerExpensesError.hidden = true;

  const payload = {
    passcode: sessionPasscode,
    date: beDateInput.value,
    particulars: document.getElementById('be-particulars').value,
    qty: beQtyInput.value,
    price: bePriceInput.value,
    amount: beAmountInput.value
  };

  broilerExpensesSubmitBtn.disabled = true;
  broilerExpensesSubmitBtn.textContent = 'Saving...';

  let result;
  try {
    result = await callApi('saveBroilerExpense', payload);
  } catch (err) {
    broilerExpensesSubmitBtn.disabled = false;
    broilerExpensesSubmitBtn.textContent = 'Save Entry';
    broilerExpensesError.hidden = false;
    return;
  }

  broilerExpensesSubmitBtn.disabled = false;
  broilerExpensesSubmitBtn.textContent = 'Save Entry';

  if (!result.ok) {
    if (result.error === 'unauthorized') {
      clearSession();
      showView('login');
      return;
    }
    broilerExpensesError.hidden = false;
    return;
  }

  broilerExpensesStatus.hidden = false;
  broilerExpensesForm.reset();
  beDateInput.valueAsDate = new Date();

  // The saved entry has no cumulativeBalance (that's only computed across the
  // full ordered list on read), so reload the log for an authoritative balance column.
  await loadBroilerExpensesLog();
});

async function initBroilerExpensesApp() {
  if (!beDateInput.value) {
    beDateInput.valueAsDate = new Date();
  }
  await loadBroilerExpensesLog();
}
window.initBroilerExpensesApp = initBroilerExpensesApp;
