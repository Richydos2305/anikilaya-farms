// Broiler Rearing Record — staff logs each day's entry manually (no auto-chaining
// between rows); both staff and admin can view the log, but only staff get the form.

let broilerRearingEntries = [];

const broilerRearingFormWrap = document.getElementById('broiler-rearing-form-wrap');
const broilerRearingForm = document.getElementById('broiler-rearing-form');
const broilerRearingStatus = document.getElementById('broiler-rearing-status');
const broilerRearingError = document.getElementById('broiler-rearing-error');
const broilerRearingSubmitBtn = broilerRearingForm.querySelector('button[type="submit"]');
const broilerRearingLogEl = document.getElementById('broiler-rearing-log');

const brDateInput = document.getElementById('br-date');
const brOpeningStockInput = document.getElementById('br-opening-stock');
const brMortalityInput = document.getElementById('br-mortality');
const brClosingStockInput = document.getElementById('br-closing-stock');

// Closing Stock = Opening Stock − Mortality, live preview only — the server
// recomputes and stores the authoritative value on save.
function recomputeBrClosingStock_() {
  const opening = Number(brOpeningStockInput.value) || 0;
  const mortality = Number(brMortalityInput.value) || 0;
  brClosingStockInput.value = opening - mortality;
}
brOpeningStockInput.addEventListener('input', recomputeBrClosingStock_);
brMortalityInput.addEventListener('input', recomputeBrClosingStock_);

function renderBroilerRearingLog() {
  broilerRearingLogEl.innerHTML = '';

  if (broilerRearingEntries.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'loading-placeholder';
    empty.textContent = 'No entries yet.';
    broilerRearingLogEl.appendChild(empty);
    return;
  }

  const table = document.createElement('table');
  table.className = 'data-log-table';

  const thead = document.createElement('thead');
  thead.innerHTML =
    '<tr><th>Date</th><th>Age</th><th>Opening</th><th>Mortality</th><th>Closing</th>' +
    '<th>Feed</th><th>Medication</th><th>Avg B/W</th><th>Comment</th></tr>';
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  broilerRearingEntries.forEach((entry) => {
    const row = document.createElement('tr');
    [
      entry.date, entry.age, entry.openingStock, entry.mortality, entry.closingStock,
      entry.feedConsumed, entry.medication, entry.avgBodyWeight, entry.comment
    ].forEach((value) => {
      const cell = document.createElement('td');
      cell.textContent = value === undefined || value === null || value === '' ? '—' : value;
      row.appendChild(cell);
    });
    tbody.appendChild(row);
  });
  table.appendChild(tbody);

  broilerRearingLogEl.appendChild(table);
}

async function loadBroilerRearingLog() {
  broilerRearingLogEl.innerHTML = '<p class="loading-placeholder">Loading log…</p>';

  let result;
  try {
    result = await callApi('listBroilerRearing', { passcode: sessionPasscode });
  } catch (err) {
    broilerRearingLogEl.innerHTML =
      '<p class="field-error">Couldn\'t load the log. Check your connection and reload.</p>';
    return;
  }

  if (!result.ok) {
    if (result.error === 'unauthorized') {
      clearSession();
      showView('login');
      return;
    }
    broilerRearingLogEl.innerHTML = '<p class="field-error">Couldn\'t load the log.</p>';
    return;
  }

  broilerRearingEntries = result.entries || [];
  renderBroilerRearingLog();
}

broilerRearingForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  broilerRearingStatus.hidden = true;
  broilerRearingError.hidden = true;

  const payload = {
    passcode: sessionPasscode,
    date: brDateInput.value,
    age: document.getElementById('br-age').value,
    openingStock: brOpeningStockInput.value,
    mortality: brMortalityInput.value,
    feedConsumed: document.getElementById('br-feed-consumed').value,
    medication: document.getElementById('br-medication').value,
    avgBodyWeight: document.getElementById('br-avg-body-weight').value,
    comment: document.getElementById('br-comment').value
  };

  broilerRearingSubmitBtn.disabled = true;
  broilerRearingSubmitBtn.textContent = 'Saving...';

  let result;
  try {
    result = await callApi('saveBroilerRearing', payload);
  } catch (err) {
    broilerRearingSubmitBtn.disabled = false;
    broilerRearingSubmitBtn.textContent = 'Save Entry';
    broilerRearingError.hidden = false;
    return;
  }

  broilerRearingSubmitBtn.disabled = false;
  broilerRearingSubmitBtn.textContent = 'Save Entry';

  if (!result.ok) {
    if (result.error === 'unauthorized') {
      clearSession();
      showView('login');
      return;
    }
    broilerRearingError.hidden = false;
    return;
  }

  broilerRearingEntries.unshift(result.entry);
  renderBroilerRearingLog();
  broilerRearingStatus.hidden = false;
  broilerRearingForm.reset();
  brDateInput.valueAsDate = new Date();
});

async function initBroilerRearingApp() {
  broilerRearingFormWrap.hidden = sessionRole !== 'staff';
  if (sessionRole === 'staff' && !brDateInput.value) {
    brDateInput.valueAsDate = new Date();
  }
  await loadBroilerRearingLog();
}
window.initBroilerRearingApp = initBroilerRearingApp;
