const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxOeKu-0x8pQhl--0itgojh6VKiwv0S1g3zMJf0l9sfpD1OYMWz6VwKfCwru1Ob0e7NEg/exec';

let sessionPasscode = null;
let sessionRole = null;
let currentFormKey = 'main';

async function callApi(action, payload = {}) {
  const res = await fetch(WEB_APP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...payload })
  });
  return res.json();
}

function clearSession() {
  sessionPasscode = null;
  sessionRole = null;
}

// Fired on page load, whenever we return to the login view, and whenever we
// land back on the chooser, so each form's getFields round trip (none need a
// passcode) overlaps with the time the person spends typing their passcode
// or looking at the chooser, instead of happening after they click a button.
//
// Each form is fetched at most once per page load — fieldsCache holds onto
// the result for the rest of the session instead of refetching every time
// someone revisits a form, so a customize-form change made elsewhere won't
// show up until the page is reloaded. A failed fetch clears its own cache
// entry so the next visit/prefetch can retry rather than being stuck.
const FORM_KEYS = ['main', 'sales', 'expenses', 'inventory'];
let fieldsCache = {};
function getFieldsCached_(formKey) {
  if (!fieldsCache[formKey]) {
    fieldsCache[formKey] = callApi('getFields', { form: formKey }).catch((err) => {
      fieldsCache[formKey] = null;
      throw err;
    });
  }
  return fieldsCache[formKey];
}
function prefetchFields() {
  FORM_KEYS.forEach((formKey) => getFieldsCached_(formKey));
}
prefetchFields();

function slugify(label, existingIds) {
  const words = label.trim().replace(/[^a-zA-Z0-9\s]/g, '').split(/\s+/).filter(Boolean);
  let base = words
    .map((w, i) => (i === 0 ? w.charAt(0).toLowerCase() + w.slice(1) : w.charAt(0).toUpperCase() + w.slice(1)))
    .join('');
  if (!base) base = 'field';
  let id = base;
  let i = 2;
  while (existingIds.includes(id)) {
    id = base + i;
    i++;
  }
  return id;
}

// ─── Views ───
const views = {
  login: document.getElementById('login-view'),
  chooser: document.getElementById('chooser-view'),
  form: document.getElementById('form-view'),
  customize: document.getElementById('customize-view'),
  receipts: document.getElementById('receipts-view'),
  broilerRearing: document.getElementById('broiler-rearing-view'),
  broilerExpenses: document.getElementById('broiler-expenses-view')
};

function showView(name) {
  Object.values(views).forEach((v) => (v.hidden = true));
  views[name].hidden = false;
}

// ─── Login ───
const loginForm = document.getElementById('login-form');
const passcodeInput = document.getElementById('passcode');
const passcodeError = document.getElementById('passcode-error');
const loginNetworkError = document.getElementById('login-network-error');
const loginSubmitBtn = loginForm.querySelector('button[type="submit"]');

loginForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const value = passcodeInput.value.trim();

  passcodeError.hidden = true;
  loginNetworkError.hidden = true;
  passcodeInput.classList.remove('error');

  loginSubmitBtn.disabled = true;
  loginSubmitBtn.textContent = 'Checking...';

  let result;
  try {
    result = await callApi('login', { passcode: value });
  } catch (err) {
    loginSubmitBtn.disabled = false;
    loginSubmitBtn.textContent = 'Enter';
    loginNetworkError.hidden = false;
    prefetchFields();
    return;
  }

  loginSubmitBtn.disabled = false;
  loginSubmitBtn.textContent = 'Enter';

  if (!result.ok) {
    if (result.error === 'invalid_passcode') {
      // prefetchedFieldsPromises are still good — same passcode, unrelated field data.
      passcodeError.hidden = false;
      passcodeInput.classList.add('error');
    } else {
      loginNetworkError.hidden = false;
      prefetchFields();
    }
    return;
  }

  sessionPasscode = value;
  sessionRole = result.role;
  loginForm.reset();
  // Receipts and Broiler Expenses are admin-only tools — hide those chooser options entirely for staff.
  document.getElementById('chooser-receipts-btn').hidden = sessionRole !== 'admin';
  document.getElementById('chooser-broiler-expenses-btn').hidden = sessionRole !== 'admin';
  showView('chooser');
});

document.getElementById('chooser-logout-btn').addEventListener('click', () => {
  clearSession();
  showView('login');
  prefetchFields();
});

async function selectForm(formKey) {
  currentFormKey = formKey;
  // Inventory's columns are edited directly in the Settings sheet, not via
  // the in-app Customize screen — both staff and admin just enter values.
  if (formKey === 'inventory' || sessionRole === 'staff') {
    await renderSubmissionForm();
    showView('form');
  } else {
    await renderFieldBuilder();
    showView('customize');
  }
}

document.getElementById('chooser-main-btn').addEventListener('click', () => selectForm('main'));
document.getElementById('chooser-sales-btn').addEventListener('click', () => selectForm('sales'));
document.getElementById('chooser-expenses-btn').addEventListener('click', () => selectForm('expenses'));
document.getElementById('chooser-inventory-btn').addEventListener('click', () => selectForm('inventory'));
document.getElementById('chooser-receipts-btn').addEventListener('click', () => {
  showView('receipts');
  if (window.initReceiptsApp) window.initReceiptsApp();
});
document.getElementById('chooser-broiler-rearing-btn').addEventListener('click', () => {
  showView('broilerRearing');
  if (window.initBroilerRearingApp) window.initBroilerRearingApp();
});
document.getElementById('chooser-broiler-expenses-btn').addEventListener('click', () => {
  showView('broilerExpenses');
  if (window.initBroilerExpensesApp) window.initBroilerExpensesApp();
});

document.getElementById('form-back-btn').addEventListener('click', () => {
  showView('chooser');
  prefetchFields();
});
document.getElementById('customize-back-btn').addEventListener('click', () => {
  showView('chooser');
  prefetchFields();
});
document.getElementById('receipts-back-btn').addEventListener('click', () => {
  showView('chooser');
  prefetchFields();
});
document.getElementById('broiler-rearing-back-btn').addEventListener('click', () => {
  showView('chooser');
  prefetchFields();
});
document.getElementById('broiler-expenses-back-btn').addEventListener('click', () => {
  showView('chooser');
  prefetchFields();
});

// ─── Submission form (fields driven by live Settings config) ───
const submissionFieldsContainer = document.getElementById('submission-fields');
const submissionForm = document.getElementById('submission-form');
const submissionStatus = document.getElementById('submission-status');
const submissionError = document.getElementById('submission-error');
const submissionSubmitBtn = submissionForm.querySelector('button[type="submit"]');

let currentFields = [];

async function renderSubmissionForm() {
  submissionStatus.hidden = true;
  submissionError.hidden = true;
  submissionFieldsContainer.innerHTML = '<p class="loading-placeholder">Loading form…</p>';

  let result;
  try {
    result = await getFieldsCached_(currentFormKey);
  } catch (err) {
    submissionFieldsContainer.innerHTML =
      '<p class="field-error">Couldn\'t load the form. Check your connection and reload.</p>';
    return;
  }

  if (!result.ok) {
    fieldsCache[currentFormKey] = null;
    submissionFieldsContainer.innerHTML =
      '<p class="field-error">Couldn\'t load the form. Please reload and try again.</p>';
    return;
  }

  currentFields = result.fields;
  submissionFieldsContainer.innerHTML = '';

  currentFields.forEach((field) => {
    const group = document.createElement('div');
    group.className = 'form-group';

    const label = document.createElement('label');
    label.setAttribute('for', 'field-' + field.id);
    label.textContent = field.label;

    let input;
    if (field.type === 'select') {
      input = document.createElement('select');
      input.id = 'field-' + field.id;
      input.className = 'form-select';
      input.required = true;

      const blankOption = document.createElement('option');
      blankOption.value = '';
      blankOption.textContent = 'Select…';
      input.appendChild(blankOption);

      (field.options || []).forEach((opt) => {
        const option = document.createElement('option');
        option.value = opt;
        option.textContent = opt;
        input.appendChild(option);
      });
    } else {
      input = document.createElement('input');
      input.id = 'field-' + field.id;
      input.className = 'form-input';
      input.required = true;
      if (field.type === 'number') {
        input.type = 'number';
        input.min = '0';
        input.step = '0.01';
        input.placeholder = '0';
      } else if (field.type === 'date') {
        input.type = 'date';
      } else {
        input.type = 'text';
        input.placeholder = field.label;
      }
    }

    group.appendChild(label);
    group.appendChild(input);
    submissionFieldsContainer.appendChild(group);
  });

  if (currentFormKey === 'sales') {
    setupSalesTotalAutoCalc_();
  }
}

// Sales-specific convenience: Total = Quantity × Price, recomputed live and
// locked from manual editing. Relies on the seeded field ids (quantity/price/
// total) — same hardcoded-by-id fragility as the Egg Production Rate calc on
// the main form, so it silently stops working if the admin renames those
// fields (renaming recomputes the id via slugify()).
function setupSalesTotalAutoCalc_() {
  const quantityEl = document.getElementById('field-quantity');
  const priceEl = document.getElementById('field-price');
  const totalEl = document.getElementById('field-total');
  if (!quantityEl || !priceEl || !totalEl) return;

  totalEl.readOnly = true;
  const recompute = () => {
    const total = (Number(quantityEl.value) || 0) * (Number(priceEl.value) || 0);
    totalEl.value = total.toFixed(2);
  };
  quantityEl.addEventListener('input', recompute);
  priceEl.addEventListener('input', recompute);
}

submissionForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  submissionStatus.hidden = true;
  submissionError.hidden = true;

  const values = {};
  currentFields.forEach((field) => {
    const el = document.getElementById('field-' + field.id);
    values[field.id] = el ? el.value : '';
  });

  submissionSubmitBtn.disabled = true;
  submissionSubmitBtn.textContent = 'Submitting...';

  let result;
  try {
    result = await callApi('submit', { passcode: sessionPasscode, form: currentFormKey, values });
  } catch (err) {
    submissionSubmitBtn.disabled = false;
    submissionSubmitBtn.textContent = 'Submit';
    submissionError.hidden = false;
    return;
  }

  submissionSubmitBtn.disabled = false;
  submissionSubmitBtn.textContent = 'Submit';

  if (!result.ok) {
    if (result.error === 'unauthorized') {
      clearSession();
      showView('login');
      return;
    }
    submissionError.hidden = false;
    return;
  }

  submissionStatus.hidden = false;
  submissionForm.reset();
});

// ─── Customize form (field builder) ───
const fieldListEl = document.getElementById('field-list');
const addFieldBtn = document.getElementById('add-field-btn');
const saveFieldsBtn = document.getElementById('save-fields-btn');
const customizeStatus = document.getElementById('customize-status');
const customizeError = document.getElementById('customize-error');

let builderFields = [];

async function renderFieldBuilder() {
  customizeStatus.hidden = true;
  customizeError.hidden = true;
  fieldListEl.innerHTML = '<p class="loading-placeholder">Loading fields…</p>';

  let result;
  try {
    result = await getFieldsCached_(currentFormKey);
  } catch (err) {
    fieldListEl.innerHTML = '<p class="field-error">Couldn\'t load fields. Check your connection and reload.</p>';
    return;
  }

  if (!result.ok) {
    fieldsCache[currentFormKey] = null;
    fieldListEl.innerHTML = '<p class="field-error">Couldn\'t load fields. Please reload and try again.</p>';
    return;
  }

  builderFields = result.fields.map((f) => ({ ...f, optionsText: (f.options || []).join(', ') }));
  drawFieldRows();
}

function drawFieldRows() {
  fieldListEl.innerHTML = '';

  builderFields.forEach((field, index) => {
    const row = document.createElement('div');
    row.className = 'custom-field-row';

    const nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.className = 'form-input';
    nameInput.value = field.label;
    nameInput.placeholder = 'Field name';
    nameInput.addEventListener('input', () => {
      builderFields[index].label = nameInput.value;
    });

    const typeSelect = document.createElement('select');
    typeSelect.className = 'custom-field-type';
    [
      ['text', 'Text'],
      ['number', 'Number'],
      ['date', 'Date'],
      ['select', 'Dropdown']
    ].forEach(([value, text]) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = text;
      if (field.type === value) option.selected = true;
      typeSelect.appendChild(option);
    });
    typeSelect.addEventListener('change', () => {
      builderFields[index].type = typeSelect.value;
      drawFieldRows();
    });

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'custom-field-delete';
    deleteBtn.setAttribute('aria-label', 'Remove field');
    deleteBtn.textContent = '×';
    deleteBtn.addEventListener('click', () => {
      builderFields.splice(index, 1);
      drawFieldRows();
    });

    row.appendChild(nameInput);
    row.appendChild(typeSelect);
    row.appendChild(deleteBtn);
    fieldListEl.appendChild(row);

    if (field.type === 'select') {
      const optionsInput = document.createElement('input');
      optionsInput.type = 'text';
      optionsInput.className = 'form-input';
      optionsInput.value = field.optionsText || '';
      optionsInput.placeholder = 'Options, comma-separated';
      optionsInput.addEventListener('input', () => {
        builderFields[index].optionsText = optionsInput.value;
      });
      fieldListEl.appendChild(optionsInput);
    }
  });
}

addFieldBtn.addEventListener('click', () => {
  builderFields.push({ id: '', label: '', type: 'text', optionsText: '' });
  drawFieldRows();
});

saveFieldsBtn.addEventListener('click', async () => {
  customizeStatus.hidden = true;
  customizeError.hidden = true;

  const usedIds = [];
  const finalFields = builderFields
    .filter((f) => f.label.trim() !== '')
    .map((f) => {
      const id = slugify(f.label, usedIds);
      usedIds.push(id);
      const field = { id, label: f.label.trim(), type: f.type };
      if (f.type === 'select') {
        field.options = (f.optionsText || '').split(',').map((o) => o.trim()).filter(Boolean);
      }
      return field;
    });

  const hasEmptySelect = finalFields.some((f) => f.type === 'select' && (!f.options || f.options.length === 0));

  if (finalFields.length === 0 || hasEmptySelect) {
    customizeError.hidden = false;
    return;
  }

  saveFieldsBtn.disabled = true;
  saveFieldsBtn.textContent = 'Saving...';

  let result;
  try {
    result = await callApi('saveFields', { passcode: sessionPasscode, form: currentFormKey, fields: finalFields });
  } catch (err) {
    saveFieldsBtn.disabled = false;
    saveFieldsBtn.textContent = 'Save';
    customizeError.hidden = false;
    return;
  }

  saveFieldsBtn.disabled = false;
  saveFieldsBtn.textContent = 'Save';

  if (!result.ok) {
    if (result.error === 'unauthorized') {
      clearSession();
      showView('login');
      return;
    }
    customizeError.hidden = false;
    return;
  }

  builderFields = result.fields.map((f) => ({ ...f, optionsText: (f.options || []).join(', ') }));
  drawFieldRows();
  customizeStatus.hidden = false;

  // Keep the cache in sync so a staff member (or the admin, revisiting this
  // form later in the same page load) sees the just-saved fields instead of
  // whatever was cached before this save.
  fieldsCache[currentFormKey] = Promise.resolve({ ok: true, fields: result.fields });
});
