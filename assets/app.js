const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxOeKu-0x8pQhl--0itgojh6VKiwv0S1g3zMJf0l9sfpD1OYMWz6VwKfCwru1Ob0e7NEg/exec';

let sessionPasscode = null;
let sessionRole = null;

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

// Fired on page load and whenever we return to the login view, so the
// getFields round trip (which needs no passcode) overlaps with the time the
// person spends typing their passcode instead of happening after login.
let prefetchedFieldsPromise = null;
function prefetchFields() {
  prefetchedFieldsPromise = callApi('getFields');
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
  form: document.getElementById('form-view'),
  customize: document.getElementById('customize-view')
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
      // prefetchedFieldsPromise is still good — same passcode, unrelated field data.
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

  if (result.role === 'staff') {
    await renderSubmissionForm();
    showView('form');
  } else {
    await renderFieldBuilder();
    showView('customize');
  }
});

document.getElementById('form-back-btn').addEventListener('click', () => {
  clearSession();
  showView('login');
  prefetchFields();
});
document.getElementById('customize-back-btn').addEventListener('click', () => {
  clearSession();
  showView('login');
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

  const fieldsPromise = prefetchedFieldsPromise || callApi('getFields');
  prefetchedFieldsPromise = null;

  let result;
  try {
    result = await fieldsPromise;
  } catch (err) {
    submissionFieldsContainer.innerHTML =
      '<p class="field-error">Couldn\'t load the form. Check your connection and reload.</p>';
    return;
  }

  if (!result.ok) {
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

    const input = document.createElement('input');
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

    group.appendChild(label);
    group.appendChild(input);
    submissionFieldsContainer.appendChild(group);
  });
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
    result = await callApi('submit', { passcode: sessionPasscode, values });
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

  const fieldsPromise = prefetchedFieldsPromise || callApi('getFields');
  prefetchedFieldsPromise = null;

  let result;
  try {
    result = await fieldsPromise;
  } catch (err) {
    fieldListEl.innerHTML = '<p class="field-error">Couldn\'t load fields. Check your connection and reload.</p>';
    return;
  }

  if (!result.ok) {
    fieldListEl.innerHTML = '<p class="field-error">Couldn\'t load fields. Please reload and try again.</p>';
    return;
  }

  builderFields = result.fields.map((f) => ({ ...f }));
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
      ['date', 'Date']
    ].forEach(([value, text]) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = text;
      if (field.type === value) option.selected = true;
      typeSelect.appendChild(option);
    });
    typeSelect.addEventListener('change', () => {
      builderFields[index].type = typeSelect.value;
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
  });
}

addFieldBtn.addEventListener('click', () => {
  builderFields.push({ id: '', label: '', type: 'text' });
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
      return { id, label: f.label.trim(), type: f.type };
    });

  if (finalFields.length === 0) {
    customizeError.hidden = false;
    return;
  }

  saveFieldsBtn.disabled = true;
  saveFieldsBtn.textContent = 'Saving...';

  let result;
  try {
    result = await callApi('saveFields', { passcode: sessionPasscode, fields: finalFields });
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

  builderFields = result.fields.map((f) => ({ ...f }));
  drawFieldRows();
  customizeStatus.hidden = false;
});
