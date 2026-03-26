/* ===================================================
   NBSC OJT Weekly Report System — script.js
   Northern Bukidnon State College | Form 9
=================================================== */

'use strict';

// ── State ──────────────────────────────────────────
const state = {
  objectives:  [],
  activities:  [],
  reflections: [],
  docs:        [],          // { url, caption, type, name, dataUrl }
  pendingFiles: []          // files staged in upload modal
};

let SCRIPT_URL = localStorage.getItem('ojt_script_url') || '';

// ── Init ────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  if (SCRIPT_URL) document.getElementById('scriptUrl').value = SCRIPT_URL;
  loadLocalDraft();
  setupDropZone();
  setupLiveProgress();
  setInterval(autoSaveDraft, 30000);  // auto-save draft every 30s
});

// ══════════════════════════════════════════════════
//  MODAL HELPERS
// ══════════════════════════════════════════════════
function openModal(name) {
  const el = document.getElementById('modal' + cap(name));
  if (!el) return;
  el.classList.add('open');
  // Focus first input
  setTimeout(() => {
    const first = el.querySelector('input:not([type=hidden]),select,textarea');
    if (first) first.focus();
  }, 100);
}
function closeModal(name) {
  const el = document.getElementById('modal' + cap(name));
  if (el) el.classList.remove('open');
  resetModal(name);
}
function resetModal(name) {
  if (name === 'objectives') {
    document.getElementById('objEditIdx').value = '-1';
    document.getElementById('objText').value = '';
    document.getElementById('objStatus').value = 'In Progress';
  } else if (name === 'activities') {
    document.getElementById('actEditIdx').value = '-1';
    document.getElementById('actTitle').value = '';
    document.getElementById('actDesc').value = '';
    document.getElementById('actDate').value = '';
    document.getElementById('actHours').value = '';
  } else if (name === 'reflections') {
    document.getElementById('refEditIdx').value = '-1';
    document.getElementById('refType').value = 'Learning';
    document.getElementById('refContent').value = '';
  } else if (name === 'documentation') {
    state.pendingFiles = [];
    document.getElementById('previewGrid').innerHTML = '';
    document.getElementById('docCaption').value = '';
    document.getElementById('fileInput').value = '';
  }
}

// Close modal on backdrop click
document.querySelectorAll('.modal-overlay').forEach(overlay => {
  overlay.addEventListener('click', e => {
    if (e.target === overlay) {
      const name = overlay.id.replace('modal', '').toLowerCase();
      closeModal(name);
    }
  });
});

// Close on Escape
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    document.querySelectorAll('.modal-overlay.open').forEach(m => {
      const name = m.id.replace('modal', '').toLowerCase();
      closeModal(name);
    });
  }
});

// ══════════════════════════════════════════════════
//  OBJECTIVES
// ══════════════════════════════════════════════════
function saveObjective() {
  const text   = document.getElementById('objText').value.trim();
  const status = document.getElementById('objStatus').value;
  const idx    = parseInt(document.getElementById('objEditIdx').value);

  if (!text) { shake('objText'); showToast('Please enter an objective statement.', 'error'); return; }

  if (idx >= 0) {
    state.objectives[idx] = { text, status };
  } else {
    state.objectives.push({ text, status });
  }

  renderObjectives();
  closeModal('objectives');
  updateStats();
  updateProgress();
  saveDraft();
  showToast(idx >= 0 ? 'Objective updated.' : 'Objective added!', 'success');
}

function editObjective(idx) {
  const obj = state.objectives[idx];
  document.getElementById('objEditIdx').value = idx;
  document.getElementById('objText').value   = obj.text;
  document.getElementById('objStatus').value = obj.status;
  openModal('objectives');
}

function deleteObjective(idx) {
  state.objectives.splice(idx, 1);
  renderObjectives();
  updateStats();
  updateProgress();
  saveDraft();
  showToast('Objective removed.', 'info');
}

function renderObjectives() {
  const list = document.getElementById('objectivesList');
  if (state.objectives.length === 0) {
    list.innerHTML = `<div class="placeholder-row"><span class="placeholder-num">1.</span><span class="placeholder-line">Click <strong>Add Objective</strong> to enter your weekly objective</span></div>`;
    return;
  }
  list.innerHTML = state.objectives.map((obj, i) => {
    const statusColors = {
      'Completed':    '#bbf7d0|#166534',
      'In Progress':  '#bfdbfe|#1e3a8a',
      'Pending':      '#fef3c7|#92400e',
      'Carried Over': '#e0e7ff|#3730a3',
    };
    const [bg, fg] = (statusColors[obj.status] || '#e5e7eb|#374151').split('|');
    return `
    <div class="obj-item">
      <span class="obj-num">${i + 1}.</span>
      <div class="obj-content">
        <div class="obj-text">${escHtml(obj.text)}</div>
        <span class="obj-status-badge" style="background:${bg};color:${fg};">${obj.status}</span>
      </div>
      <div class="obj-actions">
        <button class="icon-btn" title="Edit" onclick="editObjective(${i})">✏️</button>
        <button class="icon-btn" title="Delete" onclick="deleteObjective(${i})">🗑️</button>
      </div>
    </div>`;
  }).join('');
}

// ══════════════════════════════════════════════════
//  ACTIVITIES
// ══════════════════════════════════════════════════
function saveActivity() {
  const title = document.getElementById('actTitle').value.trim();
  const desc  = document.getElementById('actDesc').value.trim();
  const date  = document.getElementById('actDate').value;
  const hours = document.getElementById('actHours').value;
  const idx   = parseInt(document.getElementById('actEditIdx').value);

  if (!title) { shake('actTitle'); showToast('Please enter an activity title.', 'error'); return; }

  const entry = { title, desc, date, hours };
  if (idx >= 0) {
    state.activities[idx] = entry;
  } else {
    state.activities.push(entry);
  }

  renderARTable();
  closeModal('activities');
  updateStats();
  updateProgress();
  saveDraft();
  showToast(idx >= 0 ? 'Activity updated.' : 'Activity added!', 'success');
}

function editActivity(idx) {
  const a = state.activities[idx];
  document.getElementById('actEditIdx').value = idx;
  document.getElementById('actTitle').value   = a.title;
  document.getElementById('actDesc').value    = a.desc;
  document.getElementById('actDate').value    = a.date;
  document.getElementById('actHours').value   = a.hours;
  openModal('activities');
}

function deleteActivity(idx) {
  state.activities.splice(idx, 1);
  renderARTable();
  updateStats();
  updateProgress();
  saveDraft();
  showToast('Activity removed.', 'info');
}

// ══════════════════════════════════════════════════
//  REFLECTIONS
// ══════════════════════════════════════════════════
function saveReflection() {
  const type    = document.getElementById('refType').value;
  const content = document.getElementById('refContent').value.trim();
  const idx     = parseInt(document.getElementById('refEditIdx').value);

  if (!content) { shake('refContent'); showToast('Please write your reflection.', 'error'); return; }

  const entry = { type, content };
  if (idx >= 0) {
    state.reflections[idx] = entry;
  } else {
    state.reflections.push(entry);
  }

  renderARTable();
  closeModal('reflections');
  updateStats();
  updateProgress();
  saveDraft();
  showToast(idx >= 0 ? 'Reflection updated.' : 'Reflection added!', 'success');
}

function editReflection(idx) {
  const r = state.reflections[idx];
  document.getElementById('refEditIdx').value  = idx;
  document.getElementById('refType').value     = r.type;
  document.getElementById('refContent').value  = r.content;
  openModal('reflections');
}

function deleteReflection(idx) {
  state.reflections.splice(idx, 1);
  renderARTable();
  updateStats();
  updateProgress();
  saveDraft();
  showToast('Reflection removed.', 'info');
}

// ── Render Activity/Reflection table ───────────────
function renderARTable() {
  const tbody      = document.getElementById('arTableBody');
  const emptyRow   = document.getElementById('arEmptyRow');
  const maxRows    = Math.max(state.activities.length, state.reflections.length, 1);
  const hasContent = state.activities.length > 0 || state.reflections.length > 0;

  // Remove all dynamic rows
  tbody.querySelectorAll('.ar-data-row').forEach(r => r.remove());

  if (!hasContent) {
    if (emptyRow) emptyRow.style.display = '';
    return;
  }
  if (emptyRow) emptyRow.style.display = 'none';

  for (let i = 0; i < maxRows; i++) {
    const act = state.activities[i];
    const ref = state.reflections[i];
    const tr  = document.createElement('tr');
    tr.className = 'ar-data-row';
    tr.innerHTML = `
      <td>
        <div class="ar-cell ar-act-cell">
          ${act ? `
            <div class="ar-cell-badge">${act.date ? fmtDate(act.date) : ''}${act.hours ? ` · ${act.hours}h` : ''}</div>
            <div class="ar-cell-title">${escHtml(act.title)}</div>
            ${act.desc ? `<div class="ar-cell-sub">${escHtml(act.desc)}</div>` : ''}
            <div class="ar-row-actions">
              <button class="icon-btn" onclick="editActivity(${i})" title="Edit">✏️</button>
              <button class="icon-btn" onclick="deleteActivity(${i})" title="Delete">🗑️</button>
            </div>` : ''}
        </div>
      </td>
      <td>
        <div class="ar-cell">
          ${ref ? `
            <div class="ar-cell-badge ref-badge">${escHtml(ref.type)}</div>
            <div class="ar-cell-sub">${escHtml(ref.content)}</div>
            <div class="ar-row-actions">
              <button class="icon-btn" onclick="editReflection(${i})" title="Edit">✏️</button>
              <button class="icon-btn" onclick="deleteReflection(${i})" title="Delete">🗑️</button>
            </div>` : ''}
        </div>
      </td>`;
    tbody.insertBefore(tr, tbody.querySelector('.ar-blank-row'));
  }
}

// ══════════════════════════════════════════════════
//  DOCUMENTATION / FILE UPLOAD
// ══════════════════════════════════════════════════
function setupDropZone() {
  const dz = document.getElementById('dropZone');
  if (!dz) return;
  ['dragenter','dragover'].forEach(ev => {
    dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('dragging'); });
  });
  ['dragleave','drop'].forEach(ev => {
    dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('dragging'); });
  });
  dz.addEventListener('drop', e => {
    const files = Array.from(e.dataTransfer.files);
    processFiles(files);
  });
}

function handleFileSelect(event) {
  processFiles(Array.from(event.target.files));
}

function processFiles(files) {
  const MAX = 10 * 1024 * 1024; // 10MB
  const allowed = ['image/png','image/jpeg','image/gif','image/webp','application/pdf'];
  files.forEach(file => {
    if (!allowed.includes(file.type)) { showToast(`${file.name}: unsupported file type.`, 'error'); return; }
    if (file.size > MAX) { showToast(`${file.name}: exceeds 10MB limit.`, 'error'); return; }
    const reader = new FileReader();
    reader.onload = e => {
      state.pendingFiles.push({ file, dataUrl: e.target.result, name: file.name, type: file.type });
      renderPreviewGrid();
      document.getElementById('captionGroup') && (document.getElementById('captionGroup').style.display = 'block');
    };
    reader.readAsDataURL(file);
  });
}

function renderPreviewGrid() {
  const grid = document.getElementById('previewGrid');
  grid.innerHTML = state.pendingFiles.map((f, i) => {
    if (f.type === 'application/pdf') {
      return `<div class="prev-item">
        <div style="padding:.75rem;text-align:center;font-size:2rem;">📄</div>
        <div class="prev-name">${escHtml(f.name)}</div>
        <button class="prev-remove" onclick="removePending(${i})">✕</button>
      </div>`;
    }
    return `<div class="prev-item">
      <img class="prev-img" src="${f.dataUrl}" alt="${escHtml(f.name)}" />
      <div class="prev-name">${escHtml(f.name)}</div>
      <button class="prev-remove" onclick="removePending(${i})">✕</button>
    </div>`;
  }).join('');
}

function removePending(idx) {
  state.pendingFiles.splice(idx, 1);
  renderPreviewGrid();
}

function confirmUpload() {
  if (state.pendingFiles.length === 0) { showToast('No files selected.', 'warning'); return; }
  const caption = document.getElementById('docCaption').value.trim();

  state.pendingFiles.forEach(f => {
    state.docs.push({ dataUrl: f.dataUrl, name: f.name, type: f.type, caption, url: null });
  });

  // If script URL set, upload to Drive in background
  if (SCRIPT_URL) {
    state.docs.filter(d => d.url === null).forEach((doc, i) => {
      uploadFileToDrive(doc, i);
    });
  }

  renderGallery();
  closeModal('documentation');
  updateStats();
  saveDraft();
  showToast(`${state.pendingFiles.length} file(s) added!`, 'success');
}

async function uploadFileToDrive(doc, idx) {
  if (!SCRIPT_URL) return;
  try {
    const base64 = doc.dataUrl.split(',')[1];
    const res = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'uploadFile', fileName: doc.name, mimeType: doc.type, data: base64 })
    });
    const json = await res.json();
    if (json.status === 'ok') {
      doc.url = json.fileUrl;
      showToast('File uploaded to Drive ✓', 'info');
    }
  } catch (err) {
    console.warn('Drive upload failed:', err);
  }
}

function renderGallery() {
  const gallery = document.getElementById('docGallery');
  if (state.docs.length === 0) {
    gallery.innerHTML = `<div class="doc-empty"><div class="doc-empty-icon">📷</div><p>No documentation uploaded yet.<br/>Click <strong>Upload Images</strong> to add photos or files.</p></div>`;
    return;
  }
  gallery.innerHTML = `<div class="doc-grid">${state.docs.map((doc, i) => {
    if (doc.type === 'application/pdf') {
      return `<div class="doc-file-item">
        <span class="doc-file-icon">📄</span>
        <div style="flex:1;min-width:0;">
          <div style="font-weight:600;font-size:.78rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escHtml(doc.name)}</div>
          ${doc.caption ? `<div style="font-size:.68rem;color:var(--text-muted);">${escHtml(doc.caption)}</div>` : ''}
          ${doc.url ? `<a href="${doc.url}" target="_blank" style="font-size:.68rem;">View →</a>` : ''}
        </div>
        <button class="doc-remove" onclick="removeDoc(${i})">✕</button>
      </div>`;
    }
    return `<div class="doc-item">
      <img class="doc-img" src="${doc.dataUrl || doc.url}" alt="${escHtml(doc.name)}" onclick="lightbox('${doc.dataUrl || doc.url}')" style="cursor:pointer;" />
      ${doc.caption ? `<div class="doc-caption">${escHtml(doc.caption)}</div>` : ''}
      <button class="doc-remove" onclick="removeDoc(${i})">✕</button>
    </div>`;
  }).join('')}</div>`;
}

function removeDoc(idx) {
  state.docs.splice(idx, 1);
  renderGallery();
  updateStats();
  saveDraft();
  showToast('File removed.', 'info');
}

function lightbox(src) {
  const lb = document.createElement('div');
  lb.style.cssText = 'position:fixed;inset:0;z-index:5000;background:rgba(0,0,0,.9);display:flex;align-items:center;justify-content:center;cursor:pointer;';
  lb.innerHTML = `<img src="${src}" style="max-width:90vw;max-height:90vh;border-radius:8px;box-shadow:0 0 40px rgba(0,0,0,.5);" />`;
  lb.onclick = () => document.body.removeChild(lb);
  document.body.appendChild(lb);
}

// ══════════════════════════════════════════════════
//  GOOGLE SHEETS – SAVE
// ══════════════════════════════════════════════════
async function saveToSheets() {
  if (!SCRIPT_URL) {
    showToast('No Apps Script URL set. Please connect first.', 'warning');
    openModal('setup');
    return;
  }
  if (!validateForm()) return;

  showLoader('Saving to Google Sheets…');
  const payload = buildPayload();

  try {
    const res  = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'saveReport', ...payload })
    });
    const json = await res.json();
    if (json.status === 'ok') {
      const ts = new Date().toLocaleString();
      document.getElementById('lastSaved').textContent = `Last saved: ${ts}`;
      showToast('✅ Report saved to Google Sheets!', 'success');
      saveDraft();
    } else {
      showToast(`Save failed: ${json.message || 'Unknown error'}`, 'error');
    }
  } catch (err) {
    showToast('Network error. Check your script URL and try again.', 'error');
    console.error(err);
  } finally {
    hideLoader();
  }
}

// ── Load from Sheets ────────────────────────────────
async function loadFromSheets() {
  if (!SCRIPT_URL) {
    showToast('No Apps Script URL set.', 'warning');
    return;
  }
  const name    = document.getElementById('traineeName').value.trim();
  const week    = document.getElementById('weekNumber').value;
  const company = document.getElementById('company').value.trim();

  if (!name || !week) {
    showToast('Enter Name and Week before loading.', 'warning');
    return;
  }
  showLoader('Loading from Google Sheets…');

  try {
    const url  = `${SCRIPT_URL}?action=getReport&name=${encodeURIComponent(name)}&week=${encodeURIComponent(week)}&company=${encodeURIComponent(company)}`;
    const res  = await fetch(url);
    const json = await res.json();

    if (json.status === 'ok' && json.data) {
      populateForm(json.data);
      showToast('✅ Data loaded from Google Sheets!', 'success');
    } else {
      showToast('No saved report found for this Name/Week.', 'info');
    }
  } catch (err) {
    showToast('Failed to load data. Check network and script URL.', 'error');
    console.error(err);
  } finally {
    hideLoader();
  }
}

function populateForm(data) {
  document.getElementById('traineeName').value  = data.name    || '';
  document.getElementById('company').value      = data.company || '';
  document.getElementById('weekNumber').value   = data.week    || '';

  state.objectives  = safeJSON(data.objectives, []);
  state.activities  = safeJSON(data.activities, []);
  state.reflections = safeJSON(data.reflections, []);
  state.docs        = (safeJSON(data.imageUrls, []) || []).map(u => ({ url: u, dataUrl: u, name: u.split('/').pop(), type: 'image/jpeg', caption: '' }));

  document.getElementById('sigTrainee').value       = data.sig_trainee     || '';
  document.getElementById('sigSupervisor').value    = data.sig_supervisor   || '';
  document.getElementById('sigCoordinator').value   = data.sig_coordinator  || '';

  renderObjectives();
  renderARTable();
  renderGallery();
  updateStats();
  updateProgress();
}

function buildPayload() {
  return {
    name:           document.getElementById('traineeName').value.trim(),
    company:        document.getElementById('company').value.trim(),
    week:           document.getElementById('weekNumber').value,
    objectives:     JSON.stringify(state.objectives),
    activities:     JSON.stringify(state.activities),
    reflections:    JSON.stringify(state.reflections),
    imageUrls:      JSON.stringify(state.docs.map(d => d.url || d.dataUrl)),
    sig_trainee:    document.getElementById('sigTrainee').value.trim(),
    sig_supervisor: document.getElementById('sigSupervisor').value.trim(),
    sig_coordinator:document.getElementById('sigCoordinator').value.trim(),
    timestamp:      new Date().toISOString()
  };
}

// ══════════════════════════════════════════════════
//  LOCAL DRAFT (localStorage)
// ══════════════════════════════════════════════════
function saveDraft() {
  try {
    const draft = {
      ...buildPayload(),
      docs_local: state.docs.map(d => ({ ...d, dataUrl: d.dataUrl ? d.dataUrl.substring(0, 200) : '' }))
    };
    localStorage.setItem('ojt_draft', JSON.stringify(draft));
  } catch (e) { /* quota exceeded – skip */ }
}

function loadLocalDraft() {
  try {
    const raw = localStorage.getItem('ojt_draft');
    if (!raw) return;
    const draft = JSON.parse(raw);
    document.getElementById('traineeName').value = draft.name    || '';
    document.getElementById('company').value     = draft.company || '';
    document.getElementById('weekNumber').value  = draft.week    || '';
    state.objectives  = safeJSON(draft.objectives, []);
    state.activities  = safeJSON(draft.activities, []);
    state.reflections = safeJSON(draft.reflections, []);
    document.getElementById('sigTrainee').value       = draft.sig_trainee     || '';
    document.getElementById('sigSupervisor').value    = draft.sig_supervisor   || '';
    document.getElementById('sigCoordinator').value   = draft.sig_coordinator  || '';
    renderObjectives();
    renderARTable();
    updateStats();
    updateProgress();
  } catch (e) { console.warn('Draft load failed:', e); }
}

function autoSaveDraft() {
  saveDraft();
}

// ══════════════════════════════════════════════════
//  CLEAR FORM
// ══════════════════════════════════════════════════
function clearForm() {
  if (!confirm('Clear all form data? This cannot be undone.')) return;
  ['traineeName','company','sigTrainee','sigSupervisor','sigCoordinator'].forEach(id => {
    document.getElementById(id).value = '';
  });
  ['weekNumber'].forEach(id => document.getElementById(id).value = '');
  document.querySelectorAll('.sig-date-input').forEach(el => el.value = '');
  state.objectives  = [];
  state.activities  = [];
  state.reflections = [];
  state.docs        = [];
  renderObjectives();
  renderARTable();
  renderGallery();
  updateStats();
  updateProgress();
  localStorage.removeItem('ojt_draft');
  document.getElementById('lastSaved').textContent = 'Last saved: —';
  showToast('Form cleared.', 'info');
}

// ══════════════════════════════════════════════════
//  VALIDATION
// ══════════════════════════════════════════════════
function validateForm() {
  const name    = document.getElementById('traineeName').value.trim();
  const company = document.getElementById('company').value.trim();
  const week    = document.getElementById('weekNumber').value;

  if (!name)    { shake('traineeName'); showToast('Please enter trainee name.', 'error'); return false; }
  if (!company) { shake('company');     showToast('Please enter company name.', 'error'); return false; }
  if (!week)    { showToast('Please select a week.', 'error'); return false; }
  if (state.objectives.length === 0) { showToast('Please add at least one objective.', 'error'); return false; }
  return true;
}

// ══════════════════════════════════════════════════
//  STATS & PROGRESS
// ══════════════════════════════════════════════════
function updateStats() {
  document.getElementById('statObj').textContent = state.objectives.length;
  document.getElementById('statAct').textContent = state.activities.length;
  document.getElementById('statRef').textContent = state.reflections.length;
  document.getElementById('statDoc').textContent = state.docs.length;
}

function updateProgress() {
  const checks = {
    ciName:    !!document.getElementById('traineeName').value.trim(),
    ciCompany: !!document.getElementById('company').value.trim(),
    ciWeek:    !!document.getElementById('weekNumber').value,
    ciObj:     state.objectives.length > 0,
    ciAct:     state.activities.length > 0,
    ciRef:     state.reflections.length > 0,
  };
  const done  = Object.values(checks).filter(Boolean).length;
  const total = Object.keys(checks).length;
  const pct   = Math.round((done / total) * 100);

  Object.entries(checks).forEach(([id, val]) => {
    const dot = document.getElementById(id);
    if (dot) dot.classList.toggle('done', val);
  });

  document.getElementById('progPct').textContent = pct + '%';
  // Circumference = 2π × 36 ≈ 226.2
  const offset = 226.2 - (226.2 * pct / 100);
  document.getElementById('progRingFill').style.strokeDashoffset = offset;
}

function setupLiveProgress() {
  ['traineeName','company','weekNumber'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', () => { updateProgress(); saveDraft(); });
  });
}

// ══════════════════════════════════════════════════
//  EXPORT – PDF
// ══════════════════════════════════════════════════
function exportPDF() {
  showLoader('Generating PDF…');
  const name = document.getElementById('traineeName').value.trim() || 'OJT';
  const week = document.getElementById('weekNumber').value || 'Report';

  const opt = {
    margin:    [10, 10, 10, 10],
    filename:  `${name}_${week}_OJT_Report.pdf`,
    image:     { type: 'jpeg', quality: 0.95 },
    html2canvas: { scale: 2, useCORS: true, letterRendering: true },
    jsPDF:     { unit: 'mm', format: 'a4', orientation: 'portrait' },
    pagebreak: { mode: ['css', 'legacy'], before: '.page-break-before' }
  };

  // Temporarily hide non-printable elements
  document.querySelectorAll('.obj-actions,.ar-row-actions,.doc-remove,.add-btn').forEach(el => el.style.visibility = 'hidden');

  html2pdf().set(opt).from(document.getElementById('reportForm')).save()
    .then(() => {
      document.querySelectorAll('.obj-actions,.ar-row-actions,.doc-remove,.add-btn').forEach(el => el.style.visibility = '');
      hideLoader();
      showToast('PDF downloaded!', 'success');
    })
    .catch(err => {
      document.querySelectorAll('.obj-actions,.ar-row-actions,.doc-remove,.add-btn').forEach(el => el.style.visibility = '');
      hideLoader();
      showToast('PDF export failed. Try printing instead.', 'error');
      console.error(err);
    });
}

// ══════════════════════════════════════════════════
//  EXPORT – DOCX (pure JS via template approach)
// ══════════════════════════════════════════════════
function exportDOCX() {
  showLoader('Generating DOCX…');
  try {
    const name    = document.getElementById('traineeName').value.trim() || '(Name)';
    const company = document.getElementById('company').value.trim() || '(Company)';
    const week    = document.getElementById('weekNumber').value || '(Week)';

    const objText = state.objectives.map((o, i) => `${i+1}. ${o.text} [${o.status}]`).join('\n');
    const actText = state.activities.map((a, i) =>
      `${i+1}. ${a.title}${a.date ? ' ('+fmtDate(a.date)+')' : ''}${a.hours ? ' - '+a.hours+'h' : ''}\n   ${a.desc || ''}`
    ).join('\n');
    const refText = state.reflections.map((r, i) =>
      `${i+1}. [${r.type}] ${r.content}`
    ).join('\n');

    const sigTrainee     = document.getElementById('sigTrainee').value.trim();
    const sigSupervisor  = document.getElementById('sigSupervisor').value.trim();
    const sigCoordinator = document.getElementById('sigCoordinator').value.trim();

    // Build RTF-based DOC (universal compatibility trick)
    const rtf = `{\\rtf1\\ansi\\deff0
{\\fonttbl{\\f0 Times New Roman;}{\\f1 Arial;}}
{\\colortbl;\\red26\\green71\\blue49;}
\\f1\\fs20
\\qc\\b\\fs28 NORTHERN BUKIDNON STATE COLLEGE\\b0\\fs20\\par
\\qc Republic of the Philippines | Manolo Fortich, 8703 Bukidnon\\par
\\qc\\i Creando futura, Transformationis vitae, Ductae a Deo\\i0\\par
\\par
{\\pard\\qc\\b\\fs26 ON-THE-JOB TRAINING LOG SHEET\\b0\\par}
{\\pard\\qc\\b WEEKLY PROGRESS REPORT\\b0\\par}
{\\pard\\qc Term Second Semester AY 2025-2026\\par}
\\par
{\\pard\\b Form 9\\b0\\par}
\\par
\\trowd\\trgaph108\\trleft-108
\\cellx2000\\cellx9000
{\\pard\\intbl\\b Name:\\b0\\cell}{\\pard\\intbl ${rtfEsc(name)}\\cell}\\row
{\\pard\\intbl\\b Company:\\b0\\cell}{\\pard\\intbl ${rtfEsc(company)}\\cell}\\row
{\\pard\\intbl\\b Week:\\b0\\cell}{\\pard\\intbl ${rtfEsc(week)}\\cell}\\row
\\par
\\pard\\b Objective(s) for the week:\\b0\\par
${rtfEsc(objText)}\\par
\\par
\\trowd\\trgaph108\\trleft-108
\\cellx4600\\cellx9000
{\\pard\\intbl\\b ACTIVITIES:\\b0\\cell}{\\pard\\intbl\\b REFLECTIONS:\\b0\\cell}\\row
${buildARRows()}
\\par
\\b Signed:\\b0\\par
\\par
\\trowd\\trgaph108\\trleft-108
\\cellx4000\\cellx9000
{\\pard\\intbl _________________________\\par ${rtfEsc(sigTrainee)}\\par\\b Student Trainee\\b0\\cell}
{\\pard\\intbl _________________________\\par ${rtfEsc(sigSupervisor)}\\par\\b HTE Supervisor\\b0\\cell}\\row
\\par
{\\pard _________________________\\par ${rtfEsc(sigCoordinator)}\\par\\b OJT Coordinator\\b0\\par}
}`;

    function buildARRows() {
      const max = Math.max(state.activities.length, state.reflections.length, 5);
      let rows = '';
      for (let i = 0; i < max; i++) {
        const a = state.activities[i]  ? rtfEsc(state.activities[i].title)  : ' ';
        const r = state.reflections[i] ? rtfEsc(state.reflections[i].content) : ' ';
        rows += `{\\pard\\intbl ${a}\\cell}{\\pard\\intbl ${r}\\cell}\\row\n`;
      }
      return rows;
    }

    const blob = new Blob([rtf], { type: 'application/rtf' });
    const url  = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href     = url;
    link.download = `${name}_${week}_OJT_Report.doc`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    hideLoader();
    showToast('DOCX/DOC downloaded!', 'success');
  } catch (err) {
    hideLoader();
    showToast('DOCX export failed.', 'error');
    console.error(err);
  }
}

function rtfEsc(str) {
  if (!str) return '';
  return str.replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}')
            .replace(/\n/g, '\\par ');
}

// ══════════════════════════════════════════════════
//  SETUP (Apps Script URL)
// ══════════════════════════════════════════════════
function saveScriptUrl() {
  const url = document.getElementById('scriptUrl').value.trim();
  if (!url.startsWith('https://script.google.com')) {
    showToast('Invalid URL. Must be a Google Apps Script Web App URL.', 'error');
    return;
  }
  SCRIPT_URL = url;
  localStorage.setItem('ojt_script_url', SCRIPT_URL);
  showToast('✅ Connected to Google Apps Script!', 'success');
}

function showSetupGuide(e) {
  if (e) e.preventDefault();
  openModal('setup');
}

// ══════════════════════════════════════════════════
//  TOAST
// ══════════════════════════════════════════════════
function showToast(msg, type = 'info') {
  const icons = { success: '✅', error: '❌', info: 'ℹ️', warning: '⚠️' };
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.textContent = `${icons[type] || ''} ${msg}`;
  document.getElementById('toastContainer').appendChild(el);
  setTimeout(() => { el.classList.add('fade-out'); setTimeout(() => el.remove(), 350); }, 3500);
}

// ══════════════════════════════════════════════════
//  LOADER
// ══════════════════════════════════════════════════
function showLoader(msg = 'Loading…') {
  document.getElementById('loaderMsg').textContent = msg;
  document.getElementById('loader').classList.remove('hidden');
}
function hideLoader() {
  document.getElementById('loader').classList.add('hidden');
}

// ══════════════════════════════════════════════════
//  UTILITIES
// ══════════════════════════════════════════════════
function cap(str) { return str.charAt(0).toUpperCase() + str.slice(1); }

function escHtml(str) {
  if (!str) return '';
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function fmtDate(dateStr) {
  if (!dateStr) return '';
  try {
    return new Date(dateStr + 'T00:00:00').toLocaleDateString('en-PH', { month:'short', day:'numeric' });
  } catch { return dateStr; }
}

function safeJSON(val, fallback) {
  if (Array.isArray(val)) return val;
  if (typeof val === 'string') {
    try { return JSON.parse(val); } catch { return fallback; }
  }
  return fallback;
}

function shake(inputId) {
  const el = document.getElementById(inputId);
  if (!el) return;
  el.style.animation = 'none';
  el.offsetHeight; // reflow
  el.style.animation = 'shake .4s ease';
  el.addEventListener('animationend', () => { el.style.animation = ''; }, { once: true });
}

// Add shake animation to CSS dynamically
const shakeStyle = document.createElement('style');
shakeStyle.textContent = `@keyframes shake { 0%,100%{transform:translateX(0)} 20%{transform:translateX(-6px)} 40%{transform:translateX(6px)} 60%{transform:translateX(-4px)} 80%{transform:translateX(4px)} }`;
document.head.appendChild(shakeStyle);
