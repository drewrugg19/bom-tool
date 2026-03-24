/* Fabrication BOM Tool — app.js */
"use strict";

let selectedFiles = [];
let settings = {};
let adminUnlocked = false;
let currentMat = "";
let currentProjMat = "";
let adminIdleTimer = null;

const ADMIN_IDLE_TIMEOUT_MS = 2 * 60 * 1000;

const TABS = {
  run:      { title: "Run",      sub: "Upload PDFs and process a batch" },
  history:  { title: "History",  sub: "Past exports and run logs" },
  settings: { title: "Settings", sub: "Manage multipliers, materials, and exclusions" },
  admin:    { title: "Admin",    sub: "Password, imports/exports, and legend management" },
};

function switchTab(name) {
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".nav-item").forEach(n => n.classList.remove("active"));
  document.getElementById("tab-" + name).classList.add("active");
  document.querySelector(`.nav-item[data-tab="${name}"]`).classList.add("active");
  document.getElementById("page-title").textContent = TABS[name].title;
  document.getElementById("page-sub").textContent = TABS[name].sub;

  if (name === "history") loadHistory();
  if (name === "settings") loadSettingsUI();
  if (name === "admin") checkLegendStatus();
  syncProtectedPanels();
}

document.querySelectorAll(".nav-item[data-tab]").forEach(el => {
  el.addEventListener("click", () => switchTab(el.dataset.tab));
});

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const fileListEl = document.getElementById("file-list");
const runBtn = document.getElementById("run-btn");
const projectNameSelect = document.getElementById("project-name");
const projectEmptyMsg = document.getElementById("project-empty-msg");
const modeSelect = document.getElementById("mult-mode");

document.getElementById("browse-trigger").addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => {
  addFiles([...fileInput.files]);
  fileInput.value = "";
});

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("dragover", e => {
  e.preventDefault();
  dropzone.classList.add("over");
});
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("over"));
dropzone.addEventListener("drop", e => {
  e.preventDefault();
  dropzone.classList.remove("over");
  addFiles([...e.dataTransfer.files]);
});

function addFiles(files) {
  const pdfs = files.filter(f => f.name.toLowerCase().endsWith(".pdf"));
  pdfs.forEach(f => {
    if (!selectedFiles.find(s => s.name === f.name && s.size === f.size)) {
      selectedFiles.push(f);
    }
  });
  renderFileList();
}

function removeFile(name) {
  selectedFiles = selectedFiles.filter(f => f.name !== name);
  renderFileList();
}

function renderFileList() {
  fileListEl.innerHTML = "";
  selectedFiles.forEach(f => {
    const div = document.createElement("div");
    div.className = "file-item";
    div.innerHTML = `<div class="fdot"></div><div class="fname">${escHtml(f.name)}</div><span class="tag tag-blue">ready</span><div class="frem" data-name="${escHtml(f.name)}">×</div>`;
    fileListEl.appendChild(div);
  });
  fileListEl.querySelectorAll(".frem").forEach(b => b.addEventListener("click", () => removeFile(b.dataset.name)));
  updateRunButtonState();
}

modeSelect.addEventListener("change", syncRunProjectUI);
projectNameSelect.addEventListener("change", updateRunButtonState);

document.querySelectorAll(".tgl").forEach(t => {
  t.addEventListener("click", () => {
    t.classList.toggle("on");
    t.classList.toggle("off");
  });
});

function updateRunButtonState() {
  const projectRequired = modeSelect.value === "Project";
  const projectReady = !projectRequired || projectNameSelect.value.trim().length > 0;
  runBtn.disabled = selectedFiles.length === 0 || !projectReady;
}

runBtn.addEventListener("click", runBOM);

async function runBOM() {
  if (selectedFiles.length === 0) return;

  const mode = modeSelect.value;
  const project = projectNameSelect.value.trim();
  const fname = document.getElementById("export-name").value.trim() || "BOM_Export";
  const skipUC = document.getElementById("tgl-skip").classList.contains("on");

  if (mode === "Project" && !project) {
    document.getElementById("run-err").textContent = "Select a one-off project for Project mode.";
    return;
  }

  document.getElementById("run-err").textContent = "";
  document.getElementById("prog-wrap").style.display = "block";
  document.getElementById("prog-bar").style.width = "40%";
  runBtn.disabled = true;

  const fd = new FormData();
  selectedFiles.forEach(f => fd.append("pdfs", f));
  fd.append("mode", mode);
  fd.append("project", project);
  fd.append("export_filename", fname);
  fd.append("skip_unclassified", skipUC ? "true" : "false");

  try {
    const data = await fetchJson("/api/run", { method: "POST", body: fd });

    document.getElementById("prog-bar").style.width = "100%";
    setTimeout(() => {
      document.getElementById("prog-wrap").style.display = "none";
      document.getElementById("prog-bar").style.width = "0%";
    }, 600);

    document.getElementById("summary-card").style.display = "block";
    document.getElementById("s-inches").textContent = (data.total_inches || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    document.getElementById("s-rows").textContent = data.rows || 0;
    document.getElementById("s-warn").textContent = data.warn_rows || 0;
    document.getElementById("s-err").textContent = data.err_rows || 0;
    document.getElementById("summary-text").textContent = data.summary || "";

    const dlWrap = document.getElementById("dl-wrap");
    dlWrap.innerHTML = "";
    if (data.output_filename) {
      const btn = document.createElement("a");
      btn.href = `/api/download/${encodeURIComponent(data.output_filename)}`;
      btn.download = data.output_filename;
      btn.className = "btn btn-primary";
      btn.textContent = `⬇ Download ${data.output_filename}`;
      dlWrap.appendChild(btn);
    }

    showToast("BOM export complete");
  } catch (e) {
    document.getElementById("run-err").textContent = e.message;
    document.getElementById("prog-wrap").style.display = "none";
  } finally {
    updateRunButtonState();
  }
}

async function loadHistory() {
  const body = document.getElementById("history-body");
  body.innerHTML = '<p class="empty">Loading…</p>';
  try {
    const rows = await fetchJson("/api/history");
    if (!rows.length) {
      body.innerHTML = '<p class="empty">No runs yet. Completed runs will appear here.</p>';
      return;
    }
    const tbl = document.createElement("table");
    tbl.className = "hist-table";
    tbl.innerHTML = "<thead><tr><th>Date</th><th>File</th><th>PDFs</th><th>Rows</th><th>Total inches</th><th>Status</th><th></th></tr></thead>";
    const tbody = document.createElement("tbody");
    rows.forEach(r => {
      const warnTag = r.warn_rows > 0 ? `<span class="tag tag-amber">${r.warn_rows} warn</span>` : "";
      const errTag = r.err_rows > 0 ? `<span class="tag tag-red">${r.err_rows} err</span>` : "";
      const okTag = (!r.warn_rows && !r.err_rows) ? '<span class="tag tag-green">OK</span>' : "";
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td style="color:var(--tx3)">${escHtml(r.run_date || "")}</td>
        <td>${escHtml(r.export_file || "")}</td>
        <td>${r.pdf_count || 0}</td>
        <td>${r.row_count || 0}</td>
        <td style="font-weight:600;color:var(--blue)">${(r.total_inches || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td>${okTag}${warnTag}${errTag}</td>
        <td style="display:flex;gap:6px;flex-wrap:wrap">
          ${r.export_file ? `<a class="btn" href="/api/download/${encodeURIComponent(r.export_file)}" download="${escHtml(r.export_file)}" style="font-size:11px;padding:3px 10px">Download</a>` : ""}
          <button class="btn" style="font-size:11px;padding:3px 10px;border-color:#E24B4A;color:#A32D2D" onclick="deleteRun(${r.id},this)">Delete</button>
        </td>`;
      tbody.appendChild(tr);
    });
    tbl.appendChild(tbody);
    body.innerHTML = "";
    body.appendChild(tbl);
  } catch (e) {
    body.innerHTML = `<p class="empty" style="color:var(--red)">Failed to load history: ${escHtml(e.message)}</p>`;
  }
}

async function deleteRun(id, btn) {
  if (!confirm("Delete this run from history?")) return;
  btn.disabled = true;
  try {
    await fetchJson(`/api/history/${id}`, { method: "DELETE" });
    showToast("Run deleted");
    loadHistory();
  } catch (e) {
    btn.disabled = false;
    alert(e.message);
  }
}


async function loadSettingsUI() {
  settings = await fetchJson("/api/settings");
  sortMaterialTypesInState();
  renderMaterialSel();
  renderMaterials();
  renderExclusions();
  renderProjectSel();
  renderRunProjectOptions();
  syncProtectedPanels();
}

document.querySelectorAll(".snav").forEach(item => {
  item.addEventListener("click", () => {
    document.querySelectorAll(".snav").forEach(i => i.classList.remove("active"));
    document.querySelectorAll(".spanel").forEach(p => p.classList.remove("active"));
    item.classList.add("active");
    document.getElementById("panel-" + item.dataset.panel).classList.add("active");
  });
});

function renderMaterialSel() {
  const sel = document.getElementById("mult-mat-sel");
  sel.innerHTML = "";
  const mats = sortedMaterialTypes();
  mats.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    sel.appendChild(opt);
  });
  currentMat = mats.includes(currentMat) ? currentMat : (mats[0] || "");
  sel.value = currentMat;
  sel.onchange = () => {
    currentMat = sel.value;
    renderMultGrid(currentMat);
  };
  renderMultGrid(currentMat);
}

function renderMultGrid(mat) {
  const grid = document.getElementById("mult-grid");
  grid.innerHTML = "";
  const fittingTypes = settings.fitting_types || window.FBT.fittingTypes;
  const csm = settings.company_side_multipliers || {};
  const matRow = csm[mat] || {};
  fittingTypes.forEach(ft => {
    const cell = document.createElement("div");
    cell.className = "mult-cell";
    const val = matRow[ft] !== undefined ? matRow[ft] : 2.0;
    cell.innerHTML = `<div class="mult-cell-lbl">${escHtml(ft)}</div><input class="mult-inp" type="number" step="0.01" min="0" max="10" value="${val}" data-ft="${escHtml(ft)}" />`;
    grid.appendChild(cell);
  });
}

async function saveMultipliers() {
  if (!currentMat) return;
  if (!settings.company_side_multipliers) settings.company_side_multipliers = {};
  if (!settings.company_side_multipliers[currentMat]) settings.company_side_multipliers[currentMat] = {};
  document.querySelectorAll("#mult-grid .mult-inp").forEach(inp => {
    const ft = inp.dataset.ft;
    const val = parseFloat(inp.value);
    if (!Number.isNaN(val)) settings.company_side_multipliers[currentMat][ft] = val;
  });
  await postSettings({ company_side_multipliers: settings.company_side_multipliers });
  showToast("Multipliers saved");
}

function renderMaterials() {
  const list = document.getElementById("chip-list");
  list.innerHTML = "";
  sortedMaterialTypes().forEach(m => {
    const chip = document.createElement("div");
    chip.className = "chip";
    chip.innerHTML = `${escHtml(m)}<span class="chip-x" data-mat="${escHtml(m)}">×</span>`;
    list.appendChild(chip);
  });
  list.querySelectorAll(".chip-x").forEach(b => b.addEventListener("click", () => {
    settings.material_types = (settings.material_types || []).filter(x => x !== b.dataset.mat);
    sortMaterialTypesInState();
    renderMaterials();
    renderMaterialSel();
    renderProjMatSel(document.getElementById("proj-sel").value);
  }));
}

function addMaterial() {
  const inp = document.getElementById("new-mat");
  const val = inp.value.trim();
  if (!val) return;
  if (!settings.material_types) settings.material_types = [];
  if (!settings.material_types.find(m => m.toLowerCase() === val.toLowerCase())) {
    settings.material_types.push(val);
    sortMaterialTypesInState();
    renderMaterials();
    renderMaterialSel();
    renderProjMatSel(document.getElementById("proj-sel").value);
  }
  inp.value = "";
}
document.getElementById("new-mat").addEventListener("keydown", e => {
  if (e.key === "Enter") addMaterial();
});

async function saveMaterials() {
  await postSettings({ material_types: settings.material_types });
  showToast("Materials saved");
}

function renderExclusions() {
  const grid = document.getElementById("excl-grid");
  grid.innerHTML = "";
  const ex = settings.exclude_fitting_types || {};
  const fts = settings.fitting_types || window.FBT.fittingTypes;
  fts.forEach(ft => {
    const div = document.createElement("div");
    div.className = "excl-item";
    const checked = ex[ft] ? "checked" : "";
    div.innerHTML = `<input type="checkbox" id="ex-${escHtml(ft)}" ${checked} /><label for="ex-${escHtml(ft)}" style="cursor:pointer">${escHtml(ft)}</label>`;
    grid.appendChild(div);
  });
}

async function saveExclusions() {
  const ex = {};
  document.querySelectorAll("#excl-grid input[type=checkbox]").forEach(cb => {
    ex[cb.id.replace("ex-", "")] = cb.checked;
  });
  settings.exclude_fitting_types = ex;
  await postSettings({ exclude_fitting_types: ex });
  showToast("Exclusions saved");
}

function renderProjectSel() {
  const sel = document.getElementById("proj-sel");
  const prevSelection = sel.value;
  sel.innerHTML = '<option value="">— select project —</option>';
  const psm = settings.project_side_multipliers || {};
  const names = Object.keys(psm).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  names.forEach(p => {
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;
    sel.appendChild(opt);
  });
  sel.value = names.includes(prevSelection) ? prevSelection : "";
  sel.onchange = () => {
    renderProjMatSel(sel.value);
    renderProjectList();
  };
  renderProjectList();
  renderProjMatSel(sel.value);
  renderRunProjectOptions();
}

function renderProjMatSel(proj) {
  if (!proj) {
    document.getElementById("proj-mult-grid").innerHTML = "";
    return;
  }
  const sel = document.getElementById("proj-mat-sel");
  sel.innerHTML = "";
  const mats = sortedMaterialTypes();
  mats.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m;
    opt.textContent = m;
    sel.appendChild(opt);
  });
  currentProjMat = mats.includes(currentProjMat) ? currentProjMat : (mats[0] || "");
  sel.value = currentProjMat;
  sel.onchange = () => {
    currentProjMat = sel.value;
    renderProjMultGrid(proj, currentProjMat);
  };
  renderProjMultGrid(proj, currentProjMat);
}

function renderProjMultGrid(proj, mat) {
  const grid = document.getElementById("proj-mult-grid");
  grid.innerHTML = "";
  if (!proj || !mat) return;
  const fts = settings.fitting_types || window.FBT.fittingTypes;
  const psm = settings.project_side_multipliers || {};
  const matRow = (psm[proj] && psm[proj][mat]) || {};
  fts.forEach(ft => {
    const cell = document.createElement("div");
    cell.className = "mult-cell";
    const val = matRow[ft] !== undefined ? matRow[ft] : 2.0;
    cell.innerHTML = `<div class="mult-cell-lbl">${escHtml(ft)}</div><input class="mult-inp" type="number" step="0.01" min="0" max="10" value="${val}" data-ft="${escHtml(ft)}" />`;
    grid.appendChild(cell);
  });
}

function addProject() {
  const name = prompt("New one-off project name:")?.trim();
  if (!name) return;
  if (!settings.project_side_multipliers) settings.project_side_multipliers = {};
  const existingName = Object.keys(settings.project_side_multipliers).find(p => p.toLowerCase() === name.toLowerCase());
  if (existingName) {
    alert("A one-off project with that name already exists.");
    return;
  }
  settings.project_side_multipliers[name] = buildDefaultProjectTable();
  renderProjectSel();
  document.getElementById("proj-sel").value = name;
  renderProjMatSel(name);
  renderProjectList();
}

function buildDefaultProjectTable() {
  const projectTable = {};
  const fts = settings.fitting_types || window.FBT.fittingTypes;
  sortedMaterialTypes().forEach(mat => {
    projectTable[mat] = {};
    fts.forEach(ft => {
      projectTable[mat][ft] = 2.0;
    });
  });
  return projectTable;
}

function renderProjectList() {
  const list = document.getElementById("oneoff-project-list");
  if (!list) return;
  const current = document.getElementById("proj-sel").value;
  const names = getSortedProjectNames();
  if (!names.length) {
    list.innerHTML = '<p class="hint" style="margin:0">No one-off projects saved yet. Use "Add New Project" to create your first one.</p>';
    return;
  }
  list.innerHTML = "";
  names.forEach(name => {
    const row = document.createElement("div");
    row.className = "project-item";
    row.innerHTML = `
      <div class="project-name">${escHtml(name)}</div>
      <div class="project-actions">
        <button class="btn" data-action="edit" data-project="${escHtml(name)}">Edit</button>
        <button class="btn btn-danger" data-action="delete" data-project="${escHtml(name)}">Delete</button>
      </div>`;
    if (name === current) row.style.borderColor = "var(--blue)";
    list.appendChild(row);
  });
  list.querySelectorAll("button").forEach(btn => {
    const project = btn.dataset.project;
    if (btn.dataset.action === "edit") {
      btn.addEventListener("click", () => selectProjectForEditing(project));
    } else {
      btn.addEventListener("click", () => deleteProject(project));
    }
  });
}

function selectProjectForEditing(projectName) {
  const sel = document.getElementById("proj-sel");
  sel.value = projectName;
  renderProjMatSel(projectName);
  renderProjectList();
}

function deleteProject(projectName) {
  if (!confirm(`Delete one-off project "${projectName}"? This cannot be undone.`)) return;
  const psm = settings.project_side_multipliers || {};
  delete psm[projectName];
  settings.project_side_multipliers = psm;
  renderProjectSel();
  showToast(`Deleted ${projectName}`);
}

function getSortedProjectNames() {
  return Object.keys(settings.project_side_multipliers || {})
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function renderRunProjectOptions() {
  const currentValue = projectNameSelect.value;
  const names = getSortedProjectNames();
  projectNameSelect.innerHTML = "";
  if (!names.length) {
    const empty = document.createElement("option");
    empty.value = "";
    empty.textContent = "— no one-off projects available —";
    projectNameSelect.appendChild(empty);
  } else {
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "— select project —";
    projectNameSelect.appendChild(placeholder);
    names.forEach(name => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      projectNameSelect.appendChild(opt);
    });
  }
  projectNameSelect.value = names.includes(currentValue) ? currentValue : "";
  syncRunProjectUI();
}

async function saveProjectMultipliers() {
  const proj = document.getElementById("proj-sel").value;
  if (!proj) {
    alert("Select a project first.");
    return;
  }
  if (!currentProjMat) {
    alert("Add at least one material type first.");
    return;
  }
  if (!settings.project_side_multipliers) settings.project_side_multipliers = {};
  if (!settings.project_side_multipliers[proj]) settings.project_side_multipliers[proj] = {};
  if (!settings.project_side_multipliers[proj][currentProjMat]) settings.project_side_multipliers[proj][currentProjMat] = {};
  document.querySelectorAll("#proj-mult-grid .mult-inp").forEach(inp => {
    const ft = inp.dataset.ft;
    const val = parseFloat(inp.value);
    if (!Number.isNaN(val)) settings.project_side_multipliers[proj][currentProjMat][ft] = val;
  });
  await postSettings({ project_side_multipliers: settings.project_side_multipliers });
  showToast("Project multipliers saved");
}


async function unlockProtectedArea() {
  const adminInput = document.getElementById("admin-pw");
  const settingsInput = document.getElementById("settings-pw");
  const activeInput = document.activeElement === settingsInput ? settingsInput : adminInput;
  const password = (activeInput?.value || adminInput.value || settingsInput.value || "").trim();
  const adminErr = document.getElementById("lock-err");
  const settingsErr = document.getElementById("settings-lock-err");
  adminErr.textContent = "";
  settingsErr.textContent = "";
  if (!password) {
    const message = "Enter your admin password.";
    adminErr.textContent = message;
    settingsErr.textContent = message;
    return;
  }
  try {
    await fetchJson("/api/admin/verify", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password }),
    });
    adminUnlocked = true;
    adminInput.value = "";
    settingsInput.value = "";
    syncProtectedPanels();
    updateAdminStatus();
    resetAdminIdleTimer();
  } catch (e) {
    adminErr.textContent = e.message;
    settingsErr.textContent = e.message;
  }
}
document.getElementById("admin-pw").addEventListener("keydown", e => {
  if (e.key === "Enter") unlockProtectedArea();
});
document.getElementById("settings-pw").addEventListener("keydown", e => {
  if (e.key === "Enter") unlockProtectedArea();
});

async function changePassword() {
  const cur = document.getElementById("cur-pw").value;
  const np1 = document.getElementById("new-pw").value;
  const np2 = document.getElementById("new-pw2").value;
  const err = document.getElementById("pw-err");
  err.textContent = "";
  if (!cur || !np1 || !np2) {
    err.textContent = "Fill in all fields.";
    return;
  }
  if (np1 !== np2) {
    err.textContent = "New passwords do not match.";
    return;
  }
  try {
    await fetchJson("/api/admin/change-password", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ current: cur, new_password: np1 }),
    });
    showToast("Password updated");
    resetAdminIdleTimer();
    document.getElementById("cur-pw").value = "";
    document.getElementById("new-pw").value = "";
    document.getElementById("new-pw2").value = "";
  } catch (e) {
    err.textContent = e.message;
  }
}

async function resetSettings() {
  const pw = prompt("Enter admin password to confirm reset:");
  if (!pw) return;
  if (!confirm("This will reset ALL multipliers, material types, and exclusions. Run history is preserved. Continue?")) return;
  try {
    await fetchJson("/api/admin/reset-settings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ password: pw }),
    });
    showToast("Settings reset to defaults");
    resetAdminIdleTimer();
    loadSettingsUI();
  } catch (e) {
    alert(e.message);
  }
}

async function exportSettings() {
  resetAdminIdleTimer();
  const data = await fetchJson("/api/settings");
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "FBT_settings.json";
  a.click();
  URL.revokeObjectURL(url);
}

async function importSettings(input) {
  const file = input.files[0];
  if (!file) return;
  if (!confirm("This will overwrite your current settings. Continue?")) {
    input.value = "";
    return;
  }
  const text = await file.text();
  try {
    const data = JSON.parse(text);
    await postSettings(data);
    showToast("Settings imported");
    resetAdminIdleTimer();
    loadSettingsUI();
  } catch (e) {
    alert("Invalid JSON file: " + e.message);
  }
  input.value = "";
}

async function checkLegendStatus() {
  const el = document.getElementById("legend-status");
  try {
    const data = await fetchJson("/health");
    if (data.legend_cache_present) {
      el.textContent = "Legend cache is present and will be used for fitting classification.";
      el.style.color = "var(--green)";
    } else {
      el.textContent = "Legend cache not found. Upload Legend.cache.json or provide Legend.xlsx.";
      el.style.color = "var(--amber)";
    }
  } catch (e) {
    el.textContent = "Could not check legend status.";
    el.style.color = "var(--red)";
  }
}

async function uploadLegend(input) {
  const file = input.files[0];
  if (!file) return;
  const err = document.getElementById("legend-err");
  err.textContent = "";
  try {
    const text = await file.text();
    const payload = JSON.parse(text);
    await fetchJson("/api/admin/upload-legend", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    showToast("Legend cache uploaded");
    resetAdminIdleTimer();
    checkLegendStatus();
  } catch (e) {
    err.textContent = e.message;
  }
  input.value = "";
}

async function postSettings(payload) {
  return fetchJson("/api/settings", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  let data = null;
  try {
    data = await response.json();
  } catch (e) {
    if (!response.ok) throw new Error(`Request failed (${response.status})`);
    return {};
  }
  if (!response.ok || (data && data.ok === false)) {
    throw new Error((data && data.error) || `Request failed (${response.status})`);
  }
  return data;
}


function sortedMaterialTypes() {
  return [...(settings.material_types || [])].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function sortMaterialTypesInState() {
  settings.material_types = sortedMaterialTypes();
}

function updateAdminStatus() {
  const status = document.getElementById("admin-status");
  if (adminUnlocked) {
    status.innerHTML = '<svg width="11" height="11" viewBox="0 0 16 16" fill="none" stroke="var(--green)" stroke-width="1.5"><rect x="3" y="7" width="10" height="8" rx="1.5"/><path d="M5.5 7V5a2.5 2.5 0 015 0v2"/></svg> Admin unlocked';
  } else {
    status.innerHTML = '<svg width="11" height="11" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="7" width="10" height="8" rx="1.5"/><path d="M5.5 7V5a2.5 2.5 0 015 0v2"/></svg> Admin locked';
  }
}

function syncProtectedPanels() {
  const adminLock = document.getElementById("admin-lock");
  const settingsLock = document.getElementById("settings-lock");
  adminLock.style.display = adminUnlocked ? "none" : "flex";
  settingsLock.style.display = adminUnlocked ? "none" : "flex";
  updateAdminStatus();
}

function lockAdminAccess(showToastMessage = false) {
  adminUnlocked = false;
  clearTimeout(adminIdleTimer);
  adminIdleTimer = null;
  document.getElementById("admin-pw").value = "";
  document.getElementById("settings-pw").value = "";
  document.getElementById("lock-err").textContent = "";
  document.getElementById("settings-lock-err").textContent = "";
  syncProtectedPanels();
  if (showToastMessage) {
    showToast("Admin access locked");
  }
}

function resetAdminIdleTimer() {
  if (!adminUnlocked) return;
  clearTimeout(adminIdleTimer);
  adminIdleTimer = window.setTimeout(() => lockAdminAccess(true), ADMIN_IDLE_TIMEOUT_MS);
}

["click", "keydown", "mousemove", "mousedown", "touchstart", "scroll"].forEach(eventName => {
  document.addEventListener(eventName, () => {
    if (adminUnlocked) resetAdminIdleTimer();
  }, { passive: true });
});

function showToast(msg) {
  let t = document.getElementById("toast");
  if (!t) {
    t = document.createElement("div");
    t.id = "toast";
    t.style.cssText = "position:fixed;bottom:24px;right:24px;padding:10px 18px;border-radius:8px;font-size:13px;z-index:9999;opacity:0;transition:opacity .2s;pointer-events:none";
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.style.opacity = "1";
  clearTimeout(t._timer);
  t._timer = setTimeout(() => {
    t.style.opacity = "0";
  }, 2500);
}

function escHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}


(async function init() {
  settings = await fetchJson("/api/settings");
  sortMaterialTypesInState();
  renderRunProjectOptions();
  syncProtectedPanels();
  updateRunButtonState();
})();

function syncRunProjectUI() {
  const showProject = modeSelect.value === "Project";
  const names = getSortedProjectNames();
  document.getElementById("project-row").style.display = showProject ? "flex" : "none";
  projectNameSelect.style.display = showProject ? "block" : "none";
  projectEmptyMsg.style.display = showProject && names.length === 0 ? "block" : "none";
  updateRunButtonState();
}
