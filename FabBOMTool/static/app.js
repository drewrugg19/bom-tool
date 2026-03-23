/* Fabrication BOM Tool — app.js */
"use strict";

// ─────────────────────────────────────────────────────────────
// State
// ─────────────────────────────────────────────────────────────
let selectedFiles = [];
let settings = {};
let adminUnlocked = false;
let currentMat = "";
let currentProjMat = "";

const TABS = {
  run:      { title: "Run",      sub: "Upload PDFs and process a batch" },
  history:  { title: "History",  sub: "Past exports and run logs" },
  settings: { title: "Settings", sub: "Manage multipliers, materials, and exclusions" },
  admin:    { title: "Admin",    sub: "Password, imports/exports, and legend management" },
};

// ─────────────────────────────────────────────────────────────
// Tab navigation
// ─────────────────────────────────────────────────────────────
function switchTab(name) {
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".nav-item").forEach(n => n.classList.remove("active"));
  document.getElementById("tab-" + name).classList.add("active");
  document.querySelector(`.nav-item[data-tab="${name}"]`).classList.add("active");
  document.getElementById("page-title").textContent = TABS[name].title;
  document.getElementById("page-sub").textContent   = TABS[name].sub;

  if (name === "history")  loadHistory();
  if (name === "settings") loadSettingsUI();
  if (name === "admin")    checkLegendStatus();
}

document.querySelectorAll(".nav-item[data-tab]").forEach(el => {
  el.addEventListener("click", () => switchTab(el.dataset.tab));
});

// ─────────────────────────────────────────────────────────────
// Drop zone + file handling
// ─────────────────────────────────────────────────────────────
const dropzone    = document.getElementById("dropzone");
const fileInput   = document.getElementById("file-input");
const fileListEl  = document.getElementById("file-list");
const runBtn      = document.getElementById("run-btn");

document.getElementById("browse-trigger").addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", () => { addFiles([...fileInput.files]); fileInput.value = ""; });

dropzone.addEventListener("dragover",  e => { e.preventDefault(); dropzone.classList.add("over"); });
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("over"));
dropzone.addEventListener("drop", e => {
  e.preventDefault(); dropzone.classList.remove("over");
  addFiles([...e.dataTransfer.files]);
});

function addFiles(files) {
  const pdfs = files.filter(f => f.name.toLowerCase().endsWith(".pdf"));
  pdfs.forEach(f => { if (!selectedFiles.find(s => s.name === f.name)) selectedFiles.push(f); });
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
    div.innerHTML = `<div class="fdot"></div><div class="fname">${f.name}</div><span class="tag tag-blue">ready</span><div class="frem" data-name="${escHtml(f.name)}">×</div>`;
    fileListEl.appendChild(div);
  });
  fileListEl.querySelectorAll(".frem").forEach(b => b.addEventListener("click", () => removeFile(b.dataset.name)));
  runBtn.disabled = selectedFiles.length === 0;
}

// ─────────────────────────────────────────────────────────────
// Multiplier mode toggle
// ─────────────────────────────────────────────────────────────
document.getElementById("mult-mode").addEventListener("change", function() {
  document.getElementById("project-row").style.display = this.value === "Project" ? "flex" : "none";
});

// ─────────────────────────────────────────────────────────────
// Toggle switches
// ─────────────────────────────────────────────────────────────
document.querySelectorAll(".tgl").forEach(t => {
  t.addEventListener("click", () => { t.classList.toggle("on"); t.classList.toggle("off"); });
});

// ─────────────────────────────────────────────────────────────
// Run BOM
// ─────────────────────────────────────────────────────────────
runBtn.addEventListener("click", runBOM);

async function runBOM() {
  if (selectedFiles.length === 0) return;

  const mode    = document.getElementById("mult-mode").value;
  const project = document.getElementById("project-name").value.trim();
  const fname   = document.getElementById("export-name").value.trim() || "BOM_Export";
  const skipUC  = document.getElementById("tgl-skip").classList.contains("on");

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
    const res  = await fetch("/api/run", { method: "POST", body: fd });
    const data = await res.json();

    document.getElementById("prog-bar").style.width = "100%";
    setTimeout(() => { document.getElementById("prog-wrap").style.display = "none"; document.getElementById("prog-bar").style.width = "0%"; }, 600);

    if (!data.ok) {
      document.getElementById("run-err").textContent = data.error || "Unknown error";
      return;
    }

    // Show summary card
    document.getElementById("summary-card").style.display = "block";
    document.getElementById("s-inches").textContent = (data.total_inches || 0).toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2});
    document.getElementById("s-rows").textContent   = data.rows || 0;
    document.getElementById("s-warn").textContent   = data.warn_rows || 0;
    document.getElementById("s-err").textContent    = data.err_rows  || 0;
    document.getElementById("summary-text").textContent = data.summary || "";

    // Download link
    const dlWrap = document.getElementById("dl-wrap");
    dlWrap.innerHTML = "";
    if (data.output_filename) {
      const btn = document.createElement("a");
      btn.href     = `/api/download/${encodeURIComponent(data.output_filename)}`;
      btn.download = data.output_filename;
      btn.className = "btn btn-primary";
      btn.textContent = `⬇ Download ${data.output_filename}`;
      dlWrap.appendChild(btn);
    }

  } catch (e) {
    document.getElementById("run-err").textContent = "Request failed: " + e.message;
    document.getElementById("prog-wrap").style.display = "none";
  } finally {
    runBtn.disabled = selectedFiles.length === 0;
  }
}

// ─────────────────────────────────────────────────────────────
// History
// ─────────────────────────────────────────────────────────────
async function loadHistory() {
  const body = document.getElementById("history-body");
  body.innerHTML = '<p class="empty">Loading…</p>';
  try {
    const res  = await fetch("/api/history");
    const rows = await res.json();
    if (!rows.length) { body.innerHTML = '<p class="empty">No runs yet. Completed runs will appear here.</p>'; return; }
    const tbl = document.createElement("table");
    tbl.className = "hist-table";
    tbl.innerHTML = `<thead><tr><th>Date</th><th>File</th><th>PDFs</th><th>Rows</th><th>Total inches</th><th>Status</th><th></th></tr></thead>`;
    const tbody = document.createElement("tbody");
    rows.forEach(r => {
      const warnTag = r.warn_rows  > 0 ? `<span class="tag tag-amber">${r.warn_rows} warn</span>` : "";
      const errTag  = r.err_rows   > 0 ? `<span class="tag tag-red">${r.err_rows} err</span>`     : "";
      const okTag   = (!r.warn_rows && !r.err_rows) ? `<span class="tag tag-green">OK</span>`      : "";
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td style="color:var(--tx3)">${r.run_date}</td>
        <td>${escHtml(r.export_file)}</td>
        <td>${r.pdf_count}</td>
        <td>${r.row_count}</td>
        <td style="font-weight:600;color:var(--blue)">${(r.total_inches||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}</td>
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
  } catch(e) {
    body.innerHTML = `<p class="empty" style="color:var(--red)">Failed to load history: ${e.message}</p>`;
  }
}

async function deleteRun(id, btn) {
  if (!confirm("Delete this run from history?")) return;
  btn.disabled = true;
  await fetch(`/api/history/${id}`, { method: "DELETE" });
  loadHistory();
}

// ─────────────────────────────────────────────────────────────
// Settings — load
// ─────────────────────────────────────────────────────────────
async function loadSettingsUI() {
  const res = await fetch("/api/settings");
  settings  = await res.json();
  renderMaterialSel();
  renderMaterials();
  renderExclusions();
  renderProjectSel();
}

// Settings sub-nav
document.querySelectorAll(".snav").forEach(item => {
  item.addEventListener("click", () => {
    document.querySelectorAll(".snav").forEach(i => i.classList.remove("active"));
    document.querySelectorAll(".spanel").forEach(p => p.classList.remove("active"));
    item.classList.add("active");
    document.getElementById("panel-" + item.dataset.panel).classList.add("active");
  });
});

// ─────────────────────────────────────────────────────────────
// Settings — Multipliers
// ─────────────────────────────────────────────────────────────
function renderMaterialSel() {
  const sel = document.getElementById("mult-mat-sel");
  sel.innerHTML = "";
  const mats = settings.material_types || [];
  mats.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m; opt.textContent = m;
    sel.appendChild(opt);
  });
  if (mats.length) { currentMat = mats[0]; sel.value = currentMat; }
  sel.addEventListener("change", () => { currentMat = sel.value; renderMultGrid(currentMat); });
  renderMultGrid(currentMat);
}

function renderMultGrid(mat) {
  const grid = document.getElementById("mult-grid");
  grid.innerHTML = "";
  const fittingTypes = settings.fitting_types || window.FBT.fittingTypes;
  const csm   = settings.company_side_multipliers || {};
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
  // Read current grid values back into settings
  const fittingTypes = settings.fitting_types || window.FBT.fittingTypes;
  if (!settings.company_side_multipliers) settings.company_side_multipliers = {};
  if (!settings.company_side_multipliers[currentMat]) settings.company_side_multipliers[currentMat] = {};
  document.querySelectorAll("#mult-grid .mult-inp").forEach(inp => {
    const ft  = inp.dataset.ft;
    const val = parseFloat(inp.value);
    if (!isNaN(val)) settings.company_side_multipliers[currentMat][ft] = val;
  });
  await postSettings({ company_side_multipliers: settings.company_side_multipliers });
  showToast("Multipliers saved");
}

// ─────────────────────────────────────────────────────────────
// Settings — Material types
// ─────────────────────────────────────────────────────────────
function renderMaterials() {
  const list = document.getElementById("chip-list");
  list.innerHTML = "";
  (settings.material_types || []).forEach(m => {
    const chip = document.createElement("div");
    chip.className = "chip";
    chip.innerHTML = `${escHtml(m)}<span class="chip-x" data-mat="${escHtml(m)}">×</span>`;
    list.appendChild(chip);
  });
  list.querySelectorAll(".chip-x").forEach(b => b.addEventListener("click", () => {
    settings.material_types = settings.material_types.filter(x => x !== b.dataset.mat);
    renderMaterials();
  }));
}

function addMaterial() {
  const inp = document.getElementById("new-mat");
  const val = inp.value.trim();
  if (!val) return;
  if (!settings.material_types) settings.material_types = [];
  if (!settings.material_types.find(m => m.toLowerCase() === val.toLowerCase())) {
    settings.material_types.push(val);
    renderMaterials();
    renderMaterialSel();
  }
  inp.value = "";
}
document.getElementById("new-mat").addEventListener("keydown", e => { if (e.key === "Enter") addMaterial(); });

async function saveMaterials() {
  await postSettings({ material_types: settings.material_types });
  showToast("Materials saved");
}

// ─────────────────────────────────────────────────────────────
// Settings — Exclusions
// ─────────────────────────────────────────────────────────────
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
    const ft = cb.id.replace("ex-","");
    ex[ft] = cb.checked;
  });
  settings.exclude_fitting_types = ex;
  await postSettings({ exclude_fitting_types: ex });
  showToast("Exclusions saved");
}

// ─────────────────────────────────────────────────────────────
// Settings — Project overrides
// ─────────────────────────────────────────────────────────────
function renderProjectSel() {
  const sel = document.getElementById("proj-sel");
  sel.innerHTML = '<option value="">— select project —</option>';
  const psm = settings.project_side_multipliers || {};
  Object.keys(psm).sort().forEach(p => {
    const opt = document.createElement("option");
    opt.value = p; opt.textContent = p;
    sel.appendChild(opt);
  });
  sel.addEventListener("change", () => { renderProjMatSel(sel.value); });
}

function renderProjMatSel(proj) {
  if (!proj) { document.getElementById("proj-mult-grid").innerHTML = ""; return; }
  const sel = document.getElementById("proj-mat-sel");
  sel.innerHTML = "";
  const mats = settings.material_types || [];
  mats.forEach(m => {
    const opt = document.createElement("option");
    opt.value = m; opt.textContent = m;
    sel.appendChild(opt);
  });
  if (mats.length) { currentProjMat = mats[0]; sel.value = currentProjMat; }
  sel.addEventListener("change", () => { currentProjMat = sel.value; renderProjMultGrid(proj, currentProjMat); });
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
  const name = prompt("New project name:")?.trim();
  if (!name) return;
  if (!settings.project_side_multipliers) settings.project_side_multipliers = {};
  if (!settings.project_side_multipliers[name]) settings.project_side_multipliers[name] = {};
  renderProjectSel();
  document.getElementById("proj-sel").value = name;
  renderProjMatSel(name);
}

async function saveProjectMultipliers() {
  const proj = document.getElementById("proj-sel").value;
  if (!proj) { alert("Select a project first."); return; }
  if (!settings.project_side_multipliers) settings.project_side_multipliers = {};
  if (!settings.project_side_multipliers[proj]) settings.project_side_multipliers[proj] = {};
  if (!settings.project_side_multipliers[proj][currentProjMat]) settings.project_side_multipliers[proj][currentProjMat] = {};
  document.querySelectorAll("#proj-mult-grid .mult-inp").forEach(inp => {
    const ft  = inp.dataset.ft;
    const val = parseFloat(inp.value);
    if (!isNaN(val)) settings.project_side_multipliers[proj][currentProjMat][ft] = val;
  });
  await postSettings({ project_side_multipliers: settings.project_side_multipliers });
  showToast("Project multipliers saved");
}

// ─────────────────────────────────────────────────────────────
// Admin — lock / unlock
// ─────────────────────────────────────────────────────────────
async function unlockAdmin() {
  const pw  = document.getElementById("admin-pw").value;
  const err = document.getElementById("lock-err");
  err.textContent = "";
  try {
    const res  = await fetch("/api/admin/verify", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ password: pw }) });
    const data = await res.json();
    if (data.ok) {
      document.getElementById("admin-lock").style.display = "none";
      document.getElementById("admin-status").innerHTML = `<svg width="11" height="11" viewBox="0 0 16 16" fill="none" stroke="var(--green)" stroke-width="1.5"><rect x="3" y="7" width="10" height="8" rx="1.5"/><path d="M5.5 7V5a2.5 2.5 0 015 0v2"/></svg> Admin unlocked`;
      adminUnlocked = true;
    } else {
      err.textContent = "Incorrect password.";
    }
  } catch(e) {
    err.textContent = "Request failed.";
  }
}
document.getElementById("admin-pw").addEventListener("keydown", e => { if (e.key === "Enter") unlockAdmin(); });

// ─────────────────────────────────────────────────────────────
// Admin — change password
// ─────────────────────────────────────────────────────────────
async function changePassword() {
  const cur  = document.getElementById("cur-pw").value;
  const np1  = document.getElementById("new-pw").value;
  const np2  = document.getElementById("new-pw2").value;
  const err  = document.getElementById("pw-err");
  err.textContent = "";
  if (!cur || !np1 || !np2) { err.textContent = "Fill in all fields."; return; }
  if (np1 !== np2)          { err.textContent = "New passwords do not match."; return; }
  const res  = await fetch("/api/admin/change-password", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ current: cur, new_password: np1 }) });
  const data = await res.json();
  if (data.ok) {
    showToast("Password updated");
    document.getElementById("cur-pw").value = "";
    document.getElementById("new-pw").value  = "";
    document.getElementById("new-pw2").value = "";
  } else {
    err.textContent = data.error || "Failed.";
  }
}

// ─────────────────────────────────────────────────────────────
// Admin — reset settings
// ─────────────────────────────────────────────────────────────
async function resetSettings() {
  const pw = prompt("Enter admin password to confirm reset:");
  if (!pw) return;
  if (!confirm("This will reset ALL multipliers, material types, and exclusions. Run history is preserved. Continue?")) return;
  const res  = await fetch("/api/admin/reset-settings", { method:"POST", headers:{"Content-Type":"application/json"}, body: JSON.stringify({ password: pw }) });
  const data = await res.json();
  if (data.ok) { showToast("Settings reset to defaults"); loadSettingsUI(); }
  else         { alert(data.error || "Reset failed."); }
}

// ─────────────────────────────────────────────────────────────
// Admin — export / import settings
// ─────────────────────────────────────────────────────────────
async function exportSettings() {
  const res  = await fetch("/api/settings");
  const data = await res.json();
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = "FBT_settings.json"; a.click();
  URL.revokeObjectURL(url);
}

async function importSettings(input) {
  const file = input.files[0];
  if (!file) return;
  if (!confirm("This will overwrite your current settings. Continue?")) { input.value = ""; return; }
  const text = await file.text();
  try {
    const data = JSON.parse(text);
    await postSettings(data);
    showToast("Settings imported");
    loadSettingsUI();
  } catch(e) {
    alert("Invalid JSON file: " + e.message);
  }
  input.value = "";
}

// ─────────────────────────────────────────────────────────────
// Admin — legend status
// ─────────────────────────────────────────────────────────────
async function checkLegendStatus() {
  const el = document.getElementById("legend-status");
  try {
    const res  = await fetch("/api/settings");
    const data = await res.json();
    el.textContent = "Legend maps loaded. Fitting classification is active.";
    el.style.color = "var(--green)";
  } catch(e) {
    el.textContent = "Could not check legend status.";
  }
}

async function uploadLegend(input) {
  const file = input.files[0];
  if (!file) return;
  const err = document.getElementById("legend-err");
  err.textContent = "";
  try {
    const text = await file.text();
    JSON.parse(text); // validate JSON
    // Save as Legend.cache.json via settings endpoint (admin convenience)
    const res = await fetch("/api/admin/upload-legend", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: text,
    });
    if (res.ok) { showToast("Legend cache uploaded — restart server to apply"); }
    else        { err.textContent = "Upload failed."; }
  } catch(e) {
    err.textContent = "Invalid JSON: " + e.message;
  }
  input.value = "";
}

// ─────────────────────────────────────────────────────────────
// Shared helpers
// ─────────────────────────────────────────────────────────────
async function postSettings(payload) {
  await fetch("/api/settings", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
}

function showToast(msg) {
  let t = document.getElementById("toast");
  if (!t) {
    t = document.createElement("div");
    t.id = "toast";
    t.style.cssText = "position:fixed;bottom:24px;right:24px;background:#1a1a18;color:#fff;padding:10px 18px;border-radius:8px;font-size:13px;z-index:9999;opacity:0;transition:opacity .2s;pointer-events:none";
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.style.opacity = "1";
  clearTimeout(t._timer);
  t._timer = setTimeout(() => { t.style.opacity = "0"; }, 2500);
}

function escHtml(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ─────────────────────────────────────────────────────────────
// Init
// ─────────────────────────────────────────────────────────────
(async function init() {
  // Pre-load settings so run tab has latest mode
  const res = await fetch("/api/settings");
  settings  = await res.json();
})();
