"""
app.py  —  Fabrication BOM Tool (Web)
Flask entry point. All routes live here.
Run with:  python app.py
Then open: http://localhost:5000
"""

import os
import json
import shutil
from pathlib import Path
from datetime import datetime

from flask import (
    Flask, render_template, request, jsonify,
    send_file, abort
)

from core.logic import (
    APP_NAME, APP_VERSION, FITTING_TYPES, MATERIAL_TYPE_PRESET,
    load_settings, save_settings, ensure_company_defaults,
    run_bom, EXPORTS_DIR, _clean_multiplier_table,
    dedupe_case_insensitive_keep_first,
)
from core.history import save_run, get_history, get_run, delete_run

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 128 * 1024 * 1024  # 128 MB upload limit

UPLOAD_DIR = Path(__file__).parent / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
EXPORTS_DIR.mkdir(parents=True, exist_ok=True)


# ──────────────────────────────────────────────────────────────
# Main page
# ──────────────────────────────────────────────────────────────

@app.route("/")
def index():
    settings = load_settings()
    return render_template(
        "index.html",
        app_name=APP_NAME,
        app_version=APP_VERSION,
        fitting_types=FITTING_TYPES,
        material_types=settings.get("material_types", MATERIAL_TYPE_PRESET),
    )


# ──────────────────────────────────────────────────────────────
# Settings API
# ──────────────────────────────────────────────────────────────

@app.route("/api/settings", methods=["GET"])
def api_get_settings():
    s = load_settings()
    # Don't send the password over the wire
    safe = {k: v for k, v in s.items() if k != "admin_password"}
    safe["fitting_types"] = FITTING_TYPES
    return jsonify(safe)


@app.route("/api/settings", methods=["POST"])
def api_save_settings():
    data = request.get_json(force=True)
    if not data:
        return jsonify({"ok": False, "error": "No data received"}), 400

    s = load_settings()

    # Material types
    if "material_types" in data:
        mats = [str(m).strip() for m in data["material_types"] if str(m).strip()]
        if mats:
            s["material_types"] = dedupe_case_insensitive_keep_first(mats)

    # Company multipliers
    if "company_side_multipliers" in data:
        cleaned = _clean_multiplier_table(data["company_side_multipliers"])
        if cleaned:
            s["company_side_multipliers"] = cleaned

    # Project multipliers
    if "project_side_multipliers" in data:
        psm = {}
        for proj, tbl in data["project_side_multipliers"].items():
            proj_name = str(proj).strip()
            if not proj_name:
                continue
            cleaned = _clean_multiplier_table(tbl)
            psm[proj_name] = cleaned
        s["project_side_multipliers"] = psm

    # Exclusions
    if "exclude_fitting_types" in data:
        ex = {str(k): bool(v) for k, v in data["exclude_fitting_types"].items()}
        s["exclude_fitting_types"] = ex

    # Mode
    if "multiplier_mode" in data:
        mode = str(data["multiplier_mode"])
        if mode in ("Company", "Project"):
            s["multiplier_mode"] = mode

    s = ensure_company_defaults(s)
    save_settings(s)
    return jsonify({"ok": True})


# ──────────────────────────────────────────────────────────────
# Admin API
# ──────────────────────────────────────────────────────────────

@app.route("/api/admin/verify", methods=["POST"])
def api_admin_verify():
    data = request.get_json(force=True)
    pw   = str(data.get("password", "")).strip()
    s    = load_settings()
    if pw == s.get("admin_password", "FBT2026!"):
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Incorrect password"}), 401


@app.route("/api/admin/change-password", methods=["POST"])
def api_admin_change_password():
    data    = request.get_json(force=True)
    current = str(data.get("current", "")).strip()
    new_pw  = str(data.get("new_password", "")).strip()
    s       = load_settings()
    if current != s.get("admin_password", "FBT2026!"):
        return jsonify({"ok": False, "error": "Current password incorrect"}), 401
    if len(new_pw) < 4:
        return jsonify({"ok": False, "error": "Password must be at least 4 characters"}), 400
    s["admin_password"] = new_pw
    save_settings(s)
    return jsonify({"ok": True})


@app.route("/api/admin/reset-settings", methods=["POST"])
def api_admin_reset():
    data = request.get_json(force=True)
    pw   = str(data.get("password", "")).strip()
    s    = load_settings()
    if pw != s.get("admin_password", "FBT2026!"):
        return jsonify({"ok": False, "error": "Incorrect password"}), 401
    # Preserve password, reset everything else
    from core.logic import DEFAULT_SETTINGS
    new_s = json.loads(json.dumps(DEFAULT_SETTINGS))
    new_s["admin_password"] = pw
    new_s = ensure_company_defaults(new_s)
    save_settings(new_s)
    return jsonify({"ok": True})


# ──────────────────────────────────────────────────────────────
# Run BOM
# ──────────────────────────────────────────────────────────────

@app.route("/api/run", methods=["POST"])
def api_run():
    # Accept multipart/form-data: files + JSON options
    files = request.files.getlist("pdfs")
    if not files or all(f.filename == "" for f in files):
        return jsonify({"ok": False, "error": "No PDF files uploaded"}), 400

    mode            = request.form.get("mode", "Company")
    project         = request.form.get("project", "").strip()
    export_filename = request.form.get("export_filename", "BOM_Export").strip()
    skip_unclassified = request.form.get("skip_unclassified", "false").lower() == "true"

    # Save uploaded PDFs to temp upload dir
    saved_paths = []
    for f in files:
        if not f.filename.lower().endswith(".pdf"):
            continue
        safe_name = Path(f.filename).name
        dest = UPLOAD_DIR / safe_name
        f.save(str(dest))
        saved_paths.append(dest)

    if not saved_paths:
        return jsonify({"ok": False, "error": "No valid PDF files received"}), 400

    settings = load_settings()

    # Apply skip_unclassified option on-the-fly
    if skip_unclassified:
        settings["exclude_fitting_types"]["Unclassified"] = True

    try:
        result = run_bom(
            pdf_paths=[str(p) for p in saved_paths],
            settings=settings,
            export_filename=export_filename,
            mode=mode,
            project=project if mode == "Project" else "",
        )
    except Exception as e:
        return jsonify({"ok": False, "error": f"{type(e).__name__}: {e}"}), 500
    finally:
        # Clean up uploaded temp files
        for p in saved_paths:
            try:
                p.unlink()
            except Exception:
                pass

    # Save to history
    try:
        run_id = save_run(result, saved_paths, mode, project)
        result["run_id"] = run_id
    except Exception:
        result["run_id"] = None

    return jsonify(result)


# ──────────────────────────────────────────────────────────────
# Download exported Excel
# ──────────────────────────────────────────────────────────────

@app.route("/api/download/<path:filename>")
def api_download(filename):
    safe = Path(filename).name  # prevent path traversal
    path = EXPORTS_DIR / safe
    if not path.exists():
        abort(404)
    return send_file(str(path), as_attachment=True, download_name=safe)


# ──────────────────────────────────────────────────────────────
# History API
# ──────────────────────────────────────────────────────────────

@app.route("/api/history")
def api_history():
    return jsonify(get_history(50))


@app.route("/api/history/<int:run_id>")
def api_history_detail(run_id):
    run = get_run(run_id)
    if not run:
        abort(404)
    return jsonify(run)


@app.route("/api/history/<int:run_id>", methods=["DELETE"])
def api_history_delete(run_id):
    delete_run(run_id)
    return jsonify({"ok": True})


# ──────────────────────────────────────────────────────────────
# Start server
# ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    from core.logic import DATA_DIR
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    print(f"\n  {'='*44}")
    print(f"  {APP_NAME}  v{APP_VERSION}")
    print(f"  {'='*44}")
    print(f"  Open in browser: http://localhost:5000")
    print(f"  Press Ctrl+C to stop\n")
    app.run(debug=True, host="0.0.0.0", port=5000)
