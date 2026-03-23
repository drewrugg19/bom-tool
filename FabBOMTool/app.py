"""Fabrication BOM Tool Flask entry point."""

from __future__ import annotations

import json
import logging
import os
import re
from pathlib import Path
from uuid import uuid4

from flask import Flask, abort, jsonify, render_template, request, send_file
from werkzeug.exceptions import BadRequest, HTTPException

from .core.history import delete_run, get_history, get_run, save_run
from .core.logic import (
    APP_NAME,
    APP_VERSION,
    DATA_DIR,
    DEFAULT_ADMIN_PASSWORD,
    EXPORTS_DIR,
    FITTING_TYPES,
    LEGEND_CACHE_PATH,
    MATERIAL_TYPE_PRESET,
    _clean_multiplier_table,
    dedupe_case_insensitive_keep_first,
    ensure_company_defaults,
    load_settings,
    password_matches,
    run_bom,
    save_settings,
    set_admin_password,
)

MAX_CONTENT_LENGTH = 128 * 1024 * 1024
DEFAULT_HOST = os.getenv("FBT_HOST", "0.0.0.0")
DEFAULT_PORT = int(os.getenv("FBT_PORT", "5000"))
SAFE_EXPORT_RE = re.compile(r"[^A-Za-z0-9._-]+")


def _bool_env(name: str, default: bool = False) -> bool:
    raw = os.getenv(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def sanitize_export_filename(value: str) -> str:
    cleaned = SAFE_EXPORT_RE.sub("_", str(value or "").strip()).strip("._")
    return cleaned[:120] or "BOM_Export"


def create_app() -> Flask:
    app = Flask(__name__)
    app.config.update(
        APP_NAME=APP_NAME,
        APP_VERSION=APP_VERSION,
        JSON_SORT_KEYS=False,
        MAX_CONTENT_LENGTH=MAX_CONTENT_LENGTH,
    )

    upload_dir = Path(__file__).parent / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    app.logger.setLevel(logging.INFO)

    def json_body(required: bool = True) -> dict:
        data = request.get_json(silent=True)
        if not data:
            if required:
                raise BadRequest("Request body must be valid JSON.")
            return {}
        if not isinstance(data, dict):
            raise BadRequest("JSON payload must be an object.")
        return data

    def safe_settings_payload() -> dict:
        settings = load_settings()
        safe = {
            key: value
            for key, value in settings.items()
            if key not in {"admin_password", "admin_password_hash"}
        }
        safe["fitting_types"] = FITTING_TYPES
        return safe

    @app.get("/")
    def index():
        settings = load_settings()
        return render_template(
            "index.html",
            app_name=APP_NAME,
            app_version=APP_VERSION,
            fitting_types=FITTING_TYPES,
            material_types=settings.get("material_types", MATERIAL_TYPE_PRESET),
        )

    @app.get("/health")
    def healthcheck():
        return jsonify(
            {
                "ok": True,
                "app": APP_NAME,
                "version": APP_VERSION,
                "legend_cache_present": LEGEND_CACHE_PATH.exists(),
                "exports_dir": str(EXPORTS_DIR),
            }
        )

    @app.get("/api/settings")
    def api_get_settings():
        return jsonify(safe_settings_payload())

    @app.post("/api/settings")
    def api_save_settings():
        data = json_body()
        s = load_settings()

        if "material_types" in data:
            mats = [str(m).strip() for m in data["material_types"] if str(m).strip()]
            if mats:
                s["material_types"] = dedupe_case_insensitive_keep_first(mats)

        if "company_side_multipliers" in data:
            cleaned = _clean_multiplier_table(data["company_side_multipliers"])
            if cleaned:
                s["company_side_multipliers"] = cleaned

        if "project_side_multipliers" in data:
            psm = {}
            for proj, tbl in data["project_side_multipliers"].items():
                proj_name = str(proj).strip()
                if not proj_name:
                    continue
                cleaned = _clean_multiplier_table(tbl)
                psm[proj_name] = cleaned
            s["project_side_multipliers"] = psm

        if "exclude_fitting_types" in data:
            ex = {str(k): bool(v) for k, v in data["exclude_fitting_types"].items()}
            s["exclude_fitting_types"] = ex

        if "multiplier_mode" in data:
            mode = str(data["multiplier_mode"])
            if mode in ("Company", "Project"):
                s["multiplier_mode"] = mode

        s = ensure_company_defaults(s)
        save_settings(s)
        return jsonify({"ok": True})

    @app.post("/api/admin/verify")
    def api_admin_verify():
        data = json_body()
        pw = str(data.get("password", "")).strip()
        s = load_settings()
        if password_matches(s, pw):
            return jsonify({"ok": True})
        return jsonify({"ok": False, "error": "Incorrect password"}), 401

    @app.post("/api/admin/change-password")
    def api_admin_change_password():
        data = json_body()
        current = str(data.get("current", "")).strip()
        new_pw = str(data.get("new_password", "")).strip()
        s = load_settings()
        if not password_matches(s, current):
            return jsonify({"ok": False, "error": "Current password incorrect"}), 401
        if len(new_pw) < 12:
            return jsonify({"ok": False, "error": "Password must be at least 12 characters"}), 400
        s = set_admin_password(s, new_pw)
        save_settings(s)
        return jsonify({"ok": True})

    @app.post("/api/admin/reset-settings")
    def api_admin_reset():
        data = json_body()
        pw = str(data.get("password", "")).strip()
        s = load_settings()
        if not password_matches(s, pw):
            return jsonify({"ok": False, "error": "Incorrect password"}), 401
        from core.logic import DEFAULT_SETTINGS

        new_s = json.loads(json.dumps(DEFAULT_SETTINGS))
        new_s["admin_password_hash"] = s.get("admin_password_hash", "")
        new_s = ensure_company_defaults(new_s)
        save_settings(new_s)
        return jsonify({"ok": True})

    @app.post("/api/admin/upload-legend")
    def api_admin_upload_legend():
        data = json_body()
        required_keys = {"concat", "concat_nospace", "desc", "desc_nospace"}
        if not required_keys.issubset(data.keys()):
            return jsonify({"ok": False, "error": "Legend cache JSON is missing required keys."}), 400
        LEGEND_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
        LEGEND_CACHE_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
        return jsonify({"ok": True, "message": "Legend cache uploaded."})

    @app.post("/api/run")
    def api_run():
        files = request.files.getlist("pdfs")
        if not files or all(f.filename == "" for f in files):
            return jsonify({"ok": False, "error": "No PDF files uploaded"}), 400

        mode = request.form.get("mode", "Company")
        project = request.form.get("project", "").strip()
        export_filename = sanitize_export_filename(request.form.get("export_filename", "BOM_Export"))
        skip_unclassified = request.form.get("skip_unclassified", "false").lower() == "true"

        if mode == "Project" and not project:
            return jsonify({"ok": False, "error": "Project name is required when Project mode is selected."}), 400

        saved_paths: list[Path] = []
        for f in files:
            if not f.filename or not f.filename.lower().endswith(".pdf"):
                continue
            safe_name = f"{uuid4().hex}_{Path(f.filename).name}"
            dest = upload_dir / safe_name
            f.save(str(dest))
            saved_paths.append(dest)

        if not saved_paths:
            return jsonify({"ok": False, "error": "No valid PDF files received"}), 400

        settings = load_settings()
        if skip_unclassified:
            settings.setdefault("exclude_fitting_types", {})["Unclassified"] = True

        try:
            result = run_bom(
                pdf_paths=[str(p) for p in saved_paths],
                settings=settings,
                export_filename=export_filename,
                mode=mode,
                project=project if mode == "Project" else "",
            )
        except Exception as exc:
            app.logger.exception("BOM run failed")
            return jsonify({"ok": False, "error": f"{type(exc).__name__}: {exc}"}), 500
        finally:
            for path in saved_paths:
                try:
                    path.unlink(missing_ok=True)
                except OSError:
                    app.logger.warning("Failed to remove temporary upload %s", path)

        try:
            run_id = save_run(result, [str(p) for p in saved_paths], mode, project)
            result["run_id"] = run_id
        except Exception:
            app.logger.exception("Failed to persist run history")
            result["run_id"] = None

        return jsonify(result)

    @app.get("/api/download/<path:filename>")
    def api_download(filename: str):
        safe = Path(filename).name
        path = EXPORTS_DIR / safe
        if not path.exists():
            abort(404)
        return send_file(str(path), as_attachment=True, download_name=safe)

    @app.get("/api/history")
    def api_history():
        return jsonify(get_history(50))

    @app.get("/api/history/<int:run_id>")
    def api_history_detail(run_id: int):
        run = get_run(run_id)
        if not run:
            abort(404)
        return jsonify(run)

    @app.delete("/api/history/<int:run_id>")
    def api_history_delete(run_id: int):
        delete_run(run_id)
        return jsonify({"ok": True})

    @app.errorhandler(HTTPException)
    def handle_http_error(exc: HTTPException):
        if request.path.startswith("/api/") or request.path == "/health":
            return jsonify({"ok": False, "error": exc.description}), exc.code
        return exc

    @app.errorhandler(Exception)
    def handle_unexpected_error(exc: Exception):
        app.logger.exception("Unexpected application error")
        if request.path.startswith("/api/") or request.path == "/health":
            return jsonify({"ok": False, "error": "Internal server error"}), 500
        return render_template("index.html", app_name=APP_NAME, app_version=APP_VERSION, fitting_types=FITTING_TYPES, material_types=load_settings().get("material_types", MATERIAL_TYPE_PRESET)), 500

    app.logger.info(
        "Starting %s v%s with default admin password configured=%s",
        APP_NAME,
        APP_VERSION,
        DEFAULT_ADMIN_PASSWORD == "FBT2026!",
    )
    return app


app = create_app()


if __name__ == "__main__":
    print(f"\n  {'=' * 44}")
    print(f"  {APP_NAME}  v{APP_VERSION}")
    print(f"  {'=' * 44}")
    print(f"  Open in browser: http://localhost:{DEFAULT_PORT}")
    print("  Production server: gunicorn -w 4 -b 0.0.0.0:5000 app:app")
    print("  Press Ctrl+C to stop\n")
    app.run(debug=_bool_env("FLASK_DEBUG", False), host=DEFAULT_HOST, port=DEFAULT_PORT)
