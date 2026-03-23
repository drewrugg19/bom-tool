import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import sys
sys.path.insert(0, "/workspace/bom-tool/FabBOMTool")

import app as app_mod
import core.history as history_mod
import core.logic as logic_mod


class AppTestCase(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        root = Path(self.temp_dir.name)
        self.data_dir = root / "data"
        self.exports_dir = root / "exports"
        self.upload_dir = root / "uploads"
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.exports_dir.mkdir(parents=True, exist_ok=True)
        self.upload_dir.mkdir(parents=True, exist_ok=True)

        self.patchers = [
            mock.patch.object(logic_mod, "DATA_DIR", self.data_dir),
            mock.patch.object(logic_mod, "SETTINGS_FILE", self.data_dir / "settings.json"),
            mock.patch.object(logic_mod, "LEGEND_CACHE_PATH", self.data_dir / "Legend.cache.json"),
            mock.patch.object(logic_mod, "EXPORTS_DIR", self.exports_dir),
            mock.patch.object(history_mod, "DB_PATH", self.data_dir / "history.db"),
            mock.patch.object(app_mod, "DATA_DIR", self.data_dir),
            mock.patch.object(app_mod, "EXPORTS_DIR", self.exports_dir),
            mock.patch.object(app_mod, "LEGEND_CACHE_PATH", self.data_dir / "Legend.cache.json"),
        ]
        for patcher in self.patchers:
            patcher.start()

        self.app = app_mod.create_app()
        self.client = self.app.test_client()

    def tearDown(self):
        for patcher in reversed(self.patchers):
            patcher.stop()
        self.temp_dir.cleanup()

    def test_healthcheck_reports_ok(self):
        response = self.client.get("/health")
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["version"], logic_mod.APP_VERSION)

    def test_settings_response_hides_password_fields(self):
        settings = logic_mod.load_settings()
        logic_mod.set_admin_password(settings, "VerySecurePassword!")
        logic_mod.save_settings(settings)

        response = self.client.get("/api/settings")
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertNotIn("admin_password", payload)
        self.assertNotIn("admin_password_hash", payload)

    def test_verify_legacy_plaintext_password_migrates_to_hash(self):
        legacy_settings = dict(logic_mod.DEFAULT_SETTINGS)
        legacy_settings["admin_password"] = "FBT2026!"
        (self.data_dir / "settings.json").write_text(json.dumps(legacy_settings), encoding="utf-8")

        response = self.client.post("/api/admin/verify", json={"password": "FBT2026!"})
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.get_json()["ok"])

        migrated = logic_mod.load_settings()
        self.assertIn("admin_password_hash", migrated)
        self.assertNotIn("admin_password", migrated)
        self.assertTrue(logic_mod.password_matches(migrated, "FBT2026!"))

    def test_change_password_requires_strong_password(self):
        response = self.client.post(
            "/api/admin/change-password",
            json={"current": logic_mod.DEFAULT_ADMIN_PASSWORD, "new_password": "short"},
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("at least 12 characters", response.get_json()["error"])

    def test_upload_legend_validates_required_keys(self):
        response = self.client.post("/api/admin/upload-legend", json={"concat": {}})
        self.assertEqual(response.status_code, 400)

        payload = {
            "concat": {},
            "concat_nospace": {},
            "desc": {},
            "desc_nospace": {},
            "manufacturers": [],
        }
        response = self.client.post("/api/admin/upload-legend", json=payload)
        self.assertEqual(response.status_code, 200)
        saved = json.loads((self.data_dir / "Legend.cache.json").read_text(encoding="utf-8"))
        self.assertEqual(saved["manufacturers"], [])

    def test_run_requires_project_name_in_project_mode(self):
        response = self.client.post(
            "/api/run",
            data={
                "mode": "Project",
                "project": "",
                "export_filename": "demo",
                "pdfs": (io.BytesIO(b"%PDF-1.4"), "demo.pdf"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)
        self.assertIn("Project name is required", response.get_json()["error"])


if __name__ == "__main__":
    unittest.main()
