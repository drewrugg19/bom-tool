import importlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

app_mod = importlib.import_module("FabBOMTool.app")
history_mod = importlib.import_module("FabBOMTool.core.history")
logic_mod = importlib.import_module("FabBOMTool.core.logic")


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


class LogicParserRegressionTestCase(unittest.TestCase):
    def test_normalize_row_matches_legacy_shape(self):
        reject_log = []
        row = ["2", "GROOVED", "FIG 777 COUPLING", "4", '10" PVC', "B12345"]

        normalized = logic_mod.normalize_row(row, "demo.pdf", "B12345", reject_log)

        self.assertEqual(normalized["Batch"], "B12345")
        self.assertEqual(normalized["Size"], "2")
        self.assertEqual(normalized["Install Type"], "GROOVED")
        self.assertEqual(normalized["Description"], "FIG 777 COUPLING")
        self.assertEqual(normalized["Count"], 4)
        self.assertEqual(normalized["Length"], '10"')
        self.assertEqual(normalized["Material"], "PVC")
        self.assertEqual(reject_log, [])

    def test_normalize_row_rejects_metadata_rows_without_length(self):
        reject_log = []
        row = ["FABRICATION ISOMETRIC", "B12345", "REV 2"]

        normalized = logic_mod.normalize_row(row, "demo.pdf", "B12345", reject_log)

        self.assertIsNone(normalized)
        self.assertEqual(reject_log[0]["Reason"], "Missing Length pattern")

    def test_classify_fitting_type_with_legend_falls_back_to_keyword_logic(self):
        legend_maps = ({}, {}, {}, {}, set())

        self.assertEqual(
            logic_mod.classify_fitting_type_with_legend("Victaulic 90 Ell", legend_maps),
            "Elbow",
        )

    def test_classify_fitting_type_with_legend_prioritizes_short_sweep_override(self):
        legend_maps = ({"SHORT SWEEP COMBO": "Coupling"}, {}, {}, {}, set())

        self.assertEqual(
            logic_mod.classify_fitting_type_with_legend("6 in Short Sweep 90", legend_maps),
            "Elbow",
        )

    def test_classify_fitting_type_with_legend_prioritizes_combination_override(self):
        legend_maps = ({"COMBINATION WYE": "Wye"}, {}, {}, {}, set())

        self.assertEqual(
            logic_mod.classify_fitting_type_with_legend("4 in Combination Wye", legend_maps),
            "Tee",
        )

    def test_classify_fitting_type_with_legend_prioritizes_sanitary_cross_override(self):
        legend_maps = ({"SANITARY CROSS": "Cross"}, {}, {}, {}, set())

        self.assertEqual(
            logic_mod.classify_fitting_type_with_legend("4 in Sanitary Cross", legend_maps),
            "Tee",
        )

    def test_classify_fitting_type_with_legend_prioritizes_p_trap_override(self):
        legend_maps = ({}, {}, {}, {}, set())

        self.assertEqual(
            logic_mod.classify_fitting_type_with_legend("2 in P-Trap", legend_maps),
            "Tee",
        )

    def test_classify_fitting_type_with_legend_does_not_treat_sixteenth_bend_as_tee(self):
        legend_maps = ({}, {}, {}, {}, set())

        self.assertEqual(
            logic_mod.classify_fitting_type_with_legend("3 in Sixteenth Bend", legend_maps),
            "Elbow",
        )

    def test_classify_fitting_type_with_legend_maps_specific_ferrule_to_cap(self):
        legend_maps = ({}, {}, {}, {}, set())

        self.assertEqual(
            logic_mod.classify_fitting_type_with_legend(
                "Charlotte NoNH 52 S Tapped Ferrule with Southern Raised-Head Brass Plug ",
                legend_maps,
            ),
            "Cap",
        )

    def test_material_type_from_material_uses_legacy_aliases(self):
        materials = ["Copper", "PVC", "Nickel iron"]

        self.assertEqual(logic_mod.material_type_from_material("CU TUBE", materials), "Copper")
        self.assertEqual(logic_mod.material_type_from_material("nickeliron body", materials), "Nickel iron")

    def test_size_to_diameter_in_uses_largest_dimension(self):
        self.assertEqual(logic_mod.size_to_diameter_in('2 x 1-1/2'), 2.0)



if __name__ == "__main__":
    unittest.main()
