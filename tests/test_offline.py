"""Tests du mode hors ligne (offline_config.py, scripts/vendor_assets.py,
agents/llm_tiers.py).

    python -m unittest discover tests
"""
import base64
import hashlib
import http.server
import importlib
import os
import sys
import tempfile
import threading
import types
import unittest
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "scripts"))


def load_offline_config(offline: bool, **env):
    values = {"CHECKERS_OFFLINE": "true" if offline else "", "MAP_TILE_URL": "", **env}
    with mock.patch.dict(os.environ, values):
        import offline_config
        return importlib.reload(offline_config)


class AssetResolutionTest(unittest.TestCase):
    def setUp(self):
        self.vendor = tempfile.mkdtemp()

    def resolve(self, offline):
        cfg = load_offline_config(offline)
        cfg.VENDOR_DIR = self.vendor
        return cfg, cfg.resolve_assets(lambda path: "/static/" + path)

    def test_online_without_local_copy_uses_cdn(self):
        _, assets = self.resolve(offline=False)
        self.assertTrue(assets["chartjs"]["url"].startswith("https://cdn.jsdelivr.net/"))
        self.assertFalse(assets["chartjs"]["local"])

    def test_offline_without_local_copy_disables_instead_of_cdn(self):
        cfg, assets = self.resolve(offline=True)
        self.assertTrue(all(asset is None for asset in assets.values()))
        self.assertFalse(cfg.viewer_config(assets)["enabled"])

    def test_local_copy_is_preferred_and_keeps_integrity(self):
        cfg = load_offline_config(True)
        cfg.VENDOR_DIR = self.vendor
        path = os.path.join(self.vendor, cfg.ASSETS["chartjs"]["local"])
        os.makedirs(os.path.dirname(path))
        open(path, "w").close()
        assets = cfg.resolve_assets(lambda p: "/static/" + p)
        self.assertEqual(assets["chartjs"]["url"], "/static/vendor/chart.js/4.4.1/chart.umd.js")
        self.assertEqual(assets["chartjs"]["integrity"], cfg.ASSETS["chartjs"]["integrity"])

    def test_viewer_enabled_offline_only_with_all_parts(self):
        cfg = load_offline_config(True)
        cfg.VENDOR_DIR = self.vendor
        for name in ("viewer_module", "viewer_wasm", "viewer_worker"):
            path = os.path.join(self.vendor, cfg.ASSETS[name]["local"])
            os.makedirs(os.path.dirname(path), exist_ok=True)
            open(path, "w").close()
        viewer = cfg.viewer_config(cfg.resolve_assets(lambda p: "/static/" + p))
        self.assertEqual(viewer, {"enabled": True, "worker_url": "/static/vendor/thatopen/fragments-worker.mjs",
                                  "wasm_path": "/static/vendor/web-ifc/0.0.77/"})


class MapAndCspTest(unittest.TestCase):
    def test_no_external_tiles_offline(self):
        cfg = load_offline_config(True)
        self.assertIsNone(cfg.map_config()["tile_url"])
        self.assertNotIn("http", cfg.content_security_policy(None))

    def test_internal_tile_server_is_allowed_by_csp(self):
        cfg = load_offline_config(True)
        with mock.patch.dict(os.environ, {"MAP_TILE_URL": "https://{s}.tuiles.intranet.example/{z}/{x}/{y}.png"}):
            tile_url = cfg.map_config()["tile_url"]
        self.assertIn("img-src 'self' data: blob: https://*.tuiles.intranet.example",
                      cfg.content_security_policy(tile_url))

    def test_online_keeps_openstreetmap(self):
        cfg = load_offline_config(False)
        self.assertEqual(cfg.map_config()["tile_url"], cfg.OSM_TILE_URL)


class ProcessEnvironmentTest(unittest.TestCase):
    def test_ai_telemetry_disabled_offline(self):
        with mock.patch.dict(os.environ, {"CHECKERS_OFFLINE": "1"}, clear=False):
            for key in ("CREWAI_DISABLE_TELEMETRY", "OTEL_SDK_DISABLED", "CREWAI_TRACING_ENABLED"):
                os.environ.pop(key, None)
            import offline_config
            importlib.reload(offline_config).apply_process_environment()
            self.assertEqual(os.environ["CREWAI_DISABLE_TELEMETRY"], "true")
            self.assertEqual(os.environ["OTEL_SDK_DISABLED"], "true")
            self.assertEqual(os.environ["CREWAI_TRACING_ENABLED"], "false")


class CloudEscalationTest(unittest.TestCase):
    """L'escalade cloud doit être refusée en mode hors ligne même avec une
    clé configurée. CrewAI est remplacé par un module factice (il n'est pas
    requis pour exécuter les tests)."""

    def build_cloud_llm(self, offline):
        fake_crewai = types.ModuleType("crewai")
        fake_crewai.LLM = type("LLM", (), {"__init__": lambda self, **kw: None})
        fake_crewai.Crew = fake_crewai.Process = object
        with mock.patch.dict(sys.modules, {"crewai": fake_crewai}), \
             mock.patch.dict(os.environ, {"CHECKERS_OFFLINE": "true" if offline else "",
                                          "CLOUD_ANTHROPIC_API_KEY": "sk-test"}):
            import offline_config
            importlib.reload(offline_config)
            sys.modules.pop("agents.llm_tiers", None)
            llm_tiers = importlib.import_module("agents.llm_tiers")
            with mock.patch.object(llm_tiers.budget, "has_budget", return_value=True):
                return llm_tiers.build_cloud_llm()

    def test_cloud_refused_offline(self):
        self.assertIsNone(self.build_cloud_llm(offline=True))

    def test_cloud_still_possible_online(self):
        self.assertIsNotNone(self.build_cloud_llm(offline=False))


class VendorScriptTest(unittest.TestCase):
    """scripts/vendor_assets.py contre un faux CDN local : un fichier dont
    l'empreinte ne correspond pas n'est jamais écrit."""

    def setUp(self):
        self.served = tempfile.mkdtemp()
        handler = lambda *a, **kw: http.server.SimpleHTTPRequestHandler(*a, directory=self.served, **kw)
        self.server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        self.base = f"http://127.0.0.1:{self.server.server_address[1]}/"

    def tearDown(self):
        self.server.shutdown()

    def run_fetch(self, content: bytes, declared: bytes):
        import vendor_assets
        cfg = vendor_assets.offline_config
        with open(os.path.join(self.served, "chart.js"), "wb") as f:
            f.write(content)
        integrity = "sha384-" + base64.b64encode(hashlib.sha384(declared).digest()).decode()
        vendor = tempfile.mkdtemp()
        spec = dict(cfg.ASSETS["chartjs"], cdn=self.base + "chart.js", integrity=integrity)
        with mock.patch.dict(cfg.ASSETS, {"chartjs": spec}), mock.patch.object(cfg, "VENDOR_DIR", vendor):
            vendor_assets.fetch_verified("chartjs")
        return os.path.join(vendor, spec["local"])

    def test_matching_integrity_is_written(self):
        path = self.run_fetch(b"console.log(1)", declared=b"console.log(1)")
        self.assertTrue(os.path.isfile(path))

    def test_tampered_file_is_rejected(self):
        with self.assertRaises(SystemExit):
            self.run_fetch(b"alert('altered')", declared=b"console.log(1)")


if __name__ == "__main__":
    unittest.main()
