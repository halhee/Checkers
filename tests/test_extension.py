"""Tests de l'extension Blender : garantie hors ligne, enregistrement et
opérateurs du panneau, avec bpy et Bonsai simulés.

    python -m unittest discover tests
"""
import ast
import contextlib
import json
import os
import sys
import tempfile
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import support  # noqa: E402,F401
from support import PACKAGE, RULES_XLSX, add_equipment, build_model, guid  # noqa: E402

from checkers_bonsai.core import recette  # noqa: E402


class OfflineGuaranteeTest(unittest.TestCase):
    """Aucun module de l'extension n'importe de bibliothèque réseau ni ne
    lance de programme externe, et le manifeste ne demande pas la permission
    « network »."""
    FORBIDDEN_MODULES = {"socket", "ssl", "urllib", "http", "requests", "httpx", "aiohttp", "ftplib", "smtplib",
                         "webbrowser", "xmlrpc", "telnetlib", "poplib", "imaplib", "subprocess", "multiprocessing"}

    def test_no_network_or_process_imports(self):
        checked = 0
        for root, _, files in os.walk(PACKAGE):
            for name in files:
                if name.endswith(".py"):
                    checked += 1
                    with open(os.path.join(root, name), encoding="utf-8") as fh:
                        tree = ast.parse(fh.read())
                    for node in ast.walk(tree):
                        modules = [a.name for a in node.names] if isinstance(node, ast.Import) else \
                                  [node.module or ""] if isinstance(node, ast.ImportFrom) else []
                        for module in modules:
                            self.assertNotIn(module.split(".")[0], self.FORBIDDEN_MODULES, f"{name} : {module}")
        self.assertGreater(checked, 10)

    def test_manifest(self):
        import tomllib
        with open(os.path.join(PACKAGE, "blender_manifest.toml"), "rb") as fh:
            manifest = tomllib.load(fh)
        self.assertEqual(set(manifest["permissions"]), {"files"})         # surtout pas « network »
        self.assertNotIn("wheels", manifest)                               # aucune dépendance embarquée
        # Règles de Blender, faute de quoi l'extension peut être refusée à
        # l'installation : 64 caractères au plus, sans ponctuation finale.
        for text in [manifest["tagline"], *manifest["permissions"].values()]:
            self.assertLessEqual(len(text), 64, text)
            self.assertNotIn(text[-1], ".!?,;:", text)
        self.assertRegex(manifest["version"], r"^\d+\.\d+\.\d+$")
        self.assertEqual(manifest["type"], "add-on")
        for required in ("id", "name", "maintainer", "license", "blender_version_min", "schema_version"):
            self.assertTrue(manifest[required], required)


class _Collection(list):
    def add(self):
        self.append(types.SimpleNamespace())
        return self[-1]

    def clear(self):
        del self[:]


class _Object:
    def __init__(self, name):
        self.name, self.color, self.selected = name, (1.0, 1.0, 1.0, 1.0), False

    def select_set(self, value):
        self.selected = value

    def visible_get(self):
        return True


class ExtensionTest(unittest.TestCase):
    def setUp(self):
        import bpy
        import bonsai.tool as tool
        import checkers_bonsai
        from checkers_bonsai import operators
        from checkers_bonsai.props import CheckersProperties
        self.bpy, self.tool, self.ext, self.ops = bpy, tool, checkers_bonsai, operators
        checkers_bonsai.register()
        self.dir = tempfile.mkdtemp()
        self.open_model(add_equipment(build_model()))

        props = types.SimpleNamespace(
            ids_path="", rules_path=RULES_XLSX, corps_etat="ARC", check_structure=True, previous_path="",
            lot="Lot 02", phase="DOE", tester="J. Martin", marking="confidentiel", threshold_ok=0.98,
            threshold_reserve=0.9, bcf_author="", anomalies=_Collection(), anomaly_index=0)
        props.marking_text = types.MethodType(CheckersProperties.marking_text, props)
        operators.clear_results(props)
        self.props = props
        self.space = types.SimpleNamespace(type="VIEW_3D", shading=types.SimpleNamespace(color_type="MATERIAL"),
                                           as_pointer=lambda: 1)
        area = types.SimpleNamespace(type="VIEW_3D", spaces=[self.space],
                                     regions=[types.SimpleNamespace(type="WINDOW")])
        self.scene = types.SimpleNamespace(checkers=props, BIMProperties=types.SimpleNamespace(is_dirty=False))
        bpy.data.scenes = [self.scene]
        self.context = types.SimpleNamespace(
            scene=self.scene, area=area, selected_objects=[],
            view_layer=types.SimpleNamespace(objects=types.SimpleNamespace(active=None)),
            window_manager=types.SimpleNamespace(
                progress_begin=lambda *a: None, progress_end=lambda: None,
                windows=[types.SimpleNamespace(screen=types.SimpleNamespace(areas=[area]))]),
            temp_override=lambda **k: contextlib.nullcontext())

    def tearDown(self):
        self.ext.unregister()
        self.assertEqual(self.bpy.registered, [])
        self.assertEqual(self.bpy.app.handlers.load_post, [])

    def open_model(self, model):
        self.model = model
        self.tool.Ifc.file, self.tool.Ifc.path = model, ""
        self.tool.Ifc.objects = {e.GlobalId: _Object(e.GlobalId) for e in model.entities if e.is_a("IfcElement")}
        self.bpy.data.objects = {o.name: o for o in self.tool.Ifc.objects.values()}

    def run_recette(self):
        self.assertEqual(self.ops.CHECKERS_OT_run_recette().execute(self.context), {"FINISHED"})

    def export(self, operator_class, extension):
        op = operator_class()
        op.invoke(self.context, None)
        default_name = op.filepath
        op.filepath = os.path.join(self.dir, "sortie" + extension)
        self.assertEqual(op.execute(self.context), {"FINISHED"}, getattr(op, "reports", None))
        self.assertGreater(os.path.getsize(op.filepath), 0)
        return default_name, op

    def test_workflow(self):
        ops, ctx, props = self.ops, self.context, self.props
        self.assertTrue(ops.CHECKERS_OT_run_recette.poll(ctx))
        self.run_recette()
        self.assertEqual(props.verdict, recette.REJECTED)
        self.assertIn("Structure : aucun constat", props.summary)
        self.assertIn("Règles Excel : 25,0 % conformes (4 éléments)", props.summary)
        self.assertGreater(len(props.anomalies), 0)
        # Aucun identifiant d'élément n'est copié dans les propriétés de la
        # scène, donc dans le fichier .blend.
        self.assertTrue(all(not hasattr(item, "guids") for item in props.anomalies))
        self.assertNotIn(guid(3), json.dumps({k: v for k, v in vars(props).items() if isinstance(v, (str, int))}))

        select = ops.CHECKERS_OT_select_anomaly()
        select.index = 0
        self.assertEqual(select.execute(ctx), {"FINISHED"})
        self.assertTrue(any(o.selected for o in self.tool.Ifc.objects.values()))

        ops.CHECKERS_OT_color_results().execute(ctx)
        self.assertEqual(self.space.shading.color_type, "OBJECT")
        self.assertEqual(self.tool.Ifc.objects[guid(1)].color, ops.COLOR_VALID)
        self.assertEqual(self.tool.Ifc.objects[guid(3)].color, ops.COLOR_INVALID)
        self.assertEqual(self.tool.Ifc.objects[guid(5)].color, (1.0, 1.0, 1.0, 1.0))   # dalle non contrôlée (ARC)
        ops.CHECKERS_OT_reset_colors().execute(ctx)
        self.assertEqual(self.space.shading.color_type, "MATERIAL")
        self.assertEqual(self.tool.Ifc.objects[guid(3)].color, (1.0, 1.0, 1.0, 1.0))

        name, _ = self.export(ops.CHECKERS_OT_export_report, ".xlsx")
        self.assertTrue(name.startswith("Recette_Lot_02_") and name.endswith(".xlsx"), name)
        self.assertTrue(os.path.isfile(os.path.join(self.dir, "sortie.json")))          # journal, même nom
        name, _ = self.export(ops.CHECKERS_OT_export_bcf, ".bcf")
        self.assertTrue(name.endswith(".bcf"), name)
        name, op = self.export(ops.CHECKERS_OT_export_inventory, ".xlsx")
        self.assertTrue(name.startswith("Inventaire_Lot_02_"), name)
        self.assertIn("3 équipement(s), 1 local(aux)", op.reports[-1][1])

    def test_follow_up_between_deliveries(self):
        self.run_recette()
        self.export(self.ops.CHECKERS_OT_export_report, ".xlsx")
        self.props.previous_path = os.path.join(self.dir, "sortie.json")
        self.model.by_guid(guid(3))._psets = {"Pset_WallCommon": {"AcousticRating": "Rw 50"}}
        self.run_recette()
        self.assertIn("1 levée(s), 2 persistante(s), 0 nouvelle(s)", self.props.evolution)

    def test_structure_only_and_model_level_finding(self):
        self.props.rules_path = ""
        self.model.length_unit = None
        self.assertTrue(self.ops.CHECKERS_OT_run_recette.poll(self.context))
        self.run_recette()
        self.assertEqual(self.props.verdict, recette.RESERVED)
        self.assertEqual(self.props.summary, "Structure : 1 majeur(s)")
        select = self.ops.CHECKERS_OT_select_anomaly()
        select.index = 0
        self.assertEqual(select.execute(self.context), {"FINISHED"})
        self.assertIn("modèle entier", select.reports[-1][1])
        self.props.check_structure = False
        self.assertFalse(self.ops.CHECKERS_OT_run_recette.poll(self.context))           # plus aucun contrôle

    def test_unsaved_changes_warning(self):
        self.scene.BIMProperties.is_dirty = True
        self.run_recette()
        self.assertIn("non enregistrées", self.props.warnings)

    def test_results_do_not_outlive_the_model(self):
        ops, ctx = self.ops, self.context
        self.run_recette()
        self.assertFalse(ops.is_stale())
        self.open_model(build_model())                       # un autre projet est ouvert dans Bonsai
        self.assertTrue(ops.is_stale())
        self.assertFalse(ops.CHECKERS_OT_color_results.poll(ctx))
        self.assertFalse(ops.CHECKERS_OT_export_report.poll(ctx))
        self.assertFalse(ops.CHECKERS_OT_export_bcf.poll(ctx))
        select = ops.CHECKERS_OT_select_anomaly()
        select.index = 0
        self.assertEqual(select.execute(ctx), {"CANCELLED"})
        # À l'ouverture d'un fichier Blender, le panneau repart sans résultat.
        self.run_recette()
        self.assertTrue(self.props.has_result)
        self.bpy.app.handlers.load_post[0](None)
        self.assertFalse(self.props.has_result)
        self.assertEqual((self.props.verdict, len(self.props.anomalies), ops.last_result()), ("", 0, None))

    def test_errors_are_reported_not_raised(self):
        self.props.rules_path = os.path.join(self.dir, "absent.xlsx")
        op = self.ops.CHECKERS_OT_run_recette()
        self.assertEqual(op.execute(self.context), {"CANCELLED"})
        self.assertIn("Recette impossible", op.reports[-1][1])
        self.assertFalse(self.props.has_result)

    def test_example_operators(self):
        op = self.ops.CHECKERS_OT_use_example_ids()
        op.example = "exploitation_maintenance.ids"
        op.execute(self.context)
        self.assertTrue(os.path.isfile(self.props.ids_path))
        self.props.rules_path = ""
        self.ops.CHECKERS_OT_use_example_rules().execute(self.context)
        self.assertTrue(os.path.isfile(self.props.rules_path))

    def test_no_model(self):
        self.tool.Ifc.file = None
        self.assertFalse(self.ops.CHECKERS_OT_run_recette.poll(self.context))
        self.assertFalse(self.ops.CHECKERS_OT_export_inventory.poll(self.context))


if __name__ == "__main__":
    unittest.main()
