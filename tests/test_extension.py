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
    """CollectionProperty simulée ; `defaults` : valeurs par défaut des
    propriétés d'un nouvel élément, comme dans Blender."""

    def __init__(self, **defaults):
        super().__init__()
        self.defaults = defaults

    def add(self):
        self.append(types.SimpleNamespace(**self.defaults))
        return self[-1]

    def remove(self, index):
        del self[index]

    def clear(self):
        del self[:]


class _Object:
    def __init__(self, name):
        self.name, self.color, self.selected = name, (1.0, 1.0, 1.0, 1.0), False
        self.in_view_layer = True

    def select_set(self, value):
        if not self.in_view_layer:
            raise RuntimeError("Object can't be selected because it is not in View Layer")
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
            survey_path="", journals_dir="", diagnostic="",
            lot="Lot 02", phase="DOE", tester="J. Martin", marking="confidentiel", threshold_ok=0.98,
            threshold_reserve=0.9, bcf_author="", anomalies=_Collection(), anomaly_index=0,
            lots=_Collection(name="", path=""), lot_index=0, check_federation=True, check_clashes=False,
            clash_rules=_Collection(other="Ce modèle", selector_a="", selector_b="", mode="intersection",
                                    distance=0.01, severity="Majeur"),
            clash_rule_index=0)
        props.marking_text = types.MethodType(CheckersProperties.marking_text, props)
        operators.clear_results(props)
        self.props = props
        self.space = types.SimpleNamespace(
            type="VIEW_3D", shading=types.SimpleNamespace(type="MATERIAL", color_type="TEXTURE"),
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
        # La couleur d'objet ne se voit qu'en mode solide : il est activé.
        self.assertEqual((self.space.shading.type, self.space.shading.color_type), ("SOLID", "OBJECT"))
        self.assertEqual(self.tool.Ifc.objects[guid(1)].color, ops.COLOR_VALID)
        self.assertEqual(self.tool.Ifc.objects[guid(3)].color, ops.COLOR_INVALID)
        self.assertEqual(self.tool.Ifc.objects[guid(5)].color, (1.0, 1.0, 1.0, 1.0))   # dalle non contrôlée (ARC)
        ops.CHECKERS_OT_reset_colors().execute(ctx)
        self.assertEqual((self.space.shading.type, self.space.shading.color_type), ("MATERIAL", "TEXTURE"))
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


    def test_object_outside_view_layer_does_not_break_selection(self):
        self.run_recette()
        result = self.ops.last_result()
        index = next(i for i, a in enumerate(result.anomalies) if len(a["guids"]) >= 1)
        self.tool.Ifc.objects[result.anomalies[index]["guids"][0]].in_view_layer = False
        select = self.ops.CHECKERS_OT_select_anomaly()
        select.index = index
        self.assertEqual(select.execute(self.context), {"FINISHED"})

    def test_field_survey_round_trip(self):
        """Fiche de relevé produite depuis le panneau, remplie, puis contrôlée."""
        from openpyxl import load_workbook
        _, op = self.export(self.ops.CHECKERS_OT_export_survey, ".xlsx")
        self.assertIn("3 équipement(s) à vérifier", op.reports[-1][1])
        workbook = load_workbook(op.filepath)
        ws = workbook["Relevé terrain"]
        header = [c.value for c in ws[1]]
        column = {name: header.index(name) + 1 for name in header}
        for row in range(2, ws.max_row + 1):
            is_pump = ws.cell(row, column["Repère (maquette)"]).value == "PMP-001"
            ws.cell(row, column["Présent sur site (oui / non)"], "oui" if is_pump else "non")
            if is_pump:
                ws.cell(row, column["N° de série relevé"], "SN-99")          # la maquette dit SN-42
        workbook.save(op.filepath)

        self.props.survey_path, self.props.rules_path, self.props.check_structure = op.filepath, "", False
        self.run_recette()
        self.assertEqual(self.props.verdict, recette.REJECTED)
        self.assertIn("Terrain : 0,0 % sans écart (3 vérifications)", self.props.summary)
        requirements = sorted(item.requirement for item in self.props.anomalies)
        self.assertEqual(requirements, ["Terrain : absent sur site", "Terrain : absent sur site",
                                        "Terrain : n° de série différent"])

    def test_tracking_and_diagnostic(self):
        ops, ctx = self.ops, self.context
        self.assertFalse(ops.CHECKERS_OT_export_tracking.poll(ctx))          # dossier non renseigné
        self.run_recette()
        self.export(ops.CHECKERS_OT_export_report, ".xlsx")                  # écrit aussi le journal
        self.props.journals_dir = self.dir
        self.assertTrue(ops.CHECKERS_OT_export_tracking.poll(ctx))
        os.rename(os.path.join(self.dir, "sortie.xlsx"), os.path.join(self.dir, "fiche.xlsx"))
        name, op = self.export(ops.CHECKERS_OT_export_tracking, ".xlsx")
        self.assertTrue(name.startswith("Suivi_recettes_"), name)
        self.assertIn("1 recette(s)", op.reports[-1][1])

        diagnostic = ops.CHECKERS_OT_diagnostic()
        self.assertEqual(diagnostic.execute(ctx), {"FINISHED"})
        lines = self.props.diagnostic.split("\n")
        self.assertTrue(lines[0].startswith("✓ Bonsai"), lines)
        self.assertTrue(any(line.startswith("✓ Modèle IFC : schéma IFC4") for line in lines), lines)
        self.assertTrue(all(line.startswith("✓") for line in lines), lines)


    def test_panels_draw_with_valid_properties_and_operators(self):
        """Dessine chaque section du panneau avec une mise en page simulée,
        qui refuse une propriété ou un opérateur inexistant et une icône qui
        n'est pas une chaîne."""
        from checkers_bonsai import props as props_module, ui
        test = self
        known_props = set().union(*(cls.__annotations__ for cls in props_module.classes))
        known_operators = {cls.bl_idname for cls in self.ops.classes}

        class Layout:
            drawn = []

            def _same(self, *args, **kwargs):
                return self
            row = column = box = separator = _same

            def label(self, text="", icon="NONE"):
                test.assertIsInstance(text, str)
                test.assertIsInstance(icon, str)
                self.drawn.append(text)

            def prop(self, data, name, **kwargs):
                test.assertTrue(hasattr(data, name), name)
                test.assertIn(name, known_props)

            def operator(self, idname, **kwargs):
                test.assertIn(idname, known_operators)
                return types.SimpleNamespace()

            def operator_menu_enum(self, idname, prop, **kwargs):
                test.assertIn(idname, known_operators)

            def template_list(self, list_type, list_id, data, name, active_data, active_name, **kwargs):
                test.assertIn(list_type, {cls.__name__ for cls in ui.classes})
                test.assertIn(name, known_props)
                test.assertIn(active_name, known_props)

        def draw_all():
            for cls in ui.classes:
                if hasattr(cls, "draw_item"):
                    continue                                         # listes : dessinées plus bas
                if getattr(cls, "poll", None) is None or cls.poll(self.context):
                    panel = cls()
                    panel.layout = Layout()
                    panel.draw(self.context)

        self.props.rules_path = RULES_XLSX
        self.ops.CHECKERS_OT_add_lot().execute(self.context)
        self.ops.CHECKERS_OT_add_clash_rule().execute(self.context)  # coche aussi « conflits »
        self.assertTrue(self.props.check_clashes)
        draw_all()                                                   # avant toute recette
        self.run_recette()
        self.ops.CHECKERS_OT_diagnostic().execute(self.context)
        draw_all()                                                   # avec résultats et diagnostic
        self.assertTrue(any(text.startswith("Verdict : ") for text in Layout.drawn))
        row = ui.CHECKERS_UL_anomalies()
        row.draw_item(self.context, Layout(), None, self.props.anomalies[0], 0, None, "", 0)
        ui.CHECKERS_UL_lots().draw_item(self.context, Layout(), None, self.props.lots[0], 0, None, "", 0)
        ui.CHECKERS_UL_clash_rules().draw_item(self.context, Layout(), None, self.props.clash_rules[0], 0, None, "", 0)
        self.assertIn("Conflit interne (intersection, 1 cm)", Layout.drawn)
        self.tool.Ifc.file = None                                    # sans modèle : suivi et installation seulement
        Layout.drawn.clear()
        draw_all()
        self.assertIn("Ouvrez un projet IFC dans Bonsai.", Layout.drawn)


if __name__ == "__main__":
    unittest.main()
