"""Tests de l'extension (checkers_bonsai).

    python -m unittest discover tests

Exécutés sans Blender ni IfcOpenShell, grâce aux doublures de tests/fakes/.
Validation BCF contre les schémas buildingSMART si BCF_XSD_DIR pointe vers
le dossier Schemas/ de https://github.com/buildingSMART/BCF-XML (branche
release_2_1) et que lxml est installé.
"""
import ast
import os
import sys
import tempfile
import types
import unittest
import zipfile
from unittest import mock
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FAKES = os.path.join(ROOT, "tests", "fakes")
PACKAGE = os.path.join(ROOT, "checkers_bonsai")
sys.path[:0] = [FAKES, ROOT]

import ifcopenshell  # noqa: E402  (doublure)
from openpyxl import load_workbook  # noqa: E402

from checkers_bonsai.core import bcf, ids as ids_check, recette, report, rules, spatial  # noqa: E402
from checkers_bonsai.core.psets import MISSING_PARAM, MISSING_PSET, check_psets  # noqa: E402

EXAMPLES = os.path.join(PACKAGE, "examples")
RULES_XLSX = os.path.join(EXAMPLES, "regles_exemple.xlsx")
XSD_DIR = os.environ.get("BCF_XSD_DIR")


def guid(n: int) -> str:
    return f"0{n:021d}"  # 22 caractères valides pour un IfcGuid


def build_model():
    """Bâtiment : 2 étages, un local ; quelques murs, portes, dalles."""
    f = ifcopenshell.file("IFC4")
    f.add("IfcProject", Name="Projet de démonstration – bâtiment A", GlobalId=guid(900))
    building = f.add("IfcBuilding", Name="Bâtiment A", GlobalId=guid(901))
    rdc = f.add("IfcBuildingStorey", Name="RDC", GlobalId=guid(902))
    r1 = f.add("IfcBuildingStorey", Name="R+1", GlobalId=guid(903))
    ifcopenshell.aggregate(rdc, building)
    ifcopenshell.aggregate(r1, building)
    local = f.add("IfcSpace", Name="012", GlobalId=guid(904))
    ifcopenshell.aggregate(local, rdc)

    ok = {"Pset_WallCommon": {"AcousticRating": "Rw 45", "LoadBearing": True}}
    walls = [
        f.add("IfcWall", GlobalId=guid(1), Name="Mur conforme", psets=ok),
        f.add("IfcWallStandardCase", GlobalId=guid(2), Name="=HYPERLINK(\"x\")",
              psets={"Pset_WallCommon": {"AcousticRating": "  "}}),
        f.add("IfcWall", GlobalId=guid(3), Name="Mur sans PSet"),
    ]
    for wall in walls[:2]:
        ifcopenshell.contain(wall, r1)
    ifcopenshell.contain(walls[2], local)            # contenu dans un local -> niveau RDC
    door = f.add("IfcDoor", GlobalId=guid(4), Name="Porte", psets={"BaseQuantities": {"Height": "deux mètres"}})
    ifcopenshell.contain(door, rdc)
    slab = f.add("IfcSlab", GlobalId=guid(5), Name="Dalle", psets={"BaseQuantities": {"GrossArea": 12.0}})
    ifcopenshell.contain(slab, building)            # directement dans le bâtiment -> sans niveau
    return f


class RulesTest(unittest.TestCase):
    def test_registry(self):
        self.assertEqual(rules.load_registry(RULES_XLSX),
                         {"ARC": "Architecture", "STR": "Structure", "CVC": "Chauffage Ventilation Climatisation"})

    def test_rules_filtered_by_corps_etat(self):
        arc = rules.load_rules(RULES_XLSX, "ARC")
        self.assertEqual(arc["classes"], ["IfcWall", "IfcDoor", "IfcWindow", "IfcCovering"])
        self.assertEqual(arc["psets"]["IfcWall"], {"Pset_WallCommon": {"AcousticRating": "String"}})
        self.assertEqual(arc["psets"]["IfcSlab"], {"BaseQuantities": {"GrossArea": "Number"}})  # cellule vide = tous
        self.assertNotIn("IfcBeam", arc["psets"])
        everything = rules.load_rules(RULES_XLSX)
        self.assertEqual(len(everything["classes"]), 9)
        self.assertEqual(everything["psets"]["IfcWall"]["Pset_WallCommon"],
                         {"LoadBearing": "Bool", "AcousticRating": "String"})

    def test_value_types(self):
        self.assertTrue(rules.value_matches("3", "Int"))
        self.assertTrue(rules.value_matches(2.0, "Number"))
        self.assertFalse(rules.value_matches("deux", "Float"))

    def test_unknown_type_is_reported(self):
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.title = "Element_Types"
        wb.active.append(["IFC_Class"])
        wb.active.append(["IfcWall"])
        ws = wb.create_sheet("Pset_WallCommon")
        ws.append(["IFC_Class", "Parametre", "Type"])
        ws.append(["IfcWall", "LoadBearing", "Booléen"])
        path = os.path.join(tempfile.mkdtemp(), "regles.xlsx")
        wb.save(path)
        with self.assertRaises(rules.RulesFileError):
            rules.load_rules(path)


class SpatialTest(unittest.TestCase):
    def test_levels(self):
        f = build_model()
        level = {e.GlobalId: spatial.level_of(e) for e in f.entities if e.GlobalId}
        self.assertEqual(level[guid(1)], "R+1")
        self.assertEqual(level[guid(3)], "RDC")         # via le local
        self.assertEqual(level[guid(5)], spatial.NO_LEVEL)

    def test_ifc4x3_facility_part(self):
        f = ifcopenshell.file("IFC4X3_ADD2")
        port = f.add("IfcMarineFacility", Name="Port")
        quai = f.add("IfcMarinePart", Name="Quai nord")
        ifcopenshell.aggregate(quai, port)
        bollard = f.add("IfcElement", GlobalId=guid(10))
        ifcopenshell.contain(bollard, quai)
        self.assertEqual(spatial.level_of(bollard), "Quai nord")


class PsetCheckTest(unittest.TestCase):
    def setUp(self):
        self.model = build_model()
        rule_set = rules.load_rules(RULES_XLSX, "ARC")
        rule_set["classes"].append("IfcInexistant")
        self.result = check_psets(self.model, rule_set, spatial.LevelCache())

    def test_counts_and_subclasses(self):
        # 3 murs (dont un IfcWallStandardCase, couvert par la règle IfcWall) + 1 porte
        self.assertEqual(self.result["checked"], 4)
        self.assertEqual(self.result["valid_guids"], [guid(1)])

    def test_problems(self):
        problems = {(f["guid"], f["param"]): f["problem"] for f in self.result["findings"]}
        self.assertEqual(problems[(guid(2), "AcousticRating")], MISSING_PARAM)   # chaîne vide = absente
        self.assertEqual(problems[(guid(3), "")], MISSING_PSET)
        self.assertTrue(problems[(guid(4), "Height")].startswith("Type incorrect"))
        self.assertIn(("RDC", "IfcWall", "Pset_WallCommon"), self.result["anomalies"])

    def test_unknown_class_is_a_warning(self):
        self.assertTrue(any("IfcInexistant" in w for w in self.result["warnings"]))


def fake_requirement(description, failures):
    return types.SimpleNamespace(failures=[{"element": e, "reason": "Valeur absente"} for e in failures],
                                 status=not failures, to_string=lambda *a: description)


def fake_spec(name, applicable, failed, cardinality="required", status=None):
    spec = types.SimpleNamespace(
        name=name, identifier=name[:3].upper(), description="", instructions="",
        minOccurs=0 if cardinality == "optional" else 1, maxOccurs=0 if cardinality == "prohibited" else "unbounded",
        applicable_entities=applicable, failed_entities=failed,
        requirements=[fake_requirement(f"{name} : exigence", failed)])
    spec.status = (not failed and (applicable or cardinality != "required")) if status is None else status
    return spec


class IdsTest(unittest.TestCase):
    def test_dangerous_files_are_refused(self):
        for raw in (b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]><ids/>',
                    "<ids/>".encode("utf-16"),
                    b'<ids><xs:pattern value="(a+)+"/></ids>'):
            with self.assertRaises(ids_check.IdsFileError):
                ids_check.check_ids_text(raw)

    def test_example_ids_files(self):
        counts = {}
        for name in ("exploitation_maintenance.ids", "conception_minimale.ids"):
            path = os.path.join(EXAMPLES, name)
            with open(path, "rb") as fh:
                ids_check.check_ids_text(fh.read())
            document = ids_check.load_ids(path)
            counts[name] = len(document.specifications)
            self.assertTrue(ids_check.ids_title(document).startswith("Exemple"))
        self.assertEqual(counts, {"exploitation_maintenance.ids": 7, "conception_minimale.ids": 6})

    def test_example_ids_structure(self):
        """Ordre des facettes imposé par ids.xsd dans <applicability>, versions
        IFC autorisées, noms d'entités en majuscules."""
        ns = "{http://standards.buildingsmart.org/IDS}"
        order = ["entity", "partOf", "classification", "attribute", "property", "material"]
        for name in ("exploitation_maintenance.ids", "conception_minimale.ids"):
            root = ET.parse(os.path.join(EXAMPLES, name)).getroot()
            for spec in root.iter(ns + "specification"):
                self.assertTrue(set(spec.get("ifcVersion").split()) <= {"IFC2X3", "IFC4", "IFC4X3_ADD2"})
                self.assertEqual([c.tag[len(ns):] for c in spec], ["applicability", "requirements"])
                tags = [c.tag[len(ns):] for c in spec.find(ns + "applicability")]
                self.assertEqual(tags, sorted(tags, key=order.index))
                for entity in spec.iter(ns + "entity"):
                    names = [v.get("value") for v in entity.iter() if v.get("value")] or \
                            [v.text for v in entity.iter(ns + "simpleValue")]
                    self.assertTrue(all(n == n.upper() and n.startswith("IFC") for n in names), names)

    def test_check_ids_and_verdict(self):
        f = build_model()
        walls = f.by_type("IfcWall")
        doors = f.by_type("IfcDoor")
        document = types.SimpleNamespace(
            validate=lambda *a, **k: None,
            specifications=[
                fake_spec("Murs", walls, walls[1:]),
                fake_spec("Portes", doors, []),
                fake_spec("Espaces verts", [], [], cardinality="optional", status=None),
            ])
        document.specifications[2].status = None  # ignorée (autre version IFC)
        result = ids_check.check_ids(document, f, spatial.LevelCache())
        self.assertEqual((result["applicable_total"], result["failed_total"]), (4, 2))
        self.assertEqual([s["status"] for s in result["specifications"]], ["non_conforme", "conforme", "ignoree"])
        self.assertEqual(recette.ids_verdict(result, recette.RecetteSettings()), recette.REJECTED)

        required_missing = types.SimpleNamespace(validate=lambda *a, **k: None,
                                                 specifications=[fake_spec("Équipements", [], [], status=False)])
        result = ids_check.check_ids(required_missing, f, spatial.LevelCache())
        self.assertEqual(result["cardinality_failures"], 1)


def run_full_recette(marking="CONFIDENTIEL"):
    model = build_model()
    model_path = os.path.join(tempfile.mkdtemp(), "maquette.ifc")
    with open(model_path, "w") as fh:
        fh.write("ISO-10303-21;")
    walls = model.by_type("IfcWall")
    document = types.SimpleNamespace(info={"title": "Exigences DOE"}, validate=lambda *a, **k: None,
                                     specifications=[fake_spec("Murs – repère", walls, walls[2:])])
    settings = recette.RecetteSettings(
        ids_path=os.path.join(EXAMPLES, "conception_minimale.ids"), rules_path=RULES_XLSX, corps_etat="ARC",
        lot="Lot 02 – Gros œuvre", phase="DOE – Dossier des ouvrages exécutés", tester="J. Martin",
        marking=marking)
    with mock.patch.object(ids_check, "load_ids", return_value=document):
        return recette.run_recette(model, settings, model_path=model_path)


class RecetteTest(unittest.TestCase):
    def test_full_recette(self):
        result = run_full_recette()
        self.assertEqual(result.pset_verdict, recette.REJECTED)        # 1 mur conforme sur 4 éléments
        self.assertEqual(result.ids_verdict, recette.REJECTED)         # 1 mur non conforme sur 3
        self.assertEqual(result.verdict, recette.REJECTED)
        self.assertEqual(result.valid_guids, {guid(1)})                 # conforme à l'IDS et aux règles
        self.assertIn(guid(3), result.invalid_guids)
        self.assertEqual(len(result.model["file"]["sha256"]), 64)
        sources = {a["source"] for a in result.anomalies}
        self.assertEqual(sources, {"IDS", "PSet"})

    def test_thresholds(self):
        settings = recette.RecetteSettings(threshold_ok=1.0, threshold_reserve=0.5)
        self.assertEqual(recette.rate_verdict(0.99, settings), recette.RESERVED)
        self.assertEqual(recette.rate_verdict(1.0, settings), recette.VALIDATED)

    def test_requires_a_referential(self):
        with self.assertRaises(ValueError):
            recette.run_recette(build_model(), recette.RecetteSettings())


class DeliverablesTest(unittest.TestCase):
    def setUp(self):
        self.result = run_full_recette()
        self.dir = tempfile.mkdtemp()

    def test_report(self):
        path = os.path.join(self.dir, "recette.xlsx")
        report.write_report(self.result, path)
        wb = load_workbook(path)
        self.assertEqual(wb.sheetnames, ["Fiche de recette", "IDS", "Détail PSet"])
        fiche = wb["Fiche de recette"]
        texts = [str(c.value) for row in fiche.iter_rows() for c in row if c.value is not None]
        self.assertEqual(texts[0], "CONFIDENTIEL")
        self.assertIn("VERDICT : REJETÉ", texts)
        self.assertIn("Lot 02 – Gros œuvre", texts)
        self.assertIn(self.result.model["file"]["sha256"], texts)
        for ws in wb.worksheets:
            self.assertEqual(ws.oddHeader.center.text, "CONFIDENTIEL")
        details = wb["Détail PSet"]
        names = [c.value for c in details["D"]]
        self.assertIn("'=HYPERLINK(\"x\")", names)                      # formule neutralisée
        self.assertTrue(all(c.data_type != "f" for ws in wb.worksheets for row in ws.iter_rows() for c in row))

    def test_report_without_marking(self):
        result = run_full_recette(marking="")
        path = os.path.join(self.dir, "recette.xlsx")
        report.write_report(result, path)
        fiche = load_workbook(path)["Fiche de recette"]
        self.assertEqual(fiche["A1"].value, "FICHE DE RECETTE DE MAQUETTE NUMÉRIQUE (IFC)")

    def test_bcf(self):
        path = os.path.join(self.dir, "anomalies.bcf")
        stats = bcf.write_bcf(path, self.result.anomalies, author="recette@exemple.fr",
                              project_name="Projet de démonstration", lot="Lot 02", phase="DOE",
                              referentials=recette.describe_referentials(self.result), marking="CONFIDENTIEL")
        self.assertEqual(stats["topics"], len(self.result.anomalies))
        with zipfile.ZipFile(path) as zf:
            markups = [ET.fromstring(zf.read(n)) for n in zf.namelist() if n.endswith("markup.bcf")]
            topic = markups[0].find("Topic")
            self.assertIn("CONFIDENTIEL", [l.text for l in topic.findall("Labels")])
            self.assertTrue(topic.findtext("Description").startswith("CONFIDENTIEL"))
            if XSD_DIR:
                from lxml import etree
                xsd = {".bcf": "markup.xsd", ".bcfv": "visinfo.xsd", ".version": "version.xsd", ".bcfp": "project.xsd"}
                for name in zf.namelist():
                    schema = etree.XMLSchema(etree.parse(os.path.join(XSD_DIR, xsd[os.path.splitext(name)[1]])))
                    self.assertTrue(schema.validate(etree.fromstring(zf.read(name))), f"{name}: {schema.error_log}")


class OfflineGuaranteeTest(unittest.TestCase):
    """Aucun module de l'extension n'importe de bibliothèque réseau, et le
    manifeste ne demande pas la permission « network »."""
    NETWORK_MODULES = {"socket", "ssl", "urllib", "http", "requests", "httpx", "aiohttp", "ftplib", "smtplib",
                       "webbrowser", "xmlrpc", "telnetlib", "poplib", "imaplib"}

    def test_no_network_imports(self):
        for root, _, files in os.walk(PACKAGE):
            for name in files:
                if name.endswith(".py"):
                    with open(os.path.join(root, name), encoding="utf-8") as fh:
                        tree = ast.parse(fh.read())
                    for node in ast.walk(tree):
                        modules = [a.name for a in node.names] if isinstance(node, ast.Import) else \
                                  [node.module or ""] if isinstance(node, ast.ImportFrom) else []
                        for module in modules:
                            self.assertNotIn(module.split(".")[0], self.NETWORK_MODULES, f"{name} : {module}")

    def test_manifest_has_no_network_permission(self):
        with open(os.path.join(PACKAGE, "blender_manifest.toml"), encoding="utf-8") as fh:
            manifest = fh.read()
        permissions = manifest.split("[permissions]")[1].split("[")[0]
        self.assertNotIn("network", [line.split("=")[0].strip() for line in permissions.splitlines()])


class ExtensionTest(unittest.TestCase):
    """Enregistrement et opérateurs, avec bpy et Bonsai simulés."""

    def setUp(self):
        import bpy
        import bonsai.tool as tool
        import checkers_bonsai
        from checkers_bonsai import operators
        self.bpy, self.tool, self.ext, self.ops = bpy, tool, checkers_bonsai, operators
        checkers_bonsai.register()
        self.model = build_model()
        tool.Ifc.file, tool.Ifc.path = self.model, ""

        class Obj:
            def __init__(self, name):
                self.name, self.color, self.selected = name, (1.0, 1.0, 1.0, 1.0), False

            def select_set(self, value):
                self.selected = value

            def visible_get(self):
                return True

        tool.Ifc.objects = {e.GlobalId: Obj(e.GlobalId) for e in self.model.entities
                            if e.is_a("IfcElement")}
        bpy.data.objects = dict((o.name, o) for o in tool.Ifc.objects.values())

        class Coll(list):
            def add(self):
                self.append(types.SimpleNamespace())
                return self[-1]

            def clear(self):
                del self[:]

        from checkers_bonsai.props import CheckersProperties
        props = types.SimpleNamespace(
            ids_path="", rules_path=RULES_XLSX, corps_etat="ARC", lot="Lot 02", phase="DOE", tester="J. Martin",
            marking="confidentiel", threshold_ok=0.98, threshold_reserve=0.9, bcf_author="",
            anomalies=Coll(), anomaly_index=0)
        props.marking_text = types.MethodType(CheckersProperties.marking_text, props)
        self.props = props
        space = types.SimpleNamespace(type="VIEW_3D", shading=types.SimpleNamespace(color_type="MATERIAL"),
                                      as_pointer=lambda: 1)
        area = types.SimpleNamespace(type="VIEW_3D", spaces=[space], regions=[types.SimpleNamespace(type="WINDOW")])
        import contextlib
        self.space = space
        self.context = types.SimpleNamespace(
            scene=types.SimpleNamespace(checkers=props), area=area, selected_objects=[],
            view_layer=types.SimpleNamespace(objects=types.SimpleNamespace(active=None)),
            window_manager=types.SimpleNamespace(
                progress_begin=lambda *a: None, progress_end=lambda: None,
                windows=[types.SimpleNamespace(screen=types.SimpleNamespace(areas=[area]))]),
            temp_override=lambda **k: contextlib.nullcontext())

    def tearDown(self):
        self.ext.unregister()
        self.assertEqual(self.bpy.registered, [])

    def test_workflow(self):
        ops, ctx = self.ops, self.context
        self.assertTrue(ops.CHECKERS_OT_run_recette.poll(ctx))
        self.assertEqual(ops.CHECKERS_OT_run_recette().execute(ctx), {"FINISHED"})
        self.assertEqual(self.props.verdict, recette.REJECTED)
        self.assertGreater(len(self.props.anomalies), 0)

        select = ops.CHECKERS_OT_select_anomaly()
        select.index = 0
        self.assertEqual(select.execute(ctx), {"FINISHED"})
        self.assertTrue(any(o.selected for o in self.tool.Ifc.objects.values()))

        ops.CHECKERS_OT_color_results().execute(ctx)
        self.assertEqual(self.space.shading.color_type, "OBJECT")
        self.assertEqual(self.tool.Ifc.objects[guid(1)].color, ops.COLOR_VALID)
        self.assertEqual(self.tool.Ifc.objects[guid(5)].color, (1.0, 1.0, 1.0, 1.0))   # dalle non contrôlée (ARC)
        ops.CHECKERS_OT_reset_colors().execute(ctx)
        self.assertEqual(self.space.shading.color_type, "MATERIAL")

        for op_class, ext in ((ops.CHECKERS_OT_export_report, ".xlsx"), (ops.CHECKERS_OT_export_bcf, ".bcf")):
            op = op_class()
            op.invoke(ctx, None)
            self.assertTrue(op.filepath.startswith("Recette_Lot_02_") and op.filepath.endswith(ext), op.filepath)
            op.filepath = os.path.join(tempfile.mkdtemp(), "sortie" + ext)
            self.assertEqual(op.execute(ctx), {"FINISHED"})
            self.assertTrue(os.path.getsize(op.filepath) > 0)

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


if __name__ == "__main__":
    unittest.main()
