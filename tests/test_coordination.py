"""Tests de la coordination entre lots : cohérence de référence
(core/federation.py) et conflits géométriques avec IfcClash (core/clashes.py).

    python -m unittest discover tests

IfcClash est remplacé par une doublure (tests/fakes/ifcclash) qui travaille
sur des boîtes alignées sur les axes ; le vrai IfcClash est testé dans
tests_integration/.
"""
import json
import os
import sys
import tempfile
import types
import unittest
import zipfile
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import support  # noqa: E402,F401
from support import build_model, guid  # noqa: E402

import ifcopenshell  # noqa: E402  (doublure)
from openpyxl import load_workbook  # noqa: E402

from checkers_bonsai.core import (bcf, clashes, federation, journal, recette, report, spatial,  # noqa: E402
                                  tracking)
from checkers_bonsai.core.structure import BLOCKING, MAJOR, MINOR  # noqa: E402

BEAM, DUCT, DUCT_FAR = guid(300), guid(301), guid(302)


def frame(model, origin=(0.0, 0.0, 0.0), angle=0.0, rdc=0.0, r1=3.0, georef=True, eastings=652000.0, north=0.0):
    """Donne au modèle de build_model un repère : origine du bâtiment,
    altitudes d'étages, géoréférencement et nord."""
    ifcopenshell.place(model.by_guid(guid(901)), *origin, angle=angle)
    ifcopenshell.place(model.by_guid(guid(902)), origin[0], origin[1], origin[2] + rdc)
    ifcopenshell.place(model.by_guid(guid(903)), origin[0], origin[1], origin[2] + r1)
    if georef:
        model.add("IfcMapConversion", Eastings=eastings, Northings=6862000.0, OrthogonalHeight=35.0,
                  XAxisAbscissa=1.0, XAxisOrdinate=0.0, Scale=1.0,
                  TargetCRS=types.SimpleNamespace(Name="EPSG:2154"))
    import math
    model.add("IfcGeometricRepresentationContext", ContextType="Model",
              TrueNorth=types.SimpleNamespace(DirectionRatios=(math.sin(math.radians(north)),
                                                               math.cos(math.radians(north)))))
    return model


def other_lot(folder, name="STR", **frame_args):
    """Maquette d'un autre lot, même projet et même repère par défaut ; ses
    éléments ont d'autres GlobalId que ceux de build_model."""
    model = frame(build_model(), **frame_args)
    for element in model.by_type("IfcElement"):
        element.GlobalId = "1" + element.GlobalId[1:]
    path = ifcopenshell.register(os.path.join(folder, f"{name}.ifc"), model)
    return model, path


def lots(folder, *specs):
    return federation.open_lots([(name, path) for name, path in specs])


class FederationTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.model = frame(build_model())

    def check(self, **frame_args):
        _, path = other_lot(self.dir, **frame_args)
        result = federation.check_federation(self.model, lots(self.dir, ("STR", path)), spatial.LevelCache())
        return result, {c["code"]: c for c in result["checks"]}

    def test_aligned_lots(self):
        result, checks = self.check()
        self.assertEqual({c["status"] for c in checks.values()}, {"ok"}, result["checks"])
        self.assertTrue(result["frame_ok"])
        self.assertEqual(recette.severity_verdict(result), recette.VALIDATED)
        self.assertEqual(result["lots"][0]["name"], "STR")
        self.assertEqual(len(result["lots"][0]["file"]["sha256"]), 64)

    def test_shifted_origin_and_storeys(self):
        result, checks = self.check(origin=(0.0, 0.25, 0.0), rdc=0.0, r1=3.2)
        self.assertEqual(checks["C04"]["status"], "ko")
        self.assertIn("0.250 m", checks["C04"]["message"])
        self.assertEqual(checks["C05"]["status"], "ko")
        self.assertIn("R+1 : +3.000 m contre +3.200 m", checks["C05"]["message"])
        self.assertFalse(result["frame_ok"])
        self.assertEqual(recette.severity_verdict(result), recette.RESERVED)
        key = ("R+1", "IfcBuildingStorey", "Cohérence avec « STR » : Étages : mêmes altitudes")
        self.assertEqual(result["anomalies"][key]["guids"], {guid(903)})

    def test_rotation_georeferencing_north_and_units(self):
        other, path = other_lot(self.dir, angle=1.0, eastings=652000.5, north=2.0)
        other.unit_scale = 0.001
        result = federation.check_federation(self.model, lots(self.dir, ("STR", path)), spatial.LevelCache())
        checks = {c["code"]: c for c in result["checks"]}
        self.assertIn("tournée de 1.000°", checks["C04"]["message"])
        self.assertIn("décalé de 0.500", checks["C07"]["message"])
        self.assertIn("nord à 0.000° contre 2.000°", checks["C08"]["message"])
        self.assertEqual((checks["C03"]["status"], checks["C03"]["severity"]), ("ko", MINOR))
        self.assertIn("mètre contre millimètre", checks["C03"]["message"])

    def test_georeferencing_in_one_lot_only(self):
        _, checks = self.check(georef=False)
        self.assertEqual(checks["C07"]["message"], "géoréférencement présent dans un seul des deux lots")

    def test_storey_named_differently_and_project(self):
        other, path = other_lot(self.dir)
        other.by_guid(guid(902)).Name = "Rez-de-chaussée"
        other.by_type("IfcProject")[0].GlobalId = guid(990)
        result = federation.check_federation(self.model, lots(self.dir, ("STR", path)), spatial.LevelCache())
        checks = {c["code"]: c for c in result["checks"]}
        self.assertIn("« RDC » / « Rez-de-chaussée »", checks["C06"]["message"])
        self.assertEqual((checks["C02"]["status"], checks["C02"]["severity"]), ("ko", MINOR))  # même nom
        other.by_type("IfcProject")[0].Name = "Autre projet"
        result = federation.check_federation(self.model, lots(self.dir, ("STR", path)), spatial.LevelCache())
        self.assertEqual({c["code"]: c for c in result["checks"]}["C02"]["severity"], MAJOR)

    def test_shared_elements_are_blocking_but_not_spatial_structure(self):
        other, path = other_lot(self.dir)
        other.by_type("IfcDoor")[0].GlobalId = guid(4)                   # même porte dans les deux lots
        result = federation.check_federation(self.model, lots(self.dir, ("STR", path)), spatial.LevelCache())
        checks = {c["code"]: c for c in result["checks"]}
        self.assertEqual((checks["C09"]["status"], checks["C09"]["severity"]), ("ko", BLOCKING))
        self.assertEqual(checks["C09"]["count"], 1)                       # étages et bâtiment partagés : admis
        self.assertEqual(recette.severity_verdict(result), recette.REJECTED)

    def test_lot_list_errors(self):
        _, path = other_lot(self.dir)
        for specs, message in (([("STR", path), ("STR", path)], "Deux lots"),
                               ([("Ce modèle", path)], "réservé"),
                               ([("CVC", os.path.join(self.dir, "absent.ifc"))], "introuvable")):
            with self.assertRaises(federation.FederationError) as raised:
                federation.open_lots(specs)
            self.assertIn(message, str(raised.exception))
        self.assertEqual(federation.open_lots([("", path)])[0]["name"], "STR")    # nom du fichier par défaut


def add_boxes(model, beam_owner=None):
    """Une gaine qui traverse une poutre (de l'autre lot, ou du même), et une
    gaine éloignée."""
    rdc = model.by_guid(guid(902))
    duct = model.add("IfcDuctSegment", GlobalId=DUCT, Name="Gaine soufflage", _box=((0, 0, 2.6), (6, 0.4, 3.0)))
    far = model.add("IfcDuctSegment", GlobalId=DUCT_FAR, Name="Gaine reprise", _box=((0, 5, 2.6), (6, 5.4, 3.0)))
    for element in (duct, far):
        ifcopenshell.contain(element, rdc)
    if beam_owner is not None:
        beam = beam_owner.add("IfcBeam", GlobalId=BEAM, Name="Poutre", _box=((2, -1, 2.5), (2.3, 3, 3.0)))
        ifcopenshell.contain(beam, beam_owner.by_guid(guid(902)))
    return model


class ClashTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.model = add_boxes(frame(build_model()))
        self.other, self.path = other_lot(self.dir)
        beam = self.other.add("IfcBeam", GlobalId=BEAM, Name="Poutre", _box=((2, -1, 2.5), (2.3, 3, 3.0)))
        ifcopenshell.contain(beam, self.other.by_guid(guid(902)))
        self.lots = lots(self.dir, ("STR", self.path))

    def run_rules(self, *rules):
        return clashes.run_clashes(self.model, self.lots, list(rules), spatial.LevelCache())

    def test_duct_through_beam_of_other_lot(self):
        result = self.run_rules(clashes.ClashRule(other="STR", selector_a="IfcDuctSegment", selector_b="IfcBeam"))
        self.assertEqual(result["total"], 1)
        self.assertEqual(result["rules"][0]["status"], "ko")
        key = ("RDC", "IfcDuctSegment", "Conflit avec « STR » (intersection, 1 cm) : IfcBeam")
        anomaly = result["anomalies"][key]
        self.assertEqual((anomaly["guids"], anomaly["related_guids"]), ({DUCT}, {BEAM}))
        self.assertEqual(anomaly["severity"], MAJOR)
        self.assertEqual(anomaly["point"], (2.15, 0.2, 2.8))
        detail = result["details"][0]
        self.assertEqual((detail["a_guid"], detail["b_guid"], detail["type"]), (DUCT, BEAM, "dépassement"))
        self.assertEqual(recette.severity_verdict(result), recette.RESERVED)

    def test_modes_and_tolerance(self):
        # Intersection de 0,4 m : ignorée avec une tolérance de 0,5 m.
        loose = self.run_rules(clashes.ClashRule(other="STR", selector_a="IfcDuctSegment", selector_b="IfcBeam",
                                                 distance=0.5))
        self.assertEqual(loose["total"], 0)
        self.assertEqual(recette.severity_verdict(loose), recette.VALIDATED)
        # Dégagement de 5 m exigé entre gaines : les deux gaines, à 4,6 m l'une de l'autre, sont trop proches.
        clearance = self.run_rules(clashes.ClashRule(other="Ce modèle", selector_a="IfcDuctSegment",
                                                     selector_b="IfcDuctSegment", mode="clearance", distance=5.0,
                                                     severity=MINOR))
        self.assertEqual(clearance["total"], 1)                           # paire comptée une seule fois
        anomaly = next(iter(clearance["anomalies"].values()))
        self.assertEqual(anomaly["guids"], {DUCT, DUCT_FAR})               # interne : les deux éléments
        self.assertEqual(anomaly["related_guids"], set())
        self.assertEqual(recette.severity_verdict(clearance), recette.VALIDATED)   # mineur

    def test_rules_that_cannot_run(self):
        result = self.run_rules(clashes.ClashRule(other="STR", selector_a="IfcPump"),             # rien de sélectionné
                                clashes.ClashRule(other="STR", selector_a="IfcWall"))             # pas de boîte
        messages = [r["message"] for r in result["rules"]]
        self.assertIn("aucun élément avec géométrie dans la sélection de la maquette ouverte", messages)
        self.assertIn("géométrie non calculable pour les éléments sélectionnés", messages)
        self.assertEqual(result["errors"], 2)
        self.assertEqual(recette.severity_verdict(result), recette.RESERVED)
        with self.assertRaises(clashes.ClashError):
            self.run_rules(clashes.ClashRule(other="CVC"))                                         # lot inconnu
        bad = self.run_rules(clashes.ClashRule(other="STR", selector_a="pas une requête"))   # règle en erreur seule
        self.assertIn("illisible", bad["rules"][0]["message"])

    def test_ifcclash_missing(self):
        from unittest import mock
        with mock.patch.dict(sys.modules, {"ifcclash": None, "ifcclash.ifcclash": None}):
            with self.assertRaises(clashes.ClashError) as raised:
                self.run_rules(clashes.ClashRule(other="STR"))
        self.assertIn("Bonsai", str(raised.exception))


class CoordinationRecetteTest(unittest.TestCase):
    """Cohérence et conflits dans une recette complète, avec les livrables."""

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.model = add_boxes(frame(build_model()))
        self.other, self.path = other_lot(self.dir, r1=3.2)               # altitude de R+1 différente
        beam = self.other.add("IfcBeam", GlobalId=BEAM, Name="Poutre", _box=((2, -1, 2.5), (2.3, 3, 3.0)))
        ifcopenshell.contain(beam, self.other.by_guid(guid(902)))

    def recette(self, **values):
        settings = dict(check_structure=False, lot="CVC", lots=[("STR", self.path)],
                        clash_rules=[clashes.ClashRule(other="STR", selector_a="IfcDuctSegment", selector_b="IfcBeam")])
        settings.update(values)
        return recette.run_recette(self.model, recette.RecetteSettings(**settings))

    def test_recette(self):
        result = self.recette()
        self.assertEqual(result.sources, {"Cohérence", "Conflit"})
        self.assertEqual((result.federation_verdict, result.clash_verdict, result.verdict),
                         (recette.RESERVED, recette.RESERVED, recette.RESERVED))
        self.assertTrue(any("pas tous dans le même repère" in w for w in result.warnings))
        conflict = next(a for a in result.anomalies if a["source"] == "Conflit")
        self.assertEqual((conflict["guids"], conflict["related_guids"]), ([DUCT], [BEAM]))
        self.assertEqual(result.invalid_guids, {DUCT, guid(903)})
        self.assertIn("coordination avec « STR »", recette.describe_referentials(result))
        self.assertEqual(result.lots[0]["name"], "STR")

    def test_federation_only_and_internal_clashes_only(self):
        result = self.recette(clash_rules=[])
        self.assertEqual(result.sources, {"Cohérence"})
        result = self.recette(lots=[], clash_rules=[clashes.ClashRule(selector_a="IfcDuctSegment")])
        self.assertEqual(result.sources, {"Conflit"})
        self.assertEqual(result.clashes["total"], 0)
        self.assertFalse(any("même repère" in w for w in result.warnings))
        with self.assertRaises(clashes.ClashError):
            self.recette(lots=[])                                             # la règle vise un lot absent

    def test_deliverables(self):
        result = self.recette()
        folder = tempfile.mkdtemp()
        path = os.path.join(folder, "fiche.xlsx")
        report.write_report(result, path)
        workbook = load_workbook(path)
        self.assertEqual(workbook.sheetnames, ["Fiche de recette", "Cohérence", "Conflits"])
        texts = [str(c.value) for row in workbook["Fiche de recette"].iter_rows() for c in row if c.value is not None]
        self.assertTrue(any(t.startswith("Cohérence entre lots (1 lot(s))") for t in texts))
        self.assertTrue(any(t.startswith("Conflits géométriques (1 conflit(s))") for t in texts))
        self.assertIn("Lot « STR » (SHA-256)", texts)
        clash_rows = list(workbook["Conflits"].iter_rows(values_only=True))
        self.assertTrue(any(row[3] == DUCT and row[7] == BEAM for row in clash_rows))

        bcf_path = os.path.join(folder, "coordination.bcf")
        stats = bcf.write_bcf(bcf_path, result.anomalies, project_key="p", lot="CVC")
        with zipfile.ZipFile(bcf_path) as zf:
            names = zf.namelist()
            viewpoints = {n: ET.fromstring(zf.read(n)) for n in names if n.endswith(".bcfv")}
            markups = {n: ET.fromstring(zf.read(n)) for n in names if n.endswith(".bcf")}
            conflict = next(n for n, m in markups.items() if m.find("Topic/Labels[.='Conflit']") is not None)
            viewpoint = viewpoints[conflict.replace("markup.bcf", "viewpoint.bcfv")]
            selected = [c.get("IfcGuid") for c in viewpoint.find("Components/Selection")]
            colored = [c.get("IfcGuid") for c in viewpoint.find("Components/Coloring/Color")]
            self.assertEqual((selected, colored), ([DUCT, BEAM], [DUCT]))   # l'élément de l'autre lot sélectionné seulement
            camera = viewpoint.find("PerspectiveCamera")
            self.assertIsNotNone(camera)
            eye = [float(camera.findtext(f"CameraViewPoint/{a}")) for a in "XYZ"]
            self.assertEqual(eye, [2.15 - 4.0, 0.2 - 4.0, 2.8 + 3.0])
            if support.BCF_XSD_DIR:
                from lxml import etree
                xsd = {".bcf": "markup.xsd", ".bcfv": "visinfo.xsd", ".version": "version.xsd", ".bcfp": "project.xsd"}
                for name in names:
                    schema = etree.XMLSchema(etree.parse(os.path.join(support.BCF_XSD_DIR,
                                                                      xsd[os.path.splitext(name)[1]])))
                    self.assertTrue(schema.validate(etree.fromstring(zf.read(name))), f"{name}: {schema.error_log}")
        self.assertEqual(stats["components"], 3)                          # gaine + poutre, étage R+1

    def test_journal_and_tracking(self):
        result = self.recette()
        path = os.path.join(self.dir, "recette.json")
        journal.write_record(result, path)
        record = journal.load_record(path)
        self.assertEqual(record["referentials"]["lots"][0]["lot"], "STR")
        self.assertEqual(record["verdicts"]["clashes"], recette.RESERVED)
        self.assertNotIn("related_guids", json.dumps(record["anomalies"]))   # le journal ne garde que cette maquette

        # Livraison suivante, avec une nouvelle version du lot STR : signalé.
        with open(self.path, "a", encoding="utf-8") as fh:
            fh.write("\n/* version 2 */")
        second = self.recette(previous_path=path)
        self.assertTrue(any("lot « STR » a changé" in w for w in second.evolution["warnings"]))
        self.assertEqual(second.evolution["persistent_count"], 2)

        table = os.path.join(self.dir, "suivi.xlsx")
        records, _ = tracking.collect_journals(self.dir)
        tracking.write_tracking(table, records)
        header = [c.value for c in load_workbook(table)["Recettes"][1]]
        row = [c.value for c in load_workbook(table)["Recettes"][2]]
        self.assertEqual(row[header.index("Conflits")], recette.RESERVED)
        self.assertEqual(row[header.index("Cohérence entre lots")], recette.RESERVED)

    def test_lots_record_validation(self):
        record = journal.to_record(self.recette())
        record["referentials"]["lots"] = "STR.ifc"
        path = os.path.join(self.dir, "mauvais.json")
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(record, fh)
        with self.assertRaises(journal.JournalError):
            journal.load_record(path)


if __name__ == "__main__":
    unittest.main()
