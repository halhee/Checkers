"""Tests des contrôles : règles Excel, PSet, niveaux, structure et IDS.

    python -m unittest discover tests
"""
import os
import sys
import tempfile
import types
import unittest
from unittest import mock
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import support  # noqa: E402,F401  (met en place sys.path et les doublures)
from support import EXAMPLES, RULES_XLSX, build_broken_model, build_model, fake_document, fake_spec, guid

import ifcopenshell
import ifcopenshell.util.unit

from checkers_bonsai.core import ids as ids_check, recette, rules, spatial, structure
from checkers_bonsai.core.psets import MISSING_PARAM, MISSING_PSET, check_psets


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

    def test_value_types_are_strict(self):
        accepted = [(True, "Bool"), (False, "bool"), (3, "Int"), (3.0, "Int"), (2, "Float"), (2.5, "Number"),
                    ("Rw 45", "String"), (["A", "B"], "String"), ({"id": 12, "UpperBoundValue": 2.5}, "Float")]
        refused = [("oui", "Bool"), (1, "Bool"), (3.5, "Int"), ("3", "Int"), (True, "Int"), ("2.5", "Float"),
                   (True, "Number"), (3, "String"), (["A", 2], "String")]
        for value, kind in accepted:
            self.assertTrue(rules.value_matches(value, kind), (value, kind))
        for value, kind in refused:
            self.assertFalse(rules.value_matches(value, kind), (value, kind))

    def test_empty_values(self):
        for value in (None, "", "   ", [], [None, " "], {"id": 3}):
            self.assertTrue(rules.is_empty(value), value)
        for value in (0, False, "x", [None, "x"]):
            self.assertFalse(rules.is_empty(value), value)

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

    def test_unsafe_archives_are_refused(self):
        not_a_workbook = os.path.join(tempfile.mkdtemp(), "regles.xlsx")
        with open(not_a_workbook, "w") as fh:
            fh.write("ceci n'est pas un classeur")
        with self.assertRaises(rules.RulesFileError):
            rules.load_rules(not_a_workbook)
        # Contenu décompressé au-delà du plafond (bombe de décompression).
        with mock.patch.object(rules, "MAX_UNCOMPRESSED_BYTES", 100), self.assertRaises(rules.RulesFileError):
            rules.load_rules(RULES_XLSX)
        with self.assertRaises(rules.RulesFileError):
            rules.load_rules(os.path.join(tempfile.mkdtemp(), "absent.xlsx"))


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

    def test_nested_element(self):
        f = build_model()
        wall = f.by_guid(guid(1))
        accessory = f.add("IfcElement", GlobalId=guid(11), Nests=[types.SimpleNamespace(RelatingObject=wall)])
        self.assertEqual(spatial.level_of(accessory), "R+1")


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
        self.assertEqual(problems[(guid(4), "Height")], "Type incorrect (attendu Number, trouvé texte)")
        self.assertIn(("RDC", "IfcWall", "Pset_WallCommon"), self.result["anomalies"])

    def test_unknown_class_is_a_warning(self):
        self.assertTrue(any("IfcInexistant" in w for w in self.result["warnings"]))

    def test_most_specific_rule_class_wins(self):
        rule_set = {"classes": ["IfcWall", "IfcBuildingElement"],
                    "psets": {"IfcWall": {"Pset_WallCommon": {"LoadBearing": "Bool"}},
                              "IfcBuildingElement": {"Pset_Inexistant": {"X": "String"}}}}
        result = check_psets(self.model, rule_set, spatial.LevelCache())
        self.assertIn(guid(1), result["valid_guids"])             # mur jugé sur la règle IfcWall
        self.assertIn(("Sans niveau", "IfcSlab", "Pset_Inexistant"), result["anomalies"])


class StructureTest(unittest.TestCase):
    def test_sound_model(self):
        result = structure.check_structure(build_model(), spatial.LevelCache())
        self.assertEqual([c["status"] for c in result["checks"]], ["ok"] * len(structure.CHECKS))
        self.assertEqual(result["anomalies"], {})
        self.assertEqual(recette.structure_verdict(result), recette.VALIDATED)

    def test_broken_model(self):
        model, parts = build_broken_model()
        result = structure.check_structure(model, spatial.LevelCache())
        failed = {c["code"]: c["severity"] for c in result["checks"] if c["status"] == "ko"}
        self.assertEqual(failed, {
            "S01": structure.BLOCKING, "S02": structure.BLOCKING, "S03": structure.MAJOR, "S05": structure.MAJOR,
            "S06": structure.MAJOR, "S07": structure.MAJOR, "S08": structure.MINOR, "S09": structure.MINOR,
            "S10": structure.MINOR, "S11": structure.MINOR})
        self.assertEqual((result["blocking"], result["major"], result["minor"]), (2, 4, 4))
        self.assertEqual(recette.structure_verdict(result), recette.REJECTED)

        anomalies = result["anomalies"]
        orphans = anomalies[("Sans niveau", "IfcWall", "Éléments rattachés à la structure spatiale")]
        self.assertEqual(orphans["guids"], {guid(1)})             # ni la réservation, ni la volée d'escalier
        bare = anomalies[("Niveau", "IfcBeam", "Éléments avec géométrie")]
        self.assertEqual(bare["guids"], {guid(4)})                # pas l'escalier, qui est un assemblage
        self.assertEqual(anomalies[("Modèle", "IfcProject", "Unité de longueur déclarée")]["count"], 1)
        self.assertEqual(anomalies[("Modèle", "IfcBuilding", "Bâtiments découpés en étages")]["guids"], {guid(901)})

    def test_proxy_share_sets_severity(self):
        model = build_model()
        storey = model.by_guid(guid(902))
        for n in range(30, 60):                                   # dilue le proxy sous le seuil de 10 %
            ifcopenshell.contain(model.add("IfcWall", GlobalId=guid(n), Name=f"Mur {n}"), storey)
        ifcopenshell.contain(model.add("IfcBuildingElementProxy", GlobalId=guid(60), Name="Objet"), storey)
        result = structure.check_structure(model, spatial.LevelCache())
        proxies = next(c for c in result["checks"] if c["code"] == "S07")
        self.assertEqual((proxies["status"], proxies["severity"]), ("ko", structure.MINOR))
        self.assertEqual(recette.structure_verdict(result), recette.VALIDATED)

    def test_ifc2x3_and_failing_check(self):
        # IfcFacility et IfcMapConversion n'existent pas en IFC2X3 : sans effet.
        result = structure.check_structure(build_model("IFC2X3"), spatial.LevelCache())
        self.assertTrue(all(c["status"] == "ok" for c in result["checks"]))
        # Un contrôle qui plante est signalé sans faire perdre les autres.
        with mock.patch.object(ifcopenshell.util.unit, "get_project_unit", side_effect=RuntimeError("boum")):
            result = structure.check_structure(build_model(), spatial.LevelCache())
        statuses = {c["code"]: c["status"] for c in result["checks"]}
        self.assertEqual(statuses["S03"], "erreur")
        self.assertEqual(sum(1 for s in statuses.values() if s == "ok"), len(structure.CHECKS) - 1)


class IdsTest(unittest.TestCase):
    def test_dangerous_files_are_refused(self):
        for raw in (b'<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]><ids/>',
                    "<ids/>".encode("utf-16"),
                    b'<ids><xs:pattern value="(a+)+"/></ids>',
                    b"<ids>" + b" " * (ids_check.MAX_IDS_BYTES + 1)):
            with self.assertRaises(ids_check.IdsFileError):
                ids_check.check_ids_text(raw)
        with self.assertRaises(ids_check.IdsFileError):
            ids_check.load_ids(os.path.join(tempfile.mkdtemp(), "absent.ids"))

    def test_example_ids_files(self):
        counts = {}
        for name in ("exploitation_maintenance.ids", "conception_minimale.ids"):
            document = ids_check.load_ids(os.path.join(EXAMPLES, name))
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
        walls, doors = f.by_type("IfcWall"), f.by_type("IfcDoor")
        ignored = fake_spec("Espaces verts", [], [], cardinality="optional")
        ignored.status = None                                      # vise une autre version IFC
        document = fake_document(fake_spec("Murs", walls, walls[1:]), fake_spec("Portes", doors, []), ignored)
        result = ids_check.check_ids(document, f, spatial.LevelCache())
        self.assertEqual((result["applicable_total"], result["failed_total"]), (4, 2))
        self.assertEqual([s["status"] for s in result["specifications"]], ["non_conforme", "conforme", "ignoree"])
        self.assertEqual(result["applicable_guids"], {guid(1), guid(2), guid(3), guid(4)})
        self.assertEqual(len(result["details"]), 2)
        self.assertEqual(recette.ids_verdict(result, recette.RecetteSettings()), recette.REJECTED)

        required_missing = fake_document(fake_spec("Équipements", [], [], status=False))
        result = ids_check.check_ids(required_missing, f, spatial.LevelCache())
        self.assertEqual(result["cardinality_failures"], 1)
        self.assertEqual(recette.ids_verdict(result, recette.RecetteSettings()), recette.REJECTED)

    def test_prohibited_specification(self):
        f = build_model()
        doors = f.by_type("IfcDoor")
        document = fake_document(fake_spec("Aucune porte", doors, [], cardinality="prohibited", status=False))
        result = ids_check.check_ids(document, f, spatial.LevelCache())
        self.assertEqual(result["failed_total"], 1)
        self.assertIn(("RDC", "IfcDoor", "IDS : Aucune porte"), result["anomalies"])

    def test_unreadable_failure_details_keep_counts(self):
        """Si une version d'ifctester présente le détail autrement, les
        compteurs et la sélection restent exacts, avec un avertissement."""
        f = build_model()
        walls = f.by_type("IfcWall")
        spec = fake_spec("Murs", walls, walls[:1])
        spec.requirements[0].failures = [object()]
        result = ids_check.check_ids(fake_document(spec), f, spatial.LevelCache())
        self.assertEqual((result["failed_total"], result["details"]), (1, []))
        self.assertEqual(len(result["warnings"]), 1)
        self.assertEqual(result["anomalies"][("R+1", "IfcWall", "IDS : Murs")], {guid(1)})


if __name__ == "__main__":
    unittest.main()
