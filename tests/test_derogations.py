"""Tests des dérogations (core/derogations.py) et de leur effet sur la
recette, le journal, la fiche, le BCF et le tableau de suivi ; et du passage
des coordonnées IFC à la scène (core/scene.py).

    python -m unittest discover tests
"""
import json
import os
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import support  # noqa: E402,F401
from support import RULES_XLSX, build_model, guid  # noqa: E402
from test_coordination import BEAM, DUCT, add_boxes, frame, other_lot  # noqa: E402

import ifcopenshell  # noqa: E402  (doublure)
from openpyxl import load_workbook  # noqa: E402

from checkers_bonsai.core import (bcf, clashes, derogations, journal, recette, report, scene,  # noqa: E402
                                  tracking)
from checkers_bonsai.core.structure import MAJOR  # noqa: E402


def anomaly(requirement="Pset_WallCommon", source="PSet", guids=(guid(2), guid(3)), level="R+1",
            ifc_class="IfcWall", severity=None, count=None, **extra):
    return {"level": level, "ifc_class": ifc_class, "requirement": requirement, "source": source,
            "severity": severity, "count": len(guids) if count is None else count, "guids": list(guids), **extra}


class ApplyTest(unittest.TestCase):
    def document(self, *entries):
        document = derogations.new_document()
        for entry in entries:
            derogations.add(document, entry)
        return document

    def test_whole_partial_and_unused(self):
        wall = anomaly()
        whole = derogations.make_entry(wall, "Mur existant conservé", author="J. Martin", lot="ARC",
                                       date=datetime(2026, 11, 12, 10, 0))
        self.assertEqual((whole["guids"], whole["lot"], whole["date"]), ([guid(2), guid(3)], "ARC",
                                                                         "2026-11-12T10:00:00"))
        remaining, admitted, unused = derogations.apply([wall], self.document(whole), "ARC")
        self.assertEqual((remaining, len(admitted), unused), ([], 1, []))
        self.assertEqual(admitted[0]["derogation"]["reason"], "Mur existant conservé")

        partial = derogations.make_entry(wall, "Accord MOE", guids=[guid(3)])
        remaining, admitted, _ = derogations.apply([wall], self.document(partial), "ARC")
        self.assertEqual((remaining[0]["guids"], remaining[0]["count"]), ([guid(2)], 1))
        self.assertEqual((admitted[0]["guids"], admitted[0]["count"]), ([guid(3)], 1))
        # Un nouvel élément avec le même défaut n'est pas couvert.
        remaining, admitted, _ = derogations.apply([anomaly(guids=(guid(7),))], self.document(partial), "ARC")
        self.assertEqual((len(remaining), admitted), (1, []))
        # Plus aucune anomalie : la dérogation est sans objet.
        _, _, unused = derogations.apply([], self.document(partial), "ARC")
        self.assertEqual([e["id"] for e in unused], [partial["id"]])

    def test_scope_by_lot_and_model_level(self):
        finding = anomaly(requirement="IDS : Classification", source="IDS", guids=(), level="Modèle",
                          ifc_class="—", count=1)
        local = derogations.make_entry(finding, "Hors marché", lot="ELEC")
        remaining, admitted, unused = derogations.apply([finding], self.document(local), "CVC")
        self.assertEqual((len(remaining), admitted, unused), (1, [], []))     # autre lot : ignorée, pas « sans objet »
        everywhere = derogations.make_entry(finding, "Hors marché", lot="")
        remaining, admitted, _ = derogations.apply([finding], self.document(everywhere), "CVC")
        self.assertEqual((remaining, admitted[0]["count"]), ([], 1))

    def test_clash_entry_survives_tolerance_change(self):
        conflict = anomaly(requirement="Conflit avec « STR » (intersection, 1 cm) : IfcBeam", source="Conflit",
                           guids=(DUCT,), ifc_class="IfcDuctSegment", severity=MAJOR)
        entry = derogations.make_entry(conflict, "Gaine dans sa réservation")
        self.assertEqual(entry["requirement"], "Conflit avec « STR » (intersection) : IfcBeam")
        retuned = dict(conflict, requirement="Conflit avec « STR » (intersection, 5 cm) : IfcBeam")
        remaining, admitted, _ = derogations.apply([retuned], self.document(entry), "")
        self.assertEqual((remaining, len(admitted)), ([], 1))

    def test_partial_admission_splits_conflict_pairs(self):
        a, b = guid(30), guid(31)
        conflict = anomaly(requirement="Conflit avec « STR » (intersection, 1 cm) : IfcBeam", source="Conflit",
                           guids=(a, b), severity=MAJOR, related_guids=[guid(40), guid(41)], lot="STR",
                           point=(1.0, 0.0, 0.0),
                           conflicts=[{"a": a, "b": guid(40), "point": (1.0, 0.0, 0.0)},
                                      {"a": b, "b": guid(41), "point": (2.0, 0.0, 0.0)}])
        entry = derogations.make_entry(conflict, "Réservation prévue", guids=[a])
        remaining, admitted, _ = derogations.apply([conflict], self.document(entry), "")
        self.assertEqual((remaining[0]["related_guids"], remaining[0]["point"]), ([guid(41)], (2.0, 0.0, 0.0)))
        self.assertEqual([c["a"] for c in admitted[0]["conflicts"]], [a])

    def test_file_round_trip_and_validation(self):
        folder = tempfile.mkdtemp()
        path = derogations.default_path(os.path.join(folder, "CVC.ifc"))
        self.assertEqual(os.path.basename(path), "CVC.derogations.json")
        document = derogations.load_or_new(path)                       # absent : document vide
        derogations.add(document, derogations.make_entry(anomaly(), "Motif"))
        derogations.save(document, path)
        self.assertEqual(len(derogations.load(path)["derogations"]), 1)
        self.assertEqual([n for n in os.listdir(folder) if n.endswith(".tmp")], [])
        with self.assertRaises(derogations.DerogationError):
            derogations.make_entry(anomaly(), "   ")                    # motif obligatoire
        entry_id = document["derogations"][0]["id"]
        self.assertEqual(derogations.revoke(document, entry_id)["id"], entry_id)
        with self.assertRaises(derogations.DerogationError):
            derogations.revoke(document, entry_id)

        good = derogations.make_entry(anomaly(), "Motif")
        for broken, message in (({"format": "autre"}, "pas un fichier Checkers"),
                                ({**derogations.new_document(), "derogations": [dict(good, reason="")]}, "vide"),
                                ({**derogations.new_document(), "derogations": [good, good]}, "double"),
                                ({**derogations.new_document(), "derogations": [dict(good, guids="x")]},
                                 "mal formés")):
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(broken, fh)
            with self.assertRaises(derogations.DerogationError) as raised:
                derogations.load(path)
            self.assertIn(message, str(raised.exception))
        with self.assertRaises(derogations.DerogationError):
            derogations.load(os.path.join(folder, "absent.json"))


class RecetteTest(unittest.TestCase):
    """Effet des dérogations sur une recette complète."""

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.model = add_boxes(frame(build_model()))
        self.other, self.lot_path = other_lot(self.dir, r1=3.2)          # R+1 à une autre altitude : C05
        beam = self.other.add("IfcBeam", GlobalId=BEAM, Name="Poutre", _box=((2, -1, 2.5), (2.3, 3, 3.0)))
        ifcopenshell.contain(beam, self.other.by_guid(guid(902)))
        self.path = os.path.join(self.dir, "CVC.derogations.json")

    def recette(self, **values):
        settings = dict(check_structure=False, lot="CVC", lots=[("STR", self.lot_path)],
                        clash_rules=[clashes.ClashRule(other="STR", selector_a="IfcDuctSegment", selector_b="IfcBeam")])
        settings.update(values)
        return recette.run_recette(self.model, recette.RecetteSettings(**settings))

    def admit(self, result, source, reason="Admis en réunion de synthèse", **kwargs):
        document = derogations.load_or_new(self.path)
        target = next(a for a in result.anomalies if a["source"] == source)
        derogations.add(document, derogations.make_entry(target, reason, author="J. Martin", lot="CVC", **kwargs))
        derogations.save(document, self.path)

    def test_verdicts_follow_admissions(self):
        first = self.recette()
        self.assertEqual(first.verdict, recette.RESERVED)
        self.admit(first, "Conflit")
        second = self.recette(derogations_path=self.path)
        self.assertEqual((second.clash_verdict, second.federation_verdict, second.verdict),
                         (recette.VALIDATED, recette.RESERVED, recette.RESERVED))
        self.assertEqual([a["source"] for a in second.admitted], ["Conflit"])
        self.assertNotIn(DUCT, second.invalid_guids)
        self.assertTrue(any("1 anomalie(s) admise(s)" in w for w in second.warnings))
        self.admit(second, "Cohérence")
        third = self.recette(derogations_path=self.path)
        self.assertEqual((third.verdict, len(third.anomalies), third.admitted_elements), (recette.VALIDATED, 0, 2))
        # Les résultats bruts des contrôles ne changent pas : la fiche garde le constat.
        self.assertEqual(third.clashes["total"], 1)

    def test_rates_count_fully_admitted_elements(self):
        values = dict(lots=[], clash_rules=[], rules_path=RULES_XLSX, corps_etat="ARC")
        first = self.recette(**values)
        self.assertEqual(first.pset_rate, 0.25)
        wall = guid(3)
        document = derogations.new_document()
        for item in first.anomalies:
            if item["source"] == "PSet" and wall in item["guids"]:
                derogations.add(document, derogations.make_entry(item, "Mur démoli en phase 2", guids=[wall]))
        derogations.save(document, self.path)
        second = self.recette(derogations_path=self.path, **values)
        self.assertEqual(second.pset_rate, 0.5)
        self.assertEqual(second.pset["valid"], 1)                        # résultat brut inchangé
        self.assertNotIn(wall, second.invalid_guids | second.valid_guids)  # ni rouge ni vert

    def test_ids_and_survey_rates(self):
        """Taux IDS et terrain recalculés sans relancer les contrôles."""
        result = recette.RecetteResult(settings=recette.RecetteSettings(lot="CVC"), date=datetime.now(), model={})
        result.ids_result = {"specs_checked": 2, "specs_failed": 2, "cardinality_failures": 1, "applicable_total": 4,
                             "failed_total": 2, "pass_rate": 0.5, "applicable_guids": set(), "warnings": []}
        result.survey_result = {"checks": 4, "conform": 1, "pass_rate": 0.25, "conform_guids": [], "warnings": []}
        ids_failure = anomaly(requirement="IDS : Repérage", source="IDS", guids=(guid(2), guid(3)))
        cardinality = anomaly(requirement="IDS : Locaux", source="IDS", guids=(), level="Modèle", count=1)
        unmodelled = anomaly(requirement="Terrain : non modélisé", source="Terrain", guids=(), count=2)
        result.raw_anomalies = [ids_failure, cardinality, unmodelled]
        recette.finalize(result)
        self.assertEqual((result.ids_verdict, result.survey_verdict), (recette.REJECTED, recette.REJECTED))
        document = derogations.new_document()
        for item in result.raw_anomalies:
            derogations.add(document, derogations.make_entry(item, "Accord MOA"))
        result.derogation_doc = document
        recette.finalize(result)
        recette.finalize(result)                                          # idempotent
        self.assertEqual((result.ids_rate, result.ids_verdict), (1.0, recette.VALIDATED))
        self.assertEqual((result.survey_rate, result.survey_verdict), (0.75, recette.REJECTED))
        self.assertEqual(len(result.admitted), 3)
        self.assertEqual(sum("admise(s)" in w for w in result.warnings), 1)

    def test_journal_bcf_report_and_tracking(self):
        first = self.recette()
        journal_path = os.path.join(self.dir, "recette_1.json")
        journal.write_record(first, journal_path)
        open_conflict = next(a for a in first.anomalies if a["source"] == "Conflit")
        self.admit(first, "Conflit", reason="Gaine dans sa réservation")
        second = self.recette(derogations_path=self.path, previous_path=journal_path)
        e = second.evolution
        self.assertEqual((e["admitted_count"], e["resolved_count"], e["new_count"]), (1, 0, 0))

        journal_2 = os.path.join(self.dir, "recette_2.json")
        journal.write_record(second, journal_2)
        record = journal.load_record(journal_2)
        self.assertEqual(record["admitted"][0]["reason"], "Gaine dans sa réservation")
        self.assertEqual(record["referentials"]["derogations"]["name"], "CVC.derogations.json")

        # Dérogation retirée : l'anomalie redevient ouverte, persistante et non nouvelle.
        document = derogations.load(self.path)
        derogations.revoke(document, document["derogations"][0]["id"])
        derogations.save(document, self.path)
        third = self.recette(derogations_path=self.path, previous_path=journal_2)
        self.assertEqual((third.evolution["new_count"], third.evolution["persistent_count"]), (0, 2))

        bcf_path = os.path.join(self.dir, "admises.bcf")
        stats = bcf.write_bcf(bcf_path, second.anomalies, project_key="p", lot="CVC", admitted=second.admitted)
        self.assertEqual(stats["closed"], 1)
        closed_guid = bcf.topic_guid("p", "CVC", open_conflict)          # même sujet que lorsqu'il était ouvert
        with zipfile.ZipFile(bcf_path) as zf:
            markup = ET.fromstring(zf.read(f"{closed_guid}/markup.bcf"))
        self.assertEqual(markup.find("Topic").get("TopicStatus"), "Closed")
        self.assertIn("admise par dérogation", markup.findtext("Comment/Comment"))
        self.assertIn("Gaine dans sa réservation", markup.findtext("Comment/Comment"))

        fiche = os.path.join(self.dir, "fiche.xlsx")
        report.write_report(second, fiche)
        workbook = load_workbook(fiche)
        self.assertIn("Dérogations", workbook.sheetnames)
        texts = [str(c.value) for row in workbook["Dérogations"].iter_rows() for c in row if c.value is not None]
        self.assertIn("Gaine dans sa réservation", texts)
        fiche_texts = [str(c.value) for row in workbook["Fiche de recette"].iter_rows() for c in row if c.value]
        self.assertTrue(any(t.startswith("Dérogations (1 élément") for t in fiche_texts))
        self.assertTrue(any(t.startswith("Dérogations (SHA-256)") for t in fiche_texts))

        table = os.path.join(self.dir, "suivi.xlsx")
        records, _ = tracking.collect_journals(self.dir)
        tracking.write_tracking(table, records)
        ws = load_workbook(table)["Synthèse"]
        header = [c.value for c in ws[1]]
        self.assertEqual(ws[2][header.index("Anomalies admises")].value, 1)

    def test_unused_derogation_is_reported(self):
        first = self.recette()
        self.admit(first, "Conflit")
        moved = self.model.by_guid(DUCT)
        moved._box = ((0, 10, 2.6), (6, 10.4, 3.0))                      # gaine déplacée : plus de conflit
        result = self.recette(derogations_path=self.path)
        self.assertEqual(len(result.derogations["unused"]), 1)
        self.assertTrue(any("sans objet" in w for w in result.warnings))
        fiche = os.path.join(self.dir, "fiche.xlsx")
        report.write_report(result, fiche)
        texts = [str(c.value) for row in load_workbook(fiche)["Dérogations"].iter_rows() for c in row if c.value]
        self.assertIn("Dérogations sans objet (1)", texts)

    def test_unreadable_file_stops_the_recette(self):
        with open(self.path, "w", encoding="utf-8") as fh:
            fh.write("pas du JSON")
        with self.assertRaises(derogations.DerogationError):
            self.recette(derogations_path=self.path)


class SceneTest(unittest.TestCase):
    def test_identity_without_offset(self):
        self.assertEqual(scene.transform(scene.world_to_scene(None), (1, 2, 3)), (1.0, 2.0, 3.0))

    def test_offset_and_rotation(self):
        # Scène décalée de (652000, 6862000, 35) et tournée de 90° : le point
        # (652000, 6862001, 36) tombe à (1, 0, 1) dans la scène.
        matrix = scene.world_to_scene(((652000.0, 6862000.0, 35.0), 0.0, 1.0))
        x, y, z = scene.transform(matrix, (652000.0, 6862001.0, 36.0))
        self.assertAlmostEqual(x, 1.0)
        self.assertAlmostEqual(y, 0.0)
        self.assertAlmostEqual(z, 1.0)


if __name__ == "__main__":
    unittest.main()
