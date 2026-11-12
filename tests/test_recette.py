"""Tests de la recette, du suivi entre livraisons et des livrables (fiche
Excel, BCF, inventaire d'exploitation).

    python -m unittest discover tests
"""
import json
import os
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import support  # noqa: E402
from support import RULES_XLSX, add_equipment, build_broken_model, build_model, guid, run_recette  # noqa: E402

from openpyxl import load_workbook  # noqa: E402

from checkers_bonsai.core import bcf, inventory, journal, recette, report, spatial  # noqa: E402


def cell_texts(sheet) -> list:
    return [str(c.value) for row in sheet.iter_rows() for c in row if c.value is not None]


class RecetteTest(unittest.TestCase):
    def test_full_recette(self):
        result = run_recette()
        self.assertEqual(result.structure_verdict, recette.VALIDATED)
        self.assertEqual(result.pset_verdict, recette.REJECTED)        # 1 mur conforme sur 4 éléments
        self.assertEqual(result.ids_verdict, recette.REJECTED)         # 1 mur non conforme sur 3
        self.assertEqual(result.verdict, recette.REJECTED)
        self.assertEqual(result.valid_guids, {guid(1)})                 # conforme à l'IDS et aux règles
        self.assertIn(guid(3), result.invalid_guids)
        self.assertEqual(len(result.model["file"]["sha256"]), 64)
        self.assertEqual(result.model["project_guid"], guid(900))
        self.assertEqual({a["source"] for a in result.anomalies}, {"IDS", "PSet"})

    def test_ids_only_colours_conforming_elements(self):
        result = run_recette(rules_path="")
        self.assertEqual(result.valid_guids, {guid(1), guid(2)})        # applicables à l'IDS, sans anomalie
        self.assertEqual(result.invalid_guids, {guid(3)})

    def test_structure_only(self):
        model, _ = build_broken_model()
        result = recette.run_recette(model, recette.RecetteSettings())
        self.assertEqual((result.verdict, result.pset, result.ids_result), (recette.REJECTED, None, None))
        self.assertEqual(result.anomalies[0]["severity"], "Bloquant")   # les constats bloquants d'abord
        self.assertEqual(result.anomalies[-1]["severity"], "Mineur")
        sound = recette.run_recette(build_model(), recette.RecetteSettings())
        self.assertEqual((sound.verdict, sound.anomalies), (recette.VALIDATED, []))

    def test_settings_are_checked(self):
        with self.assertRaises(ValueError):                              # aucun contrôle demandé
            recette.run_recette(build_model(), recette.RecetteSettings(check_structure=False))
        with self.assertRaises(ValueError):                              # seuils inversés
            recette.run_recette(build_model(), recette.RecetteSettings(threshold_ok=0.8, threshold_reserve=0.9))
        with self.assertRaises(ValueError):
            recette.run_recette(None, recette.RecetteSettings())

    def test_thresholds(self):
        settings = recette.RecetteSettings(threshold_ok=1.0, threshold_reserve=0.5)
        self.assertEqual(recette.rate_verdict(0.99, settings), recette.RESERVED)
        self.assertEqual(recette.rate_verdict(1.0, settings), recette.VALIDATED)
        self.assertEqual(recette.rate_verdict(0.49, settings), recette.REJECTED)

    def test_unsaved_changes_are_flagged(self):
        result = recette.run_recette(build_model(), recette.RecetteSettings(), unsaved_changes=True)
        self.assertTrue(any("non enregistrées" in w for w in result.warnings))


class JournalTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.first = run_recette()
        self.journal_path = os.path.join(self.dir, "recette.json")
        journal.write_record(self.first, self.journal_path)

    def second_delivery(self, **settings):
        """Livraison suivante : le mur 3 reçoit son PSet (corrigé), la porte
        est supprimée, le mur 1 perd le sien (régression) ; le mur 2 reste
        en défaut."""
        model = build_model()
        model.by_guid(guid(3))._psets = {"Pset_WallCommon": {"AcousticRating": "Rw 50"}}
        model.entities.remove(model.by_guid(guid(4)))
        model.by_guid(guid(1))._psets = {}
        values = dict(previous_path=self.journal_path, failed_walls=slice(0, 0), model_content="ISO-10303-21; v2")
        values.update(settings)
        return run_recette(model, **values)

    def test_round_trip(self):
        record = journal.load_record(self.journal_path)
        self.assertEqual((record["lot"], record["verdict"]), ("Lot 02 – Gros œuvre", recette.REJECTED))
        self.assertEqual(record["model"]["file"]["sha256"], self.first.model["file"]["sha256"])
        self.assertEqual(len(record["anomalies"]), len(self.first.anomalies))

    def test_invalid_journals_are_refused(self):
        cases = {"texte.json": "pas du JSON", "autre.json": json.dumps({"format": "autre"}),
                 "guids.json": json.dumps({**journal.to_record(self.first), "anomalies": [
                     {"level": "RDC", "ifc_class": "IfcWall", "requirement": "X", "source": "PSet", "guids": [1]}]}),
                 "version.json": json.dumps({**journal.to_record(self.first), "version": 99})}
        for name, content in cases.items():
            path = os.path.join(self.dir, name)
            with open(path, "w", encoding="utf-8") as fh:
                fh.write(content)
            with self.assertRaises(journal.JournalError, msg=name):
                journal.load_record(path)
        with self.assertRaises(journal.JournalError):
            journal.load_record(os.path.join(self.dir, "absent.json"))
        with self.assertRaises(journal.JournalError):                    # arrête la recette avant tout calcul
            run_recette(previous_path=os.path.join(self.dir, "texte.json"))

    def test_evolution_between_deliveries(self):
        e = self.second_delivery().evolution
        self.assertEqual(e["previous_verdict"], recette.REJECTED)
        # Levées : PSet du mur 3, IDS du mur 3, quantité de la porte (supprimée).
        self.assertEqual((e["resolved_count"], e["removed_count"]), (3, 1))
        self.assertEqual((e["new_count"], e["persistent_count"]), (1, 1))  # mur 1 régresse, mur 2 persiste
        removed = [g for g in e["resolved"] if g["removed"]]
        self.assertEqual([(g["ifc_class"], g["removed"]) for g in removed], [("IfcDoor", 1)])
        self.assertEqual(e["new"][0]["guids"], [guid(1)])
        self.assertEqual(e["warnings"], [])

    def test_comparison_warnings(self):
        same = run_recette(previous_path=self.journal_path)               # même fichier, autre lot
        self.assertTrue(any("identique" in w for w in same.evolution["warnings"]))
        # Contrôle des règles Excel non relancé : ses anomalies ne sont pas
        # « levées », elles n'ont simplement pas été recontrôlées.
        other = self.second_delivery(lot="Lot 03", rules_path="")
        warnings = " ".join(other.evolution["warnings"])
        self.assertIn("autre lot", warnings)
        self.assertIn("n'ont pas été recontrôlées (contrôle non relancé : PSet)", warnings)
        self.assertEqual(other.evolution["not_rechecked_count"], 3)
        self.assertTrue(all(g["source"] != "PSet" for g in other.evolution["resolved"]))
        self.assertEqual(other.warnings[-len(other.evolution["warnings"]):], other.evolution["warnings"])
        # Référentiel modifié entre deux recettes : signalé.
        copy = os.path.join(self.dir, "regles_v2.xlsx")
        workbook = load_workbook(RULES_XLSX)
        workbook["Lisez-moi"]["A20"] = "version 2"
        workbook.save(copy)
        changed = self.second_delivery(rules_path=copy)
        self.assertTrue(any("règles Excel a changé" in w for w in changed.evolution["warnings"]))
        # Contrôle ajouté : ses anomalies sont « nouvelles » par construction.
        structure_only = recette.run_recette(build_model(), recette.RecetteSettings(lot="Lot 02 – Gros œuvre"))
        path = os.path.join(self.dir, "structure_seule.json")
        journal.write_record(structure_only, path)
        with_ids = run_recette(previous_path=path)
        self.assertTrue(any("Contrôle(s) ajouté(s)" in w and "IDS" in w for w in with_ids.evolution["warnings"]))
        self.assertEqual(len(with_ids.evolution["journal"]["sha256"]), 64)

    def test_model_level_findings_are_tracked(self):
        model, _ = build_broken_model()
        broken = recette.run_recette(model, recette.RecetteSettings(lot="Lot 02 – Gros œuvre"))
        path = os.path.join(self.dir, "structure.json")
        journal.write_record(broken, path)
        model.length_unit = "METRE"                                       # unité de longueur ajoutée
        fixed = recette.run_recette(model, recette.RecetteSettings(lot="Lot 02 – Gros œuvre", previous_path=path))
        resolved = [(g["requirement"], g["guids"]) for g in fixed.evolution["resolved"]]
        self.assertEqual(resolved, [("Unité de longueur déclarée", [])])


class ReportTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.path = os.path.join(self.dir, "recette.xlsx")

    def test_report(self):
        result = run_recette()
        report.write_report(result, self.path)
        wb = load_workbook(self.path)
        self.assertEqual(wb.sheetnames, ["Fiche de recette", "Structure", "IDS", "Détail PSet"])
        texts = cell_texts(wb["Fiche de recette"])
        self.assertEqual(texts[0], "CONFIDENTIEL")
        for expected in ("VERDICT : REJETÉ", "Lot 02 – Gros œuvre", result.model["file"]["sha256"],
                         "1. Identification", "2. Résultat", "Structure du modèle", "5. Validation", "Non déterminé"):
            self.assertIn(expected, texts)
        for ws in wb.worksheets:
            self.assertEqual(ws.oddHeader.center.text, "CONFIDENTIEL")
        names = [c.value for c in wb["Détail PSet"]["D"]]
        self.assertIn("'=HYPERLINK(\"x\")", names)                       # formule neutralisée
        self.assertTrue(all(c.data_type != "f" for ws in wb.worksheets for row in ws.iter_rows() for c in row))
        self.assertEqual(len(cell_texts(wb["Structure"])) // 4 >= 11, True)

    def test_report_without_marking_and_structure_only(self):
        model, _ = build_broken_model()
        report.write_report(recette.run_recette(model, recette.RecetteSettings()), self.path)
        wb = load_workbook(self.path)
        self.assertEqual(wb.sheetnames, ["Fiche de recette", "Structure"])
        self.assertEqual(wb["Fiche de recette"]["A1"].value, "FICHE DE RECETTE DE MAQUETTE NUMÉRIQUE (IFC)")
        self.assertIn("Structure – Bloquant", cell_texts(wb["Fiche de recette"]))
        statuses = [c.value for c in wb["Structure"]["D"]]
        self.assertEqual((statuses.count("Non conforme"), statuses.count("Conforme")), (10, 1))

    def test_report_with_evolution(self):
        first = run_recette()
        journal_path = os.path.join(self.dir, "recette.json")
        journal.write_record(first, journal_path)
        model = build_model()
        model.by_guid(guid(3))._psets = {"Pset_WallCommon": {"AcousticRating": "Rw 50"}}
        second = run_recette(model, previous_path=journal_path, failed_walls=slice(0, 0), model_content="v2")
        report.write_report(second, self.path)
        wb = load_workbook(self.path)
        self.assertEqual(wb.sheetnames[:2], ["Fiche de recette", "Évolution"])
        texts = cell_texts(wb["Fiche de recette"])
        self.assertTrue(any(t.startswith("3. Évolution depuis la recette du") for t in texts))
        self.assertIn("6. Validation", texts)
        self.assertEqual({c.value for c in wb["Évolution"]["A"]}, {"Statut", "Levée", "Persistante"})


class BcfTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def export(self, result, name="anomalies.bcf", **kwargs):
        path = os.path.join(self.dir, name)
        stats = bcf.write_bcf(path, result.anomalies, author="recette@exemple.fr",
                              project_name=result.model["project"], project_key=result.model["project_guid"],
                              lot=result.settings.lot, phase="DOE", marking=result.settings.marking,
                              referentials=recette.describe_referentials(result), **kwargs)
        with zipfile.ZipFile(path) as zf:
            files = {n: zf.read(n) for n in zf.namelist()}
        return stats, files

    def assert_valid(self, files):
        if not support.BCF_XSD_DIR:
            return
        from lxml import etree
        xsd = {".bcf": "markup.xsd", ".bcfv": "visinfo.xsd", ".version": "version.xsd", ".bcfp": "project.xsd"}
        for name, content in files.items():
            schema = etree.XMLSchema(etree.parse(os.path.join(support.BCF_XSD_DIR, xsd[os.path.splitext(name)[1]])))
            self.assertTrue(schema.validate(etree.fromstring(content)), f"{name}: {schema.error_log}")

    def test_topics(self):
        result = run_recette()
        stats, files = self.export(result)
        self.assertEqual((stats["topics"], stats["closed"]), (len(result.anomalies), 0))
        topics = [ET.fromstring(c).find("Topic") for n, c in files.items() if n.endswith("markup.bcf")]
        self.assertIn("CONFIDENTIEL", [label.text for label in topics[0].findall("Labels")])
        self.assertTrue(topics[0].findtext("Description").startswith("CONFIDENTIEL"))
        self.assertEqual({t.findtext("Stage") for t in topics}, {"DOE"})
        self.assert_valid(files)

    def test_topic_identifiers_are_stable(self):
        """Même anomalie, même identifiant d'une recette à l'autre : l'outil
        du destinataire met le sujet à jour au lieu de le dupliquer."""
        _, first = self.export(run_recette(), "a.bcf")
        _, second = self.export(run_recette(), "b.bcf")
        self.assertEqual(sorted(first), sorted(second))
        _, other_lot = self.export(run_recette(lot="Lot 03"), "c.bcf")
        self.assertFalse({n for n in first if "/" in n} & {n for n in other_lot if "/" in n})

    def test_resolved_anomalies_are_closed(self):
        journal_path = os.path.join(self.dir, "recette.json")
        journal.write_record(run_recette(), journal_path)
        model = build_model()
        model.by_guid(guid(3))._psets = {"Pset_WallCommon": {"AcousticRating": "Rw 50"}}
        model.entities.remove(model.by_guid(guid(4)))
        second = run_recette(model, previous_path=journal_path, failed_walls=slice(0, 0), model_content="v2")
        stats, files = self.export(second, resolved=second.evolution["resolved"])
        self.assertEqual(stats["closed"], 3)             # PSet et IDS du mur 3, quantité de la porte supprimée
        markups = [ET.fromstring(c) for n, c in files.items() if n.endswith("markup.bcf")]
        closed = [m for m in markups if m.find("Topic").get("TopicStatus") == "Closed"]
        self.assertEqual(len(closed), 3)
        comments = " ".join(m.findtext("Comment/Comment") for m in closed)
        self.assertIn("Anomalie levée", comments)
        self.assertIn("1 élément(s) supprimé(s)", comments)
        self.assert_valid(files)

    def test_structure_findings(self):
        model, _ = build_broken_model()
        result = recette.run_recette(model, recette.RecetteSettings())
        stats, files = self.export(result)
        priorities = {ET.fromstring(c).findtext("Topic/Priority") for n, c in files.items() if n.endswith(".bcf")}
        self.assertEqual(priorities, {"Critical", "Major", "Minor"})
        self.assertEqual(stats["skipped_guids"], 0)
        self.assert_valid(files)


class InventoryTest(unittest.TestCase):
    def setUp(self):
        self.model = add_equipment(build_model())
        self.inventory = inventory.collect(self.model, spatial.LevelCache())

    def test_equipment(self):
        rows = {r["guid"]: r for r in self.inventory["equipment"]}
        self.assertEqual(set(rows), {guid(20), guid(21), guid(22)})      # le tuyau n'est pas un équipement suivi
        pump = rows[guid(20)]
        self.assertEqual(
            (pump["tag"], pump["predefined_type"], pump["type_name"], pump["level"], pump["space"], pump["system"]),
            ("PMP-001", "CIRCULATOR", "Pompe double DN50", "RDC", "012 – Local technique", "Chauffage"))
        self.assertEqual((pump["manufacturer"], pump["model"], pump["serial"], pump["warranty_end"],
                          pump["classification"]), ("=Fabricant", "X200", "SN-42", "2028-06-30", "52.10 Pompes"))
        self.assertEqual((rows[guid(21)]["tag"], rows[guid(21)]["space"]), ("", ""))

    def test_spaces_and_fill_rates(self):
        self.assertEqual(self.inventory["spaces"][0]["area"], 18.46)
        rates = {label: (filled, total) for label, filled, total, _ in self.inventory["equipment_rates"]}
        self.assertEqual(rates["Repère"], (2, 3))
        self.assertEqual(rates["N° de série"], (1, 3))
        empty = inventory.collect(build_model("IFC2X3"), spatial.LevelCache())
        self.assertEqual(empty["equipment"], [])
        self.assertIsNone(empty["equipment_rates"][0][3])               # rien à mesurer : pas de taux

    def test_workbook(self):
        path = os.path.join(tempfile.mkdtemp(), "inventaire.xlsx")
        result = recette.run_recette(self.model, recette.RecetteSettings())
        inventory.write_inventory(path, self.inventory, result.model, marking="CONFIDENTIEL", date=result.date)
        wb = load_workbook(path)
        self.assertEqual(wb.sheetnames, ["Synthèse", "Équipements", "Locaux"])
        self.assertEqual(wb["Équipements"].max_row, 4)
        self.assertEqual(wb["Synthèse"]["A1"].value, "CONFIDENTIEL")
        for ws in wb.worksheets:
            self.assertEqual(ws.oddHeader.center.text, "CONFIDENTIEL")
        self.assertIn("'=Fabricant", cell_texts(wb["Équipements"]))      # formule neutralisée
        self.assertTrue(all(c.data_type != "f" for ws in wb.worksheets for row in ws.iter_rows() for c in row))

    def test_unreadable_information_leaves_a_blank(self):
        import ifcopenshell.util.system
        from unittest import mock
        with mock.patch.object(ifcopenshell.util.system, "get_element_systems", side_effect=RuntimeError("boum")):
            rows = inventory.collect_equipment(self.model, spatial.LevelCache())
        self.assertEqual(len(rows), 3)
        self.assertEqual({r["system"] for r in rows}, {""})


if __name__ == "__main__":
    unittest.main()
