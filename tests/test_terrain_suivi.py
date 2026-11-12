"""Tests du relevé terrain, du tableau de suivi, du diagnostic d'installation
et de la robustesse des livrables face à des données inattendues.

    python -m unittest discover tests
"""
import json
import os
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime
from unittest import mock
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import support  # noqa: E402,F401
from support import add_equipment, build_model, fake_document, fake_spec, guid, run_recette  # noqa: E402

from openpyxl import Workbook, load_workbook  # noqa: E402

from checkers_bonsai.core import (bcf, diagnostic, ids as ids_check, inventory, journal, recette, report,  # noqa: E402
                                  rules, spatial, survey, tracking)

PUMP, BARE_PUMP, LIFT = guid(20), guid(21), guid(22)


class SurveyTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.model = add_equipment(build_model())
        self.path = os.path.join(self.dir, "releve.xlsx")
        equipment = inventory.collect_equipment(self.model, spatial.LevelCache())
        survey.write_survey(self.path, equipment, {"project": "Projet", "file": {"name": "maquette.ifc"}},
                            marking="CONFIDENTIEL", date=datetime(2026, 11, 12, 9, 0))

    def fill(self, answers: dict, extra_rows=()):
        """answers : {GlobalId: {intitulé de colonne: valeur}}."""
        workbook = load_workbook(self.path)
        ws = workbook[survey.SHEET]
        header = [c.value for c in ws[1]]
        guid_column = header.index("GlobalId") + 1
        for row in range(2, ws.max_row + 1):
            for label, value in answers.get(ws.cell(row, guid_column).value, {}).items():
                ws.cell(row, header.index(label) + 1, value)
        for extra in extra_rows:
            ws.append([extra.get(label) for label in header])
        workbook.save(self.path)

    def check(self):
        return survey.check_survey(self.model, survey.load_survey(self.path), spatial.LevelCache())

    def test_sheet_to_fill(self):
        workbook = load_workbook(self.path)
        ws = workbook[survey.SHEET]
        self.assertEqual([c.value for c in ws[1]], [label for _, label, _, _ in survey.COLUMNS])
        self.assertEqual(ws.max_row, 4)                                   # 3 équipements, pas le tuyau
        row = next(r for r in ws.iter_rows(min_row=2, values_only=True) if r[2] == "PMP-001")
        self.assertEqual((row[1], row[5], row[10]), ("012 – Local technique", "SN-42", PUMP))
        self.assertEqual(row[6:10], (None, None, None, None))             # colonnes à remplir sur site
        self.assertEqual(ws.oddHeader.center.text, "CONFIDENTIEL")
        self.assertEqual(len(ws.data_validations.dataValidation), 1)      # liste oui / non
        self.assertIn("Mode d'emploi", workbook.sheetnames)

    def test_blank_sheet_means_nothing_verified(self):
        result = self.check()
        self.assertEqual((result["verified"], result["unverified"], result["pass_rate"]), (0, 3, None))
        self.assertEqual(result["anomalies"], {})

    def test_gaps_between_site_and_model(self):
        self.fill({
            PUMP: {"Présent sur site (oui / non)": "Oui", "Repère relevé": "pmp-001",   # casse ignorée : conforme
                   "N° de série relevé": "SN-99"},                                      # maquette : SN-42
            BARE_PUMP: {"Présent sur site (oui / non)": "non", "Observation": "déposée"},
            LIFT: {"Présent sur site (oui / non)": "x", "Repère relevé": "ASC-02"},     # maquette : ASC-01
        }, extra_rows=[{"Désignation": "Extracteur en toiture", "Local": "Toiture", "Repère relevé": "EXT-07"},
                       {"GlobalId": guid(999), "Présent sur site (oui / non)": "oui"}])  # absent de la maquette
        result = self.check()
        self.assertEqual((result["verified"], result["conform"], result["unmodelled"], result["unknown"]),
                         (3, 0, 1, 1))
        self.assertEqual(result["pass_rate"], 0.0)
        self.assertEqual(result["anomalies"], {
            ("RDC", "IfcPump", survey.SERIAL_DIFFERS): {PUMP},
            ("RDC", "IfcPump", survey.ABSENT): {BARE_PUMP},
            ("RDC", "IfcTransportElement", survey.TAG_DIFFERS): {LIFT}})
        self.assertEqual(result["model_anomalies"], {("Modèle", "Équipement", survey.UNMODELLED): 1})
        serial = next(d for d in result["details"] if d["problem"] == survey.SERIAL_DIFFERS)
        self.assertEqual((serial["model"], serial["found"]), ("SN-42", "SN-99"))
        self.assertTrue(any("absent de la maquette" in w for w in result["warnings"]))

    def test_numbers_typed_in_excel_and_unreadable_answers(self):
        self.model.by_guid(BARE_PUMP)._psets = {"Pset_ManufacturerOccurrence": {"SerialNumber": "123"}}
        self.fill({BARE_PUMP: {"Présent sur site (oui / non)": 1, "N° de série relevé": 123.0},  # 123.0 = « 123 »
                   LIFT: {"Présent sur site (oui / non)": "peut-être"}})
        result = self.check()
        self.assertEqual((result["verified"], result["conform"]), (1, 1))
        self.assertEqual(result["conform_guids"], [BARE_PUMP])
        self.assertTrue(any("non comprise" in w for w in result["warnings"]))

    def test_recette_with_survey(self):
        self.fill({PUMP: {"Présent sur site (oui / non)": "oui"},
                   BARE_PUMP: {"Présent sur site (oui / non)": "non"}})
        settings = recette.RecetteSettings(survey_path=self.path, check_structure=False, lot="CVC")
        result = recette.run_recette(self.model, settings)
        self.assertEqual(result.sources, {"Terrain"})
        self.assertEqual((result.survey_rate, result.survey_verdict), (0.5, recette.REJECTED))
        self.assertEqual((result.valid_guids, result.invalid_guids), ({PUMP}, {BARE_PUMP}))
        self.assertEqual([a["source"] for a in result.anomalies], ["Terrain"])
        self.assertTrue(any("1 équipement(s) non vérifié(s)" in w for w in result.warnings))

        path = os.path.join(self.dir, "fiche.xlsx")
        report.write_report(result, path)
        workbook = load_workbook(path)
        self.assertEqual(workbook.sheetnames, ["Fiche de recette", "Terrain"])
        texts = [str(c.value) for row in workbook["Fiche de recette"].iter_rows() for c in row if c.value is not None]
        self.assertIn("Relevé terrain", texts)
        self.assertTrue(any(t.startswith("releve.xlsx – ") for t in texts))          # empreinte du relevé
        record = journal.to_record(result)
        self.assertEqual((record["sources"], record["verdicts"]["survey"]), (["Terrain"], recette.REJECTED))

        bcf_path = os.path.join(self.dir, "ecarts.bcf")
        bcf.write_bcf(bcf_path, result.anomalies, lot="CVC")
        with zipfile.ZipFile(bcf_path) as zf:
            markup = next(ET.fromstring(zf.read(n)) for n in zf.namelist() if n.endswith("markup.bcf"))
        self.assertIn("relevé fait sur site : absent sur site", markup.find("Topic").findtext("Description"))

    def test_invalid_survey_files(self):
        other = os.path.join(self.dir, "autre.xlsx")
        workbook = Workbook()
        workbook.active.title = "Feuille"
        workbook.save(other)
        with self.assertRaises(survey.SurveyFileError):
            survey.load_survey(other)                                     # onglet absent
        workbook = Workbook()
        workbook.active.title = survey.SHEET
        workbook.active.append(["Désignation", "Observation"])
        workbook.save(other)
        with self.assertRaises(survey.SurveyFileError):
            survey.load_survey(other)                                     # colonnes obligatoires absentes
        with open(other, "w", encoding="utf-8") as fh:
            fh.write("pas un classeur")
        with self.assertRaises(rules.RulesFileError) as raised:
            survey.load_survey(other)
        self.assertIn("Relevé terrain", str(raised.exception))
        with self.assertRaises(ValueError):                               # arrête la recette avant tout calcul
            recette.run_recette(self.model, recette.RecetteSettings(survey_path=other))


class TrackingTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        os.makedirs(os.path.join(self.dir, "lot02"))
        first = run_recette()
        first.date = datetime(2026, 9, 1, 10, 0)
        journal.write_record(first, os.path.join(self.dir, "lot02", "recette_1.json"))
        second = run_recette(failed_walls=slice(0, 0))                     # IDS corrigé : moins d'anomalies
        second.date = datetime(2026, 10, 1, 10, 0)
        journal.write_record(second, os.path.join(self.dir, "lot02", "recette_2.json"))
        other = recette.run_recette(build_model(), recette.RecetteSettings(lot="Lot 05 – CVC", phase="EXE"))
        other.date = datetime(2026, 10, 15, 10, 0)
        journal.write_record(other, os.path.join(self.dir, "lot05.json"))
        self.counts = (tracking.open_anomalies(journal.to_record(first)),
                       tracking.open_anomalies(journal.to_record(second)))
        with open(os.path.join(self.dir, "autre.json"), "w", encoding="utf-8") as fh:
            json.dump({"sans": "rapport"}, fh)                              # ignoré en silence
        broken = {**journal.to_record(first), "version": 99}
        with open(os.path.join(self.dir, "ancien.json"), "w", encoding="utf-8") as fh:
            json.dump(broken, fh)                                           # journal Checkers refusé : signalé

    def test_collect_and_summarize(self):
        records, rejected = tracking.collect_journals(self.dir)
        self.assertEqual([r["_file"] for r in records],
                         [os.path.join("lot02", "recette_1.json"), os.path.join("lot02", "recette_2.json"),
                          "lot05.json"])
        self.assertEqual([name for name, _ in rejected], ["ancien.json"])
        rows = tracking.summarize(records)
        self.assertEqual([(r["lot"], r["count"], r["verdict"]) for r in rows],
                         [("Lot 02 – Gros œuvre", 2, recette.REJECTED), ("Lot 05 – CVC", 1, recette.VALIDATED)])
        self.assertEqual(rows[0]["open"], self.counts[1])
        self.assertEqual(rows[0]["delta"], self.counts[1] - self.counts[0])
        self.assertLess(rows[0]["delta"], 0)                               # anomalies en baisse
        self.assertIsNone(rows[1]["delta"])                                # première recette du lot

    def test_workbook(self):
        records, rejected = tracking.collect_journals(self.dir)
        path = os.path.join(self.dir, "suivi.xlsx")
        tracking.write_tracking(path, records, rejected, marking="CONFIDENTIEL", date=datetime(2026, 11, 12))
        workbook = load_workbook(path)
        self.assertEqual(workbook.sheetnames, ["Synthèse", "Recettes"])
        summary = list(workbook["Synthèse"].iter_rows(values_only=True))
        self.assertEqual(summary[1][1], "Lot 02 – Gros œuvre")
        self.assertEqual(summary[1][4], recette.REJECTED)
        self.assertLess(summary[1][6], 0)                                  # écart négatif : anomalies en baisse
        self.assertTrue(any(row[0] and "Journal refusé : ancien.json" in str(row[0]) for row in summary))
        self.assertEqual(workbook["Recettes"].max_row, 4)
        self.assertEqual(workbook["Synthèse"].oddHeader.center.text, "CONFIDENTIEL")

    def test_missing_folder(self):
        with self.assertRaises(ValueError):
            tracking.collect_journals(os.path.join(self.dir, "absent"))


class DiagnosticTest(unittest.TestCase):
    def test_everything_in_order(self):
        lines = diagnostic.run(build_model(), bonsai_found=True)
        self.assertTrue(all(line["ok"] for line in lines), lines)
        self.assertIn("✓ Modèle IFC : schéma IFC4, 5 élément(s)", diagnostic.to_text(lines))

    def test_missing_library_and_function_are_explained(self):
        required = (("module_absent", ("fonction",)), ("ifcopenshell.util.system", ("fonction_absente",)))
        with mock.patch.object(diagnostic, "REQUIRED", required):
            lines = {line["label"]: line for line in diagnostic.run(None, bonsai_found=False)}
        self.assertFalse(lines["Bonsai"]["ok"])
        self.assertIn("introuvable", lines["module_absent"]["detail"])
        self.assertIn("fonction_absente", lines["ifcopenshell.util.system"]["detail"])
        self.assertFalse(lines["Modèle IFC"]["ok"])


class RobustnessTest(unittest.TestCase):
    """Données inattendues dans la maquette ou les référentiels : les
    livrables doivent rester lisibles et l'erreur, explicite."""

    def test_control_characters_do_not_break_the_deliverables(self):
        model = build_model()
        model.by_guid(guid(3)).Name = "Mur\x0bavec\x00caractères de contrôle"
        result = run_recette(model, marking="R&D")
        folder = tempfile.mkdtemp()
        path = os.path.join(folder, "fiche.xlsx")
        report.write_report(result, path)                                  # openpyxl refuserait ces caractères
        workbook = load_workbook(path)
        names = [c.value for c in workbook["Détail PSet"]["D"]]
        self.assertIn("Muraveccaractères de contrôle", names)
        # « & » introduit un code de mise en forme dans un en-tête Excel (« &D » = date).
        self.assertEqual(workbook["Fiche de recette"].oddHeader.center.text, "R&&D")
        bcf_path = os.path.join(folder, "anomalies.bcf")
        bcf.write_bcf(bcf_path, [{**result.anomalies[0], "requirement": "Exigence\x08 illisible"}])
        with zipfile.ZipFile(bcf_path) as zf:
            for name in zf.namelist():
                ET.fromstring(zf.read(name))                               # XML bien formé

    def test_required_ids_specification_without_element_is_a_visible_anomaly(self):
        model = build_model()
        document = fake_document(fake_spec("Équipements – identification", [], [], status=False))
        settings = recette.RecetteSettings(ids_path="exigences.ids", check_structure=False, lot="CVC")
        with mock.patch.object(ids_check, "load_ids", return_value=document):
            result = recette.run_recette(model, settings)
        self.assertEqual(result.verdict, recette.REJECTED)
        self.assertEqual([(a["level"], a["ifc_class"], a["count"], a["guids"]) for a in result.anomalies],
                         [("Modèle", "Aucun élément", 1, [])])
        description = bcf.topic_description(result.anomalies[0])
        self.assertIn("Aucun élément de la maquette ne correspond", description)
        self.assertIn("maquette entière", description)

    def test_rules_file_mistakes_are_explained(self):
        folder = tempfile.mkdtemp()

        def workbook_with(types_header, types_rows, pset_rows):
            workbook = Workbook()
            workbook.active.title = "Element_Types"
            workbook.active.append(types_header)
            for row in types_rows:
                workbook.active.append(row)
            ws = workbook.create_sheet("Pset_WallCommon")
            ws.append(["IFC_Class", "Parametre", "Type"])
            for row in pset_rows:
                ws.append(row)
            path = os.path.join(folder, f"regles_{len(os.listdir(folder))}.xlsx")
            workbook.save(path)
            return path

        with self.assertRaises(rules.RulesFileError):                      # colonne IFC_Class mal nommée
            rules.load_rules(workbook_with(["Classe"], [["IfcWall"]], []))
        # Règle sur une classe non déclarée : elle ne serait jamais appliquée.
        loaded = rules.load_rules(workbook_with(["IFC_Class"], [["IfcWall"]],
                                                [["IfcWall", "LoadBearing", "Bool"], ["IfcSlab", "LoadBearing", "Bool"]]))
        self.assertEqual(len(loaded["warnings"]), 1)
        self.assertIn("IfcSlab", loaded["warnings"][0])

    def test_malformed_journal_blocks_are_refused(self):
        folder = tempfile.mkdtemp()
        record = journal.to_record(run_recette())
        cases = {"file": {**record, "model": {**record["model"], "file": "maquette.ifc"}},
                 "referentials": {**record, "referentials": {"ids": "exigences.ids"}},
                 "sources": {**record, "sources": "IDS"}}
        for name, content in cases.items():
            path = os.path.join(folder, name + ".json")
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(content, fh)
            with self.assertRaises(journal.JournalError, msg=name):
                journal.load_record(path)


if __name__ == "__main__":
    unittest.main()
