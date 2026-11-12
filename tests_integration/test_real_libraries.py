"""Tests d'intégration avec les vraies bibliothèques IfcOpenShell et
ifctester (celles que Bonsai fournit), sans Blender.

    pip install ifcopenshell ifctester ifcclash openpyxl
    python -m unittest discover tests_integration

Les tests de tests/ utilisent des doublures : ils vérifient la logique de
Checkers, pas son accord avec les bibliothèques réelles. Ceux-ci construisent
une petite maquette avec IfcOpenShell et déroulent une recette complète :
contrôles de structure, IDS d'exemple validés par ifctester, règles Excel,
relevé terrain, fiche, BCF, journal, inventaire. Ils sont sautés si
IfcOpenShell n'est pas installé.
"""
import os
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

try:
    import ifcopenshell
    import ifcopenshell.api.aggregate
    import ifcopenshell.api.owner
    import ifcopenshell.api.project
    import ifcopenshell.api.pset
    import ifcopenshell.api.root
    import ifcopenshell.api.spatial
    import ifcopenshell.api.unit
    import ifctester  # noqa: F401
    HAS_LIBRARIES = True
except ImportError:
    HAS_LIBRARIES = False

EXAMPLES = os.path.join(ROOT, "checkers_bonsai", "examples")


def create(model, ifc_class, name=None, **attributes):
    entity = ifcopenshell.api.root.create_entity(model, ifc_class=ifc_class, name=name)
    for key, value in attributes.items():
        setattr(entity, key, value)
    return entity


def add_pset(model, product, name, properties):
    pset = ifcopenshell.api.pset.add_pset(model, product=product, name=name)
    ifcopenshell.api.pset.edit_pset(model, pset=pset, properties=properties)


def build_model(schema="IFC4"):
    """Projet, site, bâtiment, un étage, un local ; un mur conforme, un mur
    sans propriétés, une porte, une pompe repérée et une pompe nue."""
    model = ifcopenshell.api.project.create_file(schema)
    if schema == "IFC2X3":
        # L'historique de propriété est obligatoire en IFC2X3 : IfcOpenShell
        # exige qu'un auteur et une application existent dans le fichier.
        ifcopenshell.api.owner.add_application(model)
        ifcopenshell.api.owner.add_person_and_organisation(
            model, person=ifcopenshell.api.owner.add_person(model),
            organisation=ifcopenshell.api.owner.add_organisation(model))
    project = create(model, "IfcProject", "Projet d'intégration")
    ifcopenshell.api.unit.assign_unit(model)
    site = create(model, "IfcSite", "Site")
    building = create(model, "IfcBuilding", "Bâtiment A")
    storey = create(model, "IfcBuildingStorey", "RDC")
    space = create(model, "IfcSpace", "012", LongName="Local technique")
    ifcopenshell.api.aggregate.assign_object(model, products=[site], relating_object=project)
    ifcopenshell.api.aggregate.assign_object(model, products=[building], relating_object=site)
    ifcopenshell.api.aggregate.assign_object(model, products=[storey], relating_object=building)
    ifcopenshell.api.aggregate.assign_object(model, products=[space], relating_object=storey)

    wall_class = "IfcWallStandardCase" if schema == "IFC2X3" else "IfcWall"
    wall_ok = create(model, wall_class, "Mur conforme")
    add_pset(model, wall_ok, "Pset_WallCommon", {"LoadBearing": True, "IsExternal": False, "AcousticRating": "Rw 45"})
    wall_bare = create(model, wall_class, "Mur sans propriétés")
    door = create(model, "IfcDoor", "Porte")
    pump_class = "IfcFlowMovingDevice" if schema == "IFC2X3" else "IfcPump"
    pump = create(model, pump_class, "Pompe de circulation", Tag="PMP-001")
    add_pset(model, pump, "Pset_ManufacturerTypeInformation", {"Manufacturer": "Fabricant", "ModelReference": "X200"})
    add_pset(model, pump, "Pset_ManufacturerOccurrence", {"SerialNumber": "SN-42"})
    pump_bare = create(model, pump_class, "Pompe de relevage")
    ifcopenshell.api.spatial.assign_container(model, products=[wall_ok, wall_bare, door, pump_bare],
                                              relating_structure=storey)
    ifcopenshell.api.spatial.assign_container(model, products=[pump], relating_structure=space)
    return model, {"wall_ok": wall_ok, "wall_bare": wall_bare, "door": door, "pump": pump, "pump_bare": pump_bare,
                   "storey": storey, "space": space}


@unittest.skipUnless(HAS_LIBRARIES, "IfcOpenShell et ifctester ne sont pas installés")
class RealLibrariesTest(unittest.TestCase):
    def setUp(self):
        from checkers_bonsai.core import recette
        self.recette = recette
        self.dir = tempfile.mkdtemp()
        self.model, self.parts = build_model()
        self.model_path = os.path.join(self.dir, "maquette.ifc")
        self.model.write(self.model_path)

    def settings(self, **values):
        defaults = dict(ids_path=os.path.join(EXAMPLES, "conception_minimale.ids"),
                        rules_path=os.path.join(EXAMPLES, "regles_exemple.xlsx"), corps_etat="ARC",
                        lot="Lot 02", phase="DOE", tester="Intégration")
        defaults.update(values)
        return self.recette.RecetteSettings(**defaults)

    def test_diagnostic(self):
        from checkers_bonsai.core import diagnostic
        lines = diagnostic.run(self.model)
        self.assertTrue(all(line["ok"] for line in lines), diagnostic.to_text(lines))

    def test_example_ids_pass_the_official_schema(self):
        from checkers_bonsai.core import ids as ids_check
        for name in ("conception_minimale.ids", "exploitation_maintenance.ids"):
            document = ids_check.load_ids(os.path.join(EXAMPLES, name))
            self.assertTrue(ids_check.ids_title(document).startswith("Exemple"), name)

    def test_full_recette_and_deliverables(self):
        from checkers_bonsai.core import bcf, journal, report
        result = self.recette.run_recette(self.model, self.settings(), model_path=self.model_path,
                                          unsaved_changes=False)
        parts = self.parts

        # Structure : la maquette est saine, hors constats mineurs (pas de
        # géométrie ni de géoréférencement dans ce modèle de test).
        statuses = {c["code"]: c["status"] for c in result.structure["checks"]}
        self.assertNotIn("erreur", statuses.values(), result.structure["checks"])
        for code in ("S01", "S02", "S03", "S04", "S05", "S06"):
            self.assertEqual(statuses[code], "ok", code)
        self.assertEqual(result.structure_verdict, self.recette.VALIDATED)

        # IDS, par le vrai ifctester : le mur sans propriétés échoue, pas l'autre.
        ids_result = result.ids_result
        self.assertGreater(ids_result["specs_checked"], 0)
        failed = set().union(*(a["guids"] for a in result.anomalies if a["source"] == "IDS"))
        self.assertIn(parts["wall_bare"].GlobalId, failed)
        self.assertNotIn(parts["wall_ok"].GlobalId, failed)
        self.assertIn(parts["door"].GlobalId, failed)                       # ni degré coupe-feu, ni IsExternal
        self.assertNotIn(parts["space"].GlobalId, failed)                   # numéro et désignation renseignés
        self.assertTrue(ids_result["details"], "détail des échecs IDS vide")
        self.assertTrue(all(d["reason"] and d["requirement"] for d in ids_result["details"]))
        self.assertEqual(result.warnings, [])

        # Règles Excel (corps d'état ARC) : même constat sur les murs.
        pset_failed = set().union(*(a["guids"] for a in result.anomalies if a["source"] == "PSet"))
        self.assertIn(parts["wall_bare"].GlobalId, pset_failed)
        self.assertNotIn(parts["wall_ok"].GlobalId, pset_failed)
        # Coloration : sans les constats de structure (ce modèle de test n'a
        # pas de géométrie), le mur conforme est vert, l'autre rouge.
        colours = self.recette.run_recette(self.model, self.settings(check_structure=False))
        self.assertIn(parts["wall_ok"].GlobalId, colours.valid_guids)
        self.assertIn(parts["wall_bare"].GlobalId, colours.invalid_guids)
        levels = {a["level"] for a in result.anomalies if a["source"] != "Structure"}
        self.assertEqual(levels, {"RDC"})
        self.assertEqual(result.verdict, self.recette.REJECTED)

        # Traçabilité.
        self.assertEqual(result.model["schema"], "IFC4")
        self.assertEqual(result.model["project"], "Projet d'intégration")
        self.assertEqual(len(result.model["file"]["sha256"]), 64)

        # Livrables.
        from openpyxl import load_workbook
        fiche = os.path.join(self.dir, "fiche.xlsx")
        report.write_report(result, fiche)
        self.assertEqual(load_workbook(fiche).sheetnames, ["Fiche de recette", "Structure", "IDS", "Détail PSet"])
        bcf_path = os.path.join(self.dir, "anomalies.bcf")
        stats = bcf.write_bcf(bcf_path, result.anomalies, project_key=result.model["project_guid"], lot="Lot 02")
        self.assertEqual(stats["skipped_guids"], 0)                          # vrais GlobalId : tous valides
        with zipfile.ZipFile(bcf_path) as zf:
            for name in zf.namelist():
                ET.fromstring(zf.read(name))
        journal_path = os.path.join(self.dir, "recette.json")
        journal.write_record(result, journal_path)

        # Livraison suivante : le mur est complété, ses anomalies sont levées.
        add_pset(self.model, parts["wall_bare"], "Pset_WallCommon",
                 {"LoadBearing": False, "IsExternal": True, "AcousticRating": "Rw 40"})
        second = self.recette.run_recette(self.model, self.settings(previous_path=journal_path),
                                          model_path=self.model_path)
        resolved = set().union(*(g["guids"] for g in second.evolution["resolved"]))
        self.assertIn(parts["wall_bare"].GlobalId, resolved)
        self.assertEqual(second.evolution["new_count"], 0)

    def test_inventory_and_field_survey(self):
        from openpyxl import load_workbook
        from checkers_bonsai.core import inventory, spatial, survey
        level_of = spatial.LevelCache()
        data = inventory.collect(self.model, level_of)
        pump = next(row for row in data["equipment"] if row["tag"] == "PMP-001")
        self.assertEqual((pump["manufacturer"], pump["model"], pump["serial"]), ("Fabricant", "X200", "SN-42"))
        self.assertEqual((pump["level"], pump["space"]), ("RDC", "012 – Local technique"))
        self.assertEqual(len(data["equipment"]), 2)
        self.assertEqual(data["spaces"][0]["number"], "012")
        inventory.write_inventory(os.path.join(self.dir, "inventaire.xlsx"), data, {"project": "Projet"})

        path = os.path.join(self.dir, "releve.xlsx")
        survey.write_survey(path, data["equipment"], {"project": "Projet"})
        workbook = load_workbook(path)
        ws = workbook[survey.SHEET]
        header = [c.value for c in ws[1]]
        for row in range(2, ws.max_row + 1):
            tagged = ws.cell(row, header.index("Repère (maquette)") + 1).value == "PMP-001"
            ws.cell(row, header.index("Présent sur site (oui / non)") + 1, "oui" if tagged else "non")
            if tagged:
                ws.cell(row, header.index("N° de série relevé") + 1, "SN-99")
        workbook.save(path)
        result = self.recette.run_recette(self.model, self.recette.RecetteSettings(
            survey_path=path, check_structure=False, lot="CVC"))
        found = {a["requirement"]: a["guids"] for a in result.anomalies}
        self.assertEqual(found, {survey.SERIAL_DIFFERS: [self.parts["pump"].GlobalId],
                                 survey.ABSENT: [self.parts["pump_bare"].GlobalId]})

    def test_exploitation_ids_on_equipment(self):
        """IDS d'exemple « exploitation » : la pompe repérée passe
        l'identification (EXP-01), la pompe nue non."""
        settings = self.recette.RecetteSettings(
            ids_path=os.path.join(EXAMPLES, "exploitation_maintenance.ids"), check_structure=False)
        result = self.recette.run_recette(self.model, settings)
        by_spec = {s["identifier"]: s for s in result.ids_result["specifications"]}
        self.assertEqual(by_spec["EXP-06"]["status"], "ignoree")             # spécification IFC2X3
        self.assertEqual((by_spec["EXP-01"]["applicable"], by_spec["EXP-01"]["failed"]), (2, 1))
        self.assertEqual((by_spec["EXP-02"]["applicable"], by_spec["EXP-02"]["failed"]), (2, 1))
        self.assertEqual(by_spec["EXP-07"]["status"], "conforme")            # local identifié

    def test_ifc2x3_model(self):
        model, parts = build_model("IFC2X3")
        result = self.recette.run_recette(model, self.settings())
        self.assertNotIn("erreur", [c["status"] for c in result.structure["checks"]])
        self.assertEqual(result.model["schema"], "IFC2X3")
        # La règle Excel « IfcWall » couvre IfcWallStandardCase.
        self.assertEqual(result.pset["checked"], 3)                          # 2 murs + 1 porte
        self.assertIn(parts["wall_ok"].GlobalId, result.pset["valid_guids"])
        self.assertGreater(result.ids_result["specs_checked"], 0)


try:
    import ifcopenshell.api.context
    import ifcopenshell.api.geometry
    import numpy
    from ifcclash import ifcclash  # noqa: F401
    HAS_IFCCLASH = HAS_LIBRARIES
except ImportError:
    HAS_IFCCLASH = False


def build_lot(storey_z=0.0, building_y=0.0, eastings=652000.0, project_guid=None):
    """Maquette d'un lot avec une géométrie réelle : RDC, R+1 à 3 m,
    contexte Body, géoréférencement Lambert 93."""
    model = ifcopenshell.api.project.create_file("IFC4")
    project = create(model, "IfcProject", "Projet de coordination")
    if project_guid:
        project.GlobalId = project_guid
    ifcopenshell.api.unit.assign_unit(model)
    model3d = ifcopenshell.api.context.add_context(model, context_type="Model")
    body = ifcopenshell.api.context.add_context(model, context_type="Model", context_identifier="Body",
                                                target_view="MODEL_VIEW", parent=model3d)
    crs = model.createIfcProjectedCRS(Name="EPSG:2154")
    model.createIfcMapConversion(SourceCRS=model3d, TargetCRS=crs, Eastings=eastings, Northings=6862000.0,
                                 OrthogonalHeight=35.0, XAxisAbscissa=1.0, XAxisOrdinate=0.0, Scale=1.0)
    site = create(model, "IfcSite", "Site")
    building = create(model, "IfcBuilding", "Bâtiment A")
    rdc = create(model, "IfcBuildingStorey", "RDC")
    r1 = create(model, "IfcBuildingStorey", "R+1")
    ifcopenshell.api.aggregate.assign_object(model, products=[site], relating_object=project)
    ifcopenshell.api.aggregate.assign_object(model, products=[building], relating_object=site)
    ifcopenshell.api.aggregate.assign_object(model, products=[rdc, r1], relating_object=building)
    for product, (x, y, z) in ((site, (0, 0, 0)), (building, (0, building_y, 0)), (rdc, (0, building_y, 0)),
                               (r1, (0, building_y, 3.0 + storey_z))):
        place(model, product, x, y, z)
    return model, body, rdc


def place(model, product, x, y, z):
    matrix = numpy.eye(4)
    matrix[:3, 3] = (x, y, z)
    ifcopenshell.api.geometry.edit_object_placement(model, product=product, matrix=matrix)


def box(model, body, storey, ifc_class, name, origin, size):
    """Élément parallélépipédique : origine (x, y, z) et dimensions (dx, dy, dz)
    en mètres."""
    element = create(model, ifc_class, name)
    ifcopenshell.api.spatial.assign_container(model, products=[element], relating_structure=storey)
    place(model, element, *origin)
    representation = ifcopenshell.api.geometry.add_wall_representation(
        model, context=body, length=size[0], thickness=size[1], height=size[2])
    ifcopenshell.api.geometry.assign_representation(model, product=element, representation=representation)
    return element


@unittest.skipUnless(HAS_IFCCLASH, "IfcOpenShell ou IfcClash ne sont pas installés")
class CoordinationTest(unittest.TestCase):
    """Cohérence entre lots et conflits géométriques avec le vrai IfcClash :
    une gaine CVC pénètre dans une poutre du lot structure."""

    def setUp(self):
        from checkers_bonsai.core import clashes, federation, recette, spatial
        self.clashes, self.federation, self.recette, self.spatial = clashes, federation, recette, spatial
        self.dir = tempfile.mkdtemp()
        self.cvc, body, rdc = build_lot()
        project_guid = self.cvc.by_type("IfcProject")[0].GlobalId
        # La gaine soufflage s'arrête au milieu de la poutre : son extrémité y pénètre de 15 cm, sans face
        # commune. La gaine reprise passe à côté.
        self.duct = box(self.cvc, body, rdc, "IfcDuctSegment", "Gaine soufflage", (0, 0, 2.6), (2.15, 0.4, 0.3))
        self.far = box(self.cvc, body, rdc, "IfcDuctSegment", "Gaine reprise", (0, 5, 2.6), (2.15, 0.4, 0.3))
        structure, body, rdc = build_lot(project_guid=project_guid)
        self.beam = box(structure, body, rdc, "IfcBeam", "Poutre", (2, -1, 2.5), (0.3, 4, 0.5))
        self.str_path = os.path.join(self.dir, "STR.ifc")
        structure.write(self.str_path)
        shifted, body, rdc = build_lot(storey_z=0.2, building_y=0.25, eastings=652000.5, project_guid=project_guid)
        box(shifted, body, rdc, "IfcPipeSegment", "Canalisation", (10, 10, 1), (2, 0.1, 0.1))
        self.plb_path = os.path.join(self.dir, "PLB.ifc")
        shifted.write(self.plb_path)

    def test_federation(self):
        lots = self.federation.open_lots([("STR", self.str_path), ("PLB", self.plb_path)])
        result = self.federation.check_federation(self.cvc, lots, self.spatial.LevelCache())
        status = {(c["lot"], c["code"]): c["status"] for c in result["checks"]}
        for code in ("C01", "C02", "C03", "C04", "C05", "C06", "C07", "C08", "C09"):
            self.assertEqual(status[("STR", code)], "ok", (code, result["checks"]))
        for code in ("C04", "C05", "C07"):
            self.assertEqual(status[("PLB", code)], "ko", (code, result["checks"]))
        messages = {c["code"]: c["message"] for c in result["checks"] if c["lot"] == "PLB"}
        self.assertIn("0.250 m", messages["C04"])
        self.assertIn("R+1", messages["C05"])
        self.assertFalse(result["frame_ok"])

    def test_clashes(self):
        lots = self.federation.open_lots([("STR", self.str_path)])
        rules = [self.clashes.ClashRule(other="STR", selector_a="IfcDuctSegment", selector_b="IfcBeam"),
                 self.clashes.ClashRule(selector_a="IfcDuctSegment", selector_b="IfcDuctSegment",
                                        mode="clearance", distance=5.0)]
        result = self.clashes.run_clashes(self.cvc, lots, rules, self.spatial.LevelCache())
        self.assertEqual([r["status"] for r in result["rules"]], ["ko", "ko"], result["rules"])
        external = [d for d in result["details"] if d["b_guid"] == self.beam.GlobalId]
        self.assertEqual([d["a_guid"] for d in external], [self.duct.GlobalId])     # la gaine éloignée passe
        internal = [d for d in result["details"] if d["b_guid"] != self.beam.GlobalId]
        self.assertEqual(len(internal), 1)                                           # paire comptée une fois
        point = next(a["point"] for a in result["anomalies"].values() if a["related_guids"])
        self.assertTrue(1.9 <= point[0] <= 2.4 and 2.4 <= point[2] <= 3.1, point)

    def test_other_lot_geometry_for_the_3d_view(self):
        """Géométrie d'un élément d'un autre lot, telle que la vue 3D
        l'affiche : en mètres, placement compris."""
        from checkers_bonsai.core import scene
        structure = ifcopenshell.open(self.str_path)
        vertices, triangles = scene.element_mesh(structure.by_guid(self.beam.GlobalId))
        self.assertGreaterEqual(len(triangles), 12)
        for axis, (low, high) in enumerate(((2.0, 2.3), (-1.0, 3.0), (2.5, 3.0))):
            values = [v[axis] for v in vertices]
            self.assertAlmostEqual(min(values), low, places=4)
            self.assertAlmostEqual(max(values), high, places=4)

    def test_derogation_for_a_clash(self):
        from checkers_bonsai.core import derogations
        settings = self.recette.RecetteSettings(
            check_structure=False, lot="CVC", lots=[("STR", self.str_path)],
            clash_rules=[self.clashes.ClashRule(other="STR", selector_a="IfcDuctSegment", selector_b="IfcBeam")])
        first = self.recette.run_recette(self.cvc, settings)
        conflict = next(a for a in first.anomalies if a["source"] == "Conflit")
        path = os.path.join(self.dir, "CVC.derogations.json")
        document = derogations.new_document()
        derogations.add(document, derogations.make_entry(conflict, "Réservation prévue", lot="CVC"))
        derogations.save(document, path)
        settings.derogations_path = path
        second = self.recette.run_recette(self.cvc, settings)
        self.assertEqual((second.clash_verdict, len(second.admitted)), (self.recette.VALIDATED, 1))

    def test_recette(self):
        settings = self.recette.RecetteSettings(
            check_structure=False, lot="CVC", lots=[("STR", self.str_path)],
            clash_rules=[self.clashes.ClashRule(other="STR", selector_a="IfcDuctSegment", selector_b="IfcBeam")])
        result = self.recette.run_recette(self.cvc, settings)
        self.assertEqual(result.federation_verdict, self.recette.VALIDATED)
        self.assertEqual(result.clash_verdict, self.recette.RESERVED)
        self.assertEqual(result.invalid_guids, {self.duct.GlobalId})
        from checkers_bonsai.core import bcf, report
        report.write_report(result, os.path.join(self.dir, "fiche.xlsx"))
        stats = bcf.write_bcf(os.path.join(self.dir, "conflits.bcf"), result.anomalies, project_key="p", lot="CVC")
        self.assertEqual(stats["components"], 2)


if __name__ == "__main__":
    unittest.main()
