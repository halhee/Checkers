"""Tests de l'export BCF (checkers_core/bcf.py).

    python -m unittest discover tests

Validation contre les schémas officiels buildingSMART si BCF_XSD_DIR pointe
vers le dossier Schemas/ du dépôt https://github.com/buildingSMART/BCF-XML
(branche release_2_1) et que lxml est installé ; sinon seule la structure
est vérifiée.
"""
import os
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from checkers_core.bcf import write_bcf  # noqa: E402

GUID_A = "2O2Fr$t4X7Zf8NOew3FLOH"
GUID_B = "1hOSvn6df7F8_7GcBWlRGQ"
ANOMALIES = [
    {"lot": "Architecture", "floor": "R+1", "ifc_class": "IfcWall", "pset": "Pset_WallCommon",
     "count": 2, "guids": [GUID_A, GUID_B]},
    {"lot": "Structure", "floor": "Sans étage", "ifc_class": "IfcDoor",
     "pset": "IDS : Portes -- résistance au feu", "count": 1, "guids": [GUID_A, "pas-un-guid"]},
    {"lot": "Structure", "floor": "R+2", "ifc_class": "IfcSlab", "pset": "BaseQuantities", "count": 3},
]

XSD_DIR = os.environ.get("BCF_XSD_DIR")


class BcfExportTest(unittest.TestCase):
    def setUp(self):
        self.path = os.path.join(tempfile.mkdtemp(), "anomalies.bcf")
        self.stats = write_bcf(self.path, ANOMALIES, author="recette@example.org",
                               project_name="Projet de démonstration — bâtiment A", rules_description="IDS « exemple.ids »")
        self.zip = zipfile.ZipFile(self.path)

    def tearDown(self):
        self.zip.close()

    def test_one_topic_per_anomaly(self):
        markups = [n for n in self.zip.namelist() if n.endswith("/markup.bcf")]
        self.assertEqual(len(markups), 3)
        self.assertEqual(self.stats, {"topics": 3, "components": 3, "skipped_guids": 1})
        self.assertIn("bcf.version", self.zip.namelist())

    def test_topic_content(self):
        titles, labels = set(), set()
        for name in self.zip.namelist():
            if name.endswith("/markup.bcf"):
                topic = ET.fromstring(self.zip.read(name)).find("Topic")
                self.assertEqual(name.split("/")[0], topic.get("Guid"))
                titles.add(topic.findtext("Title"))
                labels.update(label.text for label in topic.findall("Labels"))
                self.assertEqual(topic.findtext("CreationAuthor"), "recette@example.org")
        self.assertIn("IfcWall – R+1 – Pset_WallCommon", titles)
        self.assertTrue({"Checkers", "Architecture", "Structure", "IDS", "PSet"} <= labels)

    def test_viewpoint_selects_and_colors_elements(self):
        for name in self.zip.namelist():
            if not name.endswith("/markup.bcf"):
                continue
            markup = ET.fromstring(self.zip.read(name))
            if markup.find("Topic").findtext("Title").startswith("IfcWall"):
                folder = name.split("/")[0]
                vp = ET.fromstring(self.zip.read(f"{folder}/viewpoint.bcfv"))
                self.assertEqual(vp.get("Guid"), markup.find("Viewpoints").get("Guid"))
                selected = [c.get("IfcGuid") for c in vp.find("Components/Selection")]
                colored = [c.get("IfcGuid") for c in vp.find("Components/Coloring/Color")]
                self.assertEqual(selected, [GUID_A, GUID_B])
                self.assertEqual(colored, selected)

    @unittest.skipUnless(XSD_DIR, "BCF_XSD_DIR non défini : validation XSD sautée")
    def test_valid_against_official_schemas(self):
        from lxml import etree
        schemas = {ext: etree.XMLSchema(etree.parse(os.path.join(XSD_DIR, xsd)))
                   for ext, xsd in ((".bcf", "markup.xsd"), (".bcfv", "visinfo.xsd"),
                                    (".version", "version.xsd"), (".bcfp", "project.xsd"))}
        for name in self.zip.namelist():
            schema = schemas[os.path.splitext(name)[1]]
            document = etree.fromstring(self.zip.read(name))
            self.assertTrue(schema.validate(document), f"{name} : {schema.error_log}")


if __name__ == "__main__":
    unittest.main()
