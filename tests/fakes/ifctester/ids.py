"""Doublure d'ifctester.ids : lecture minimale d'un IDS (titre, noms des
spécifications), sans validation XSD ni contrôle d'un modèle."""
from types import SimpleNamespace
from xml.etree import ElementTree as ET

NS = "{http://standards.buildingsmart.org/IDS}"


def open(path, validate=False):
    root = ET.parse(path).getroot()
    if root.tag != NS + "ids":
        raise ValueError("Not an IDS document")
    info = {child.tag[len(NS):]: child.text for child in root.find(NS + "info")}
    specs = [SimpleNamespace(name=s.get("name"), identifier=s.get("identifier"), ifcVersion=s.get("ifcVersion"))
             for s in root.iter(NS + "specification")]
    return SimpleNamespace(info=info, specifications=specs)
