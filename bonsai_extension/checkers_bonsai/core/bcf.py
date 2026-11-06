"""Export des anomalies de Checkers au format BCF 2.1 (BIM Collaboration
Format, buildingSMART).

Un sujet (topic) BCF par anomalie (niveau, classe IFC, PSet ou
spécification IDS), avec un point de vue qui sélectionne et colore en rouge
les éléments concernés. Le fichier .bcf s'ouvre dans les outils des
maîtres d'œuvre et des entreprises (BIMcollab, Solibri, Revit/Navisworks
via plugin, Bonsai...) : chaque non-conformité y est rattachée aux objets
de la maquette, au lieu de vivre seulement dans un tableau Excel.

Version 2.1 plutôt que 3.0 : c'est la plus largement prise en charge par
les logiciels du marché. Les fichiers produits sont validés contre les
schémas XSD officiels (voir tests/test_bonsai_extension.py).

Bibliothèque standard uniquement.
"""
import re
import uuid
import zipfile
from datetime import datetime, timezone
from typing import Iterable, Optional
from xml.etree import ElementTree as ET

BCF_VERSION = "2.1"
# Couleur des éléments non conformes dans les points de vue (RGB hexadécimal,
# majuscules, comme l'exige visinfo.xsd).
ANOMALY_COLOR = "E53935"

# IfcGuid : 22 caractères de l'alphabet base64 d'IFC (visinfo.xsd). Un
# identifiant non conforme rendrait tout le fichier invalide : il est écarté
# et compté (voir write_bcf), plutôt que de produire un BCF refusé à
# l'ouverture.
_IFC_GUID_RE = re.compile(r"^[0-9A-Za-z_$]{22}$")


def _xml(root: ET.Element) -> bytes:
    ET.indent(root)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sub(parent: ET.Element, tag: str, text: Optional[str] = None, **attrs) -> ET.Element:
    element = ET.SubElement(parent, tag, {k: v for k, v in attrs.items() if v is not None})
    if text is not None:
        element.text = text
    return element


def _version_file() -> bytes:
    root = ET.Element("Version", VersionId=BCF_VERSION)
    _sub(root, "DetailedVersion", BCF_VERSION)
    return _xml(root)


def _project_file(project_id: str, project_name: Optional[str]) -> bytes:
    root = ET.Element("ProjectExtension")
    project = _sub(root, "Project", ProjectId=project_id)
    if project_name:
        _sub(project, "Name", project_name)
    _sub(root, "ExtensionSchema", "")
    return _xml(root)


def topic_title(anomaly: dict) -> str:
    """Titre court, lisible dans la liste des sujets d'un outil BCF."""
    return f"{anomaly['ifc_class']} – {anomaly['level']} – {anomaly['requirement']}"


def topic_description(anomaly: dict, lot: str = "", referentials: str = "", marking: str = "") -> str:
    requirement = anomaly["requirement"]
    if anomaly.get("source") == "IDS":
        what = f"Non-conformité à la spécification {requirement[len('IDS : '):] if requirement.startswith('IDS : ') else requirement}."
    else:
        what = f"PSet « {requirement} » absent ou incomplet (paramètre manquant, vide ou de type incorrect)."
    lines = [marking] if marking else []
    lines += [
        what,
        f"Niveau : {anomaly['level']} — classe IFC : {anomaly['ifc_class']}" + (f" — lot : {lot}" if lot else "") + ".",
        f"Éléments concernés : {anomaly['count']} (sélectionnés dans le point de vue).",
    ]
    if referentials:
        lines.append(f"Référentiel : {referentials}.")
    lines.append("Détail élément par élément : rapport de recette Checkers (onglets « Détail PSet » / « IDS »).")
    return "\n".join(lines)


def _markup_file(topic_guid: str, viewpoint_guid: str, anomaly: dict, index: int, author: str, date: str,
                 lot: str, referentials: str, marking: str, labels: list) -> bytes:
    root = ET.Element("Markup")
    topic = _sub(root, "Topic", Guid=topic_guid, TopicType="Error", TopicStatus="Open")
    _sub(topic, "Title", topic_title(anomaly))
    _sub(topic, "Priority", "Normal")
    _sub(topic, "Index", str(index))
    for label in labels + [anomaly.get("source") or "PSet"]:
        _sub(topic, "Labels", label)
    _sub(topic, "CreationDate", date)
    _sub(topic, "CreationAuthor", author)
    _sub(topic, "Description", topic_description(anomaly, lot, referentials, marking))
    viewpoint = _sub(root, "Viewpoints", Guid=viewpoint_guid)
    _sub(viewpoint, "Viewpoint", "viewpoint.bcfv")
    return _xml(root)


def _viewpoint_file(viewpoint_guid: str, guids: list) -> bytes:
    """Point de vue sans caméra (facultative en BCF 2.1) : les éléments
    concernés sont sélectionnés et colorés, le reste du modèle reste
    visible ; l'outil de destination cadre la sélection."""
    root = ET.Element("VisualizationInfo", Guid=viewpoint_guid)
    components = _sub(root, "Components")
    if guids:
        selection = _sub(components, "Selection")
        for guid in guids:
            _sub(selection, "Component", IfcGuid=guid)
    _sub(components, "Visibility", DefaultVisibility="true")
    if guids:
        color = _sub(_sub(components, "Coloring"), "Color", Color=ANOMALY_COLOR)
        for guid in guids:
            _sub(color, "Component", IfcGuid=guid)
    return _xml(root)


def write_bcf(path: str, anomalies: Iterable[dict], author: str = "Checkers",
              project_name: Optional[str] = None, lot: str = "", referentials: str = "",
              marking: str = "", phase: str = "", creation_date: Optional[datetime] = None) -> dict:
    """Écrit un fichier BCF 2.1 avec un sujet par anomalie.

    `anomalies` : dicts level, ifc_class, requirement, source, count, guids
    (voir recette.run_recette). La mention de confidentialité éventuelle figure en
    tête de chaque description et en étiquette.

    Retourne {"topics": nb de sujets, "components": nb d'éléments référencés,
    "skipped_guids": nb d'identifiants écartés car non conformes}."""
    date = (creation_date or datetime.now(timezone.utc)).astimezone().isoformat(timespec="seconds")
    author = (author or "").strip() or "Checkers"
    labels = [label for label in ("Checkers", lot, phase, marking) if label]
    stats = {"topics": 0, "components": 0, "skipped_guids": 0}

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("bcf.version", _version_file())
        zf.writestr("project.bcfp", _project_file(str(uuid.uuid4()), project_name))
        for index, anomaly in enumerate(anomalies, 1):
            guids = []
            for guid in anomaly.get("guids") or []:
                if isinstance(guid, str) and _IFC_GUID_RE.match(guid):
                    guids.append(guid)
                else:
                    stats["skipped_guids"] += 1
            topic_guid, viewpoint_guid = str(uuid.uuid4()), str(uuid.uuid4())
            zf.writestr(f"{topic_guid}/markup.bcf", _markup_file(
                topic_guid, viewpoint_guid, anomaly, index, author, date, lot, referentials, marking, labels))
            zf.writestr(f"{topic_guid}/viewpoint.bcfv", _viewpoint_file(viewpoint_guid, guids))
            stats["topics"] += 1
            stats["components"] += len(guids)
    return stats
