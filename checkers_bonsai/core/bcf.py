"""Export des anomalies de Checkers au format BCF 2.1 (BIM Collaboration
Format, buildingSMART).

Un sujet (topic) BCF par anomalie (niveau, classe IFC, exigence), avec un
point de vue qui sélectionne et colore en rouge les éléments concernés. Le
fichier .bcf s'ouvre dans les outils des maîtres d'œuvre et des entreprises
(BIMcollab, Solibri, Revit/Navisworks via plugin, Bonsai...) : chaque
non-conformité y est rattachée aux objets de la maquette, au lieu de vivre
seulement dans un tableau Excel.

Les identifiants des sujets sont stables : la même anomalie (même projet,
même lot, même niveau, même classe, même exigence) garde le même identifiant
d'une recette à l'autre. Réimporter le BCF de la livraison suivante met donc
à jour les sujets existants au lieu de les dupliquer, et les anomalies
levées depuis la recette précédente sont transmises comme sujets clos.

Version 2.1 plutôt que 3.0 : c'est la plus largement prise en charge par
les logiciels du marché. Les fichiers produits sont validés contre les
schémas XSD officiels (voir tests/test_recette.py).

Bibliothèque standard uniquement.
"""
import re
import uuid
import zipfile
from datetime import datetime, timezone
from typing import Iterable, Optional
from xml.etree import ElementTree as ET

from .excel_safety import clean_text
from .clashes import stable_requirement

BCF_VERSION = "2.1"
# Couleur des éléments non conformes dans les points de vue (RGB hexadécimal,
# majuscules, comme l'exige visinfo.xsd).
ANOMALY_COLOR = "E53935"
PRIORITIES = {"Bloquant": "Critical", "Majeur": "Major", "Mineur": "Minor"}
OPEN, CLOSED = "Open", "Closed"

# IfcGuid : 22 caractères de l'alphabet base64 d'IFC (visinfo.xsd). Un
# identifiant non conforme rendrait tout le fichier invalide : il est écarté
# et compté (voir write_bcf), plutôt que de produire un BCF refusé à
# l'ouverture.
_IFC_GUID_RE = re.compile(r"^[0-9A-Za-z_$]{22}$")
_NAMESPACE = uuid.uuid5(uuid.NAMESPACE_URL, "https://github.com/halhee/Checkers/bcf")


def stable_guid(*parts: str) -> str:
    """Identifiant déterministe : mêmes éléments de clé, même identifiant."""
    return str(uuid.uuid5(_NAMESPACE, "|".join(parts)))


def topic_guid(project_key: str, lot: str, anomaly: dict) -> str:
    source = anomaly.get("source") or ""
    return stable_guid(project_key, lot, source, anomaly["level"], anomaly["ifc_class"],
                       stable_requirement(source, anomaly["requirement"]))


def _xml(root: ET.Element) -> bytes:
    ET.indent(root)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _sub(parent: ET.Element, tag: str, text: Optional[str] = None, **attrs) -> ET.Element:
    element = ET.SubElement(parent, tag, {k: clean_text(v) for k, v in attrs.items() if v is not None})
    if text is not None:
        element.text = clean_text(text)
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
    requirement, source = anomaly["requirement"], anomaly.get("source")
    has_elements = bool(anomaly.get("guids"))
    if source == "IDS":
        name = requirement[len("IDS : "):] if requirement.startswith("IDS : ") else requirement
        what = f"Non-conformité à la spécification {name}."
        if not has_elements:
            what += " Aucun élément de la maquette ne correspond à cette spécification obligatoire."
    elif source == "Structure":
        what = f"Contrôle de structure non satisfait : {requirement}."
    elif source == "Terrain":
        name = requirement[len("Terrain : "):] if requirement.startswith("Terrain : ") else requirement
        what = f"Écart entre la maquette et le relevé fait sur site : {name}."
    elif source == "Cohérence":
        what = f"{requirement} : les maquettes de ces lots ne partagent pas le même cadre de référence."
    elif source == "Conflit":
        what = f"{requirement}."
        if anomaly.get("related_guids"):
            what += (f" Éléments de l'autre lot concernés : {len(anomaly['related_guids'])} (sélectionnés dans le "
                     f"point de vue, avec ceux de cette maquette).")
    else:
        what = f"PSet « {requirement} » absent ou incomplet (paramètre manquant, vide ou de type incorrect)."
    lines = [marking] if marking else []
    lines += [
        what,
        f"Niveau : {anomaly['level']} — classe IFC : {anomaly['ifc_class']}" + (f" — lot : {lot}" if lot else "") + ".",
        (f"Éléments concernés : {anomaly['count']} (sélectionnés dans le point de vue)." if has_elements
         else "Constat portant sur la maquette entière, sans élément à sélectionner."),
    ]
    if referentials:
        lines.append(f"Référentiel : {referentials}.")
    lines.append("Détail élément par élément : fiche de recette Checkers (onglets de détail).")
    return "\n".join(lines)


def _markup_file(guid: str, viewpoint_guid: str, anomaly: dict, index: int, status: str, author: str, date: str,
                 lot: str, phase: str, referentials: str, marking: str, labels: list,
                 comment: Optional[str]) -> bytes:
    root = ET.Element("Markup")
    topic = _sub(root, "Topic", Guid=guid, TopicType="Error", TopicStatus=status)
    _sub(topic, "Title", topic_title(anomaly))
    _sub(topic, "Priority", PRIORITIES.get(anomaly.get("severity") or "", "Normal"))
    _sub(topic, "Index", str(index))
    for label in labels + [anomaly.get("source") or "PSet"]:
        _sub(topic, "Labels", label)
    _sub(topic, "CreationDate", date)
    _sub(topic, "CreationAuthor", author)
    if phase:
        _sub(topic, "Stage", phase)
    _sub(topic, "Description", topic_description(anomaly, lot, referentials, marking))
    if comment:
        node = _sub(root, "Comment", Guid=stable_guid(guid, "comment", date))
        _sub(node, "Date", date)
        _sub(node, "Author", author)
        _sub(node, "Comment", comment)
    viewpoint = _sub(root, "Viewpoints", Guid=viewpoint_guid)
    _sub(viewpoint, "Viewpoint", "viewpoint.bcfv")
    return _xml(root)


# Caméra des points de vue de conflits : placée en retrait et en surplomb
# du point de conflit, tournée vers lui (coordonnées du projet, en mètres).
CAMERA_OFFSET = (-4.0, -4.0, 3.0)
CAMERA_FIELD_OF_VIEW = 60.0


def _normalize(vector):
    length = sum(c * c for c in vector) ** 0.5 or 1.0
    return tuple(c / length for c in vector)


def _cross(a, b):
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0])


def _camera(parent, point) -> None:
    eye = tuple(p + o for p, o in zip(point, CAMERA_OFFSET))
    direction = _normalize(tuple(p - e for p, e in zip(point, eye)))
    up = _normalize(_cross(_cross(direction, (0.0, 0.0, 1.0)), direction))
    camera = _sub(parent, "PerspectiveCamera")
    for tag, values in (("CameraViewPoint", eye), ("CameraDirection", direction), ("CameraUpVector", up)):
        node = _sub(camera, tag)
        for axis, value in zip("XYZ", values):
            _sub(node, axis, f"{value:.6f}")
    _sub(camera, "FieldOfView", f"{CAMERA_FIELD_OF_VIEW:g}")


def _viewpoint_file(viewpoint_guid: str, guids: list, color: bool, related: list = (), point=None) -> bytes:
    """Point de vue : les éléments concernés sont sélectionnés et colorés, le
    reste du modèle reste visible. `related` : éléments d'un autre lot
    (conflits), sélectionnés mais pas colorés. Avec `point` (conflit), une
    caméra est placée face au point ; sinon pas de caméra (facultative en BCF
    2.1) et l'outil de destination cadre la sélection."""
    root = ET.Element("VisualizationInfo", Guid=viewpoint_guid)
    components = _sub(root, "Components")
    if guids or related:
        selection = _sub(components, "Selection")
        for guid in list(guids) + list(related):
            _sub(selection, "Component", IfcGuid=guid)
    _sub(components, "Visibility", DefaultVisibility="true")
    if guids and color:
        node = _sub(_sub(components, "Coloring"), "Color", Color=ANOMALY_COLOR)
        for guid in guids:
            _sub(node, "Component", IfcGuid=guid)
    if point is not None and len(point) == 3:
        _camera(root, tuple(float(v) for v in point))
    return _xml(root)


def write_bcf(path: str, anomalies: Iterable[dict], author: str = "Checkers",
              project_name: Optional[str] = None, project_key: str = "", lot: str = "", referentials: str = "",
              marking: str = "", phase: str = "", resolved: Iterable[dict] = (),
              creation_date: Optional[datetime] = None, admitted: Iterable[dict] = ()) -> dict:
    """Écrit un fichier BCF 2.1 avec un sujet ouvert par anomalie.

    `anomalies` : dicts level, ifc_class, requirement, source, severity,
    count, guids (voir recette.run_recette). `resolved` : groupes d'anomalies
    levées depuis la recette précédente (journal.compare) ; ceux qui n'ont
    plus aucun élément en anomalie sont écrits comme sujets clos, comme les
    anomalies admises par dérogation (`admitted`, avec le motif en
    commentaire).
    `project_key` (GlobalId du projet IFC de préférence) entre dans les
    identifiants stables des sujets. La mention de confidentialité éventuelle
    figure en tête de chaque description et en étiquette.

    Retourne {"topics", "closed", "components", "skipped_guids"}."""
    moment = (creation_date or datetime.now(timezone.utc)).astimezone()
    date = moment.isoformat(timespec="seconds")
    author = (author or "").strip() or "Checkers"
    project_key = project_key or project_name or ""
    labels = [label for label in ("Checkers", lot, marking) if label]
    stats = {"topics": 0, "closed": 0, "components": 0, "skipped_guids": 0}
    open_guids = set()

    def clean(guids) -> list:
        kept = []
        for guid in guids or []:
            if isinstance(guid, str) and _IFC_GUID_RE.match(guid):
                kept.append(guid)
            else:
                stats["skipped_guids"] += 1
        return kept

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("bcf.version", _version_file())
        zf.writestr("project.bcfp", _project_file(stable_guid(project_key, "project"), project_name))
        index = 0
        for anomaly in anomalies:
            index += 1
            guid = topic_guid(project_key, lot, anomaly)
            open_guids.add(guid)
            guids = clean(anomaly.get("guids"))
            viewpoint = stable_guid(guid, "viewpoint")
            zf.writestr(f"{guid}/markup.bcf", _markup_file(
                guid, viewpoint, anomaly, index, OPEN, author, date, lot, phase, referentials, marking, labels, None))
            related = clean(anomaly.get("related_guids"))
            zf.writestr(f"{guid}/viewpoint.bcfv", _viewpoint_file(viewpoint, guids, color=True, related=related,
                                                                   point=anomaly.get("point")))
            stats["topics"] += 1
            stats["components"] += len(guids) + len(related)
        def closing_comment(group):
            if group.get("derogation"):
                d = group["derogation"]
                by = f" par {d['author']}" if d.get("author") else ""
                when = f" le {d['date'][8:10]}/{d['date'][5:7]}/{d['date'][:4]}" if len(d.get("date") or "") >= 10 \
                    else ""
                return f"Anomalie admise par dérogation (accordée{by}{when}) : {d.get('reason', '')}"
            comment = f"Anomalie levée, constatée à la recette du {moment:%d/%m/%Y}."
            if group.get("removed"):
                comment += f" Dont {group['removed']} élément(s) supprimé(s) de la maquette."
            return comment

        closed_guids = set()
        for group in list(admitted) + list(resolved):
            guid = topic_guid(project_key, lot, group)
            if guid in open_guids or guid in closed_guids:
                continue  # il reste des éléments en anomalie : le sujet reste ouvert
            closed_guids.add(guid)
            index += 1
            comment = closing_comment(group)
            viewpoint = stable_guid(guid, "viewpoint")
            zf.writestr(f"{guid}/markup.bcf", _markup_file(
                guid, viewpoint, group, index, CLOSED, author, date, lot, phase, referentials, marking, labels,
                comment))
            zf.writestr(f"{guid}/viewpoint.bcfv", _viewpoint_file(viewpoint, clean(group.get("guids")), color=False))
            stats["topics"] += 1
            stats["closed"] += 1
    return stats
