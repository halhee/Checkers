"""Traçabilité de la recette : ce qui a été contrôlé, avec quoi, et quand.

Les empreintes SHA-256 identifient sans ambiguïté la maquette et les
référentiels utilisés : en cas de contestation d'une recette, on peut
prouver quelle version de quel fichier a été contrôlée.
"""
import hashlib
import os
from datetime import datetime


def sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_info(path: str) -> dict:
    """Nom, taille, date de modification et empreinte d'un fichier, ou {}
    s'il n'existe pas (ex. projet Bonsai jamais enregistré)."""
    if not path or not os.path.isfile(path):
        return {}
    return {
        "name": os.path.basename(path),
        "size": os.path.getsize(path),
        "modified": datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d/%m/%Y %H:%M"),
        "sha256": sha256(path),
    }


def _header_value(getter):
    try:
        value = getter()
    except Exception:
        return ""
    if isinstance(value, (tuple, list)):
        value = ", ".join(str(v) for v in value if v)
    return str(value or "").strip()


def model_info(ifc_file, path: str = "") -> dict:
    """En-tête du fichier IFC (application auteur, organisation, date
    d'export), schéma, projet, et empreinte du fichier sur disque."""
    header = getattr(ifc_file, "header", None)
    file_name = getattr(header, "file_name", None)
    projects = []
    try:
        projects = ifc_file.by_type("IfcProject")
    except Exception:
        pass
    info = {
        "schema": _header_value(lambda: ifc_file.schema_identifier) or _header_value(lambda: ifc_file.schema),
        "project": (projects[0].Name or "") if projects else "",
        "originating_system": _header_value(lambda: file_name.originating_system),
        "author": _header_value(lambda: file_name.author),
        "organization": _header_value(lambda: file_name.organization),
        "timestamp": _header_value(lambda: file_name.time_stamp),
    }
    info["file"] = file_info(path)
    return info
