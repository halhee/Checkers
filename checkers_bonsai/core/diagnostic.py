"""Diagnostic de l'installation : vérifie, sur le poste, que tout ce dont
l'extension a besoin est présent et répond comme attendu.

Utile au premier lancement et après une mise à jour de Blender ou de Bonsai,
en particulier sur un poste sans accès Internet, où l'on ne peut pas chercher
la cause d'une erreur en ligne. Chaque ligne dit ce qui a été vérifié et, en
cas d'échec, quoi faire.
"""
import importlib
import os

EXAMPLES_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "examples")
EXAMPLE_IDS = ("exploitation_maintenance.ids", "conception_minimale.ids")

# Fonctions des bibliothèques fournies par Bonsai dont l'extension dépend.
REQUIRED = (
    ("ifcopenshell.util.element", ("get_psets", "get_container", "get_type", "get_predefined_type")),
    ("ifcopenshell.util.unit", ("get_project_unit",)),
    ("ifcopenshell.util.classification", ("get_references",)),
    ("ifcopenshell.util.system", ("get_element_systems",)),
    ("ifcopenshell.util.selector", ("filter_elements",)),
    # Un tuple : l'une des fonctions suffit (get_local_placement avant 0.9).
    ("ifcopenshell.util.placement", (("get_placement", "get_local_placement"),)),
    ("ifctester.ids", ("open",)),
    ("ifcclash.ifcclash", ("Clasher", "ClashSettings")),
    ("openpyxl", ("Workbook", "load_workbook")),
)
BONSAI_HINT = "vérifiez que Bonsai est installé et activé"


def _line(label: str, ok: bool, detail: str = "") -> dict:
    return {"label": label, "ok": ok, "detail": detail}


def _version(module_name: str) -> str:
    try:
        module = importlib.import_module(module_name.split(".")[0])
    except Exception:
        return ""
    return str(getattr(module, "version", None) or getattr(module, "__version__", None) or "")


def run(ifc_file=None, bonsai_found=None) -> list:
    """Liste de {"label", "ok", "detail"}. `bonsai_found` : True, False, ou
    None si l'appelant ne sait pas (hors de Blender)."""
    lines = []
    if bonsai_found is not None:
        lines.append(_line("Bonsai", bonsai_found, "" if bonsai_found else BONSAI_HINT))

    for module_name, functions in REQUIRED:
        try:
            module = importlib.import_module(module_name)
        except Exception as exc:
            lines.append(_line(module_name, False, f"introuvable ({type(exc).__name__}) : {BONSAI_HINT}"))
            continue
        missing = [" ou ".join(name) if isinstance(name, tuple) else name for name in functions
                   if not any(hasattr(module, n) for n in (name if isinstance(name, tuple) else (name,)))]
        version = _version(module_name)
        if missing:
            lines.append(_line(module_name, False, "fonction(s) absente(s) : " + ", ".join(missing)
                               + " (version de Bonsai non prise en charge)"))
        else:
            lines.append(_line(module_name, True, f"version {version}" if version else ""))

    # Les IDS d'exemple passent-ils la validation du schéma officiel sur ce poste ?
    try:
        from . import ids as ids_check
        for name in EXAMPLE_IDS:
            document = ids_check.load_ids(os.path.join(EXAMPLES_DIR, name))
            lines.append(_line(f"IDS d'exemple {name}", True, f"{len(document.specifications)} spécifications"))
    except Exception as exc:
        lines.append(_line("IDS d'exemple", False, str(exc)))

    if ifc_file is None:
        lines.append(_line("Modèle IFC", False, "aucun projet IFC ouvert dans Bonsai"))
    else:
        try:
            schema = getattr(ifc_file, "schema_identifier", None) or ifc_file.schema
            count = len(ifc_file.by_type("IfcElement"))
            lines.append(_line("Modèle IFC", True, f"schéma {schema}, {count} élément(s)"))
        except Exception as exc:
            lines.append(_line("Modèle IFC", False, f"illisible ({type(exc).__name__})"))
    return lines


def to_text(lines: list) -> str:
    return "\n".join(f"{'✓' if line['ok'] else '✗'} {line['label']}" + (f" : {line['detail']}" if line["detail"] else "")
                     for line in lines)
