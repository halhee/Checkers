"""Lien entre l'extension et le moteur de Checkers (checkers_core).

Volontairement sans `bpy` : tout ce qui est ici peut être testé hors de
Blender. L'interface (panneau, opérateurs) vit dans operators.py / ui.py.
"""
import importlib
import os
import sys
import tempfile
import time
from datetime import datetime

LIB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib")

AUTO = "AUTO"


def ensure_lib_on_path():
    """Rend importables les modules de Checkers copiés dans lib/ (carbon,
    ids_check, checkers_core...). Appelé à l'enregistrement de l'extension."""
    if LIB_DIR not in sys.path:
        sys.path.insert(0, LIB_DIR)


def remove_lib_from_path():
    if LIB_DIR in sys.path:
        sys.path.remove(LIB_DIR)


def bonsai_tool():
    """Module `bonsai.tool`, ou None si Bonsai n'est pas chargé. Bonsai est
    lui-même une extension (`bl_ext.<dépôt>.bonsai`) et s'enregistre aussi
    sous le nom `bonsai` ; on retombe sur une recherche dans sys.modules si
    ce n'est pas le cas dans une version future."""
    try:
        return importlib.import_module("bonsai.tool")
    except ImportError:
        pass
    for name in list(sys.modules):
        if name.endswith(".bonsai"):
            try:
                return importlib.import_module(name + ".tool")
            except ImportError:
                continue
    return None


def current_ifc():
    """Modèle IFC ouvert dans Bonsai (`ifcopenshell.file`), ou None."""
    tool = bonsai_tool()
    if tool is None:
        return None
    try:
        return tool.Ifc.get()
    except Exception:
        return None


def current_ifc_path() -> str:
    tool = bonsai_tool()
    if tool is None:
        return ""
    try:
        return tool.Ifc.get_path() or ""
    except Exception:
        return ""


def default_lot_name() -> str:
    path = current_ifc_path()
    return os.path.splitext(os.path.basename(path))[0] if path else "Maquette"


_registry_cache = {}


def load_corps_etat(rules_path: str):
    """Onglet "Corps_Etat" du fichier de règles ({code: nom}) ou None.
    Mis en cache par (chemin, date de modification) : appelé à chaque
    rafraîchissement du panneau pour remplir la liste déroulante."""
    if not rules_path or not os.path.isfile(rules_path):
        return None
    key = (rules_path, os.path.getmtime(rules_path))
    if key not in _registry_cache:
        ensure_lib_on_path()
        try:
            import corps_etat
            _registry_cache.clear()
            _registry_cache[key] = corps_etat.load_registry(rules_path)
        except Exception as exc:
            print(f"Checkers : lecture de l'onglet Corps_Etat impossible ({exc})")
            _registry_cache[key] = None
    return _registry_cache[key]


def run_analysis(ifc_file, lot_name: str, rules_path: str = "", ids_path: str = "",
                 corps_etat_code: str = AUTO, output_dir: str = None) -> dict:
    """Lance le moteur de Checkers (le même que l'application web) sur le
    modèle déjà ouvert. Retourne le résultat de process_files enrichi de
    "report_path" (rapport Excel complet) et "duration_s".

    Les préconisations IA sont désactivées ici (CrewAI/Ollama n'ont pas
    leur place dans Blender) ; tout le reste est identique à l'appli web :
    contrôle PSet (Excel), IDS, bilan carbone, Fiche de Recette."""
    if ifc_file is None:
        raise ValueError("Aucun modèle IFC ouvert dans Bonsai.")
    rules_path = rules_path or None
    ids_path = ids_path or None
    if rules_path is None and ids_path is None:
        raise ValueError("Choisissez au moins un fichier de règles (Excel et/ou IDS).")
    for path in (rules_path, ids_path):
        if path is not None and not os.path.isfile(path):
            raise ValueError(f"Fichier introuvable : {path}")

    ensure_lib_on_path()
    from checkers_core.analysis import process_files

    output_dir = output_dir or tempfile.mkdtemp(prefix="checkers_")
    report_path = os.path.join(output_dir, f"checkers_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    explicit_code = None if corps_etat_code in (None, "", AUTO) else corps_etat_code
    lot_name = lot_name.strip() or default_lot_name()

    start = time.time()
    result = process_files(
        output_dir, [(lot_name, ifc_file, explicit_code)], rules_path, report_path,
        label=lot_name, ids_file_path=ids_path,
        ids_filename=os.path.basename(ids_path) if ids_path else None,
        with_recommendations=False, collect_elements=True,
    )
    result["report_path"] = report_path
    result["duration_s"] = time.time() - start
    return result


def element_status(result: dict):
    """(conformes, non conformes) en ensembles de GlobalId. Un élément
    non conforme pour une règle l'est tout court (il peut apparaître dans
    les deux listes si plusieurs lots le contrôlent) ; les éléments en échec
    IDS sont non conformes même sans fichier Excel."""
    elements = result.get("elements") or {}
    invalid = set(elements.get("invalid", []))
    for anomaly in result.get("anomalies", []):
        invalid.update(anomaly.get("guids", []))
    valid = set(elements.get("valid", [])) - invalid
    return valid, invalid
