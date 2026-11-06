"""Fichier Excel de règles PSet (format historique de Checkers).

Structure attendue :
- onglet « Element_Types » : colonne IFC_Class (classes contrôlées) ;
- un onglet par PSet (nom de l'onglet = nom du PSet) : colonnes IFC_Class,
  Parametre, Type (String, Int, Float/Number, Bool) ;
- facultatif : colonne « Corps_Etat » sur chaque ligne (vide = tous les corps
  d'état ; « ARC, STR » = règle partagée) et onglet « Corps_Etat » (colonnes
  Code, Nom) qui déclare les disciplines.

Lecture avec openpyxl seul (fourni par Bonsai), en lecture seule et sans
évaluer de formule.
"""
import os
from typing import Optional

from openpyxl import load_workbook

RESERVED_SHEETS = ("Element_Types", "Corps_Etat")
SUPPORTED_TYPES = {"string": str, "int": int, "float": float, "number": float, "bool": bool}

# Un fichier de règles fait quelques dizaines de Ko : au-delà, refus net
# plutôt qu'un chargement interminable dans Blender.
MAX_RULES_BYTES = 20 * 1024 * 1024


class RulesFileError(ValueError):
    """Fichier de règles illisible ou mal structuré."""


def _rows(sheet):
    """Lignes d'un onglet sous forme de dicts {en-tête: valeur}, en sautant
    les lignes entièrement vides."""
    iterator = sheet.iter_rows(values_only=True)
    try:
        header = [str(h).strip() if h is not None else "" for h in next(iterator)]
    except StopIteration:
        return
    for values in iterator:
        if values is None or all(v is None or (isinstance(v, str) and not v.strip()) for v in values):
            continue
        yield dict(zip(header, values))


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def applies_to(corps_etat_cell, code: Optional[str]) -> bool:
    """Une cellule Corps_Etat vide s'applique à tous ; sinon liste de codes
    séparés par « , » ou « ; »."""
    if not code:
        return True
    codes = [c.strip().upper() for c in _text(corps_etat_cell).replace(";", ",").split(",") if c.strip()]
    return not codes or code.upper() in codes


def _open(path: str):
    if not os.path.isfile(path):
        raise RulesFileError(f"Fichier de règles introuvable : {path}")
    if os.path.getsize(path) > MAX_RULES_BYTES:
        raise RulesFileError("Fichier de règles trop volumineux")
    try:
        return load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise RulesFileError(f"Fichier Excel illisible ({type(exc).__name__})") from exc


def load_registry(path: str) -> Optional[dict]:
    """{code: nom} de l'onglet « Corps_Etat », ou None s'il est absent."""
    workbook = _open(path)
    try:
        if "Corps_Etat" not in workbook.sheetnames:
            return None
        registry = {}
        for row in _rows(workbook["Corps_Etat"]):
            code, nom = _text(row.get("Code")), _text(row.get("Nom"))
            if code and nom:
                registry[code] = nom
        return registry or None
    finally:
        workbook.close()


def load_rules(path: str, code: Optional[str] = None) -> dict:
    """{"classes": [IFC_Class...], "psets": {IFC_Class: {PSet: {Paramètre: Type}}}}
    filtrés pour le corps d'état `code` (None = toutes les lignes)."""
    workbook = _open(path)
    try:
        if "Element_Types" not in workbook.sheetnames:
            raise RulesFileError("Onglet « Element_Types » absent du fichier de règles")
        classes = []
        for row in _rows(workbook["Element_Types"]):
            ifc_class = _text(row.get("IFC_Class"))
            if ifc_class and applies_to(row.get("Corps_Etat"), code) and ifc_class not in classes:
                classes.append(ifc_class)

        psets = {}
        for sheet_name in workbook.sheetnames:
            if sheet_name in RESERVED_SHEETS:
                continue
            rows = list(_rows(workbook[sheet_name]))
            if rows and not {"IFC_Class", "Parametre", "Type"} <= set(rows[0]):
                continue  # onglet de notes (ex. « Lisez-moi ») : ignoré
            for row in rows:
                ifc_class, param, param_type = (_text(row.get(k)) for k in ("IFC_Class", "Parametre", "Type"))
                if not (ifc_class and param) or not applies_to(row.get("Corps_Etat"), code):
                    continue
                if param_type.lower() not in SUPPORTED_TYPES:
                    raise RulesFileError(
                        f"Type « {param_type} » non reconnu (onglet {sheet_name}, {ifc_class}.{param}) : "
                        f"attendu String, Int, Float, Number ou Bool")
                psets.setdefault(ifc_class, {}).setdefault(sheet_name, {})[param] = param_type
        return {"classes": classes, "psets": psets}
    finally:
        workbook.close()


def value_matches(value, param_type: str) -> bool:
    """Même règle que l'application web : la valeur doit être convertible
    dans le type attendu (Int accepte 3.0, Bool accepte toute valeur)."""
    try:
        SUPPORTED_TYPES[param_type.lower()](value)
        return True
    except (ValueError, TypeError):
        return False
