"""Fichier Excel de règles PSet (exemple : examples/regles_exemple.xlsx).

Structure attendue :
- onglet « Element_Types » : colonne IFC_Class (classes contrôlées) ;
- un onglet par PSet (nom de l'onglet = nom du PSet) : colonnes IFC_Class,
  Parametre, Type (String, Int, Float/Number, Bool) ;
- facultatif : colonne « Corps_Etat » sur chaque ligne (vide = tous les corps
  d'état ; « ARC, STR » = règle partagée) et onglet « Corps_Etat » (colonnes
  Code, Nom) qui déclare les disciplines.

Le fichier est traité comme une entrée non fiable : taille limitée, archive
contrôlée avant ouverture (bombe de décompression), lecture seule et sans
évaluer de formule. openpyxl (fourni par Bonsai) n'est importé qu'au moment
de la lecture, pour que son absence donne un message clair.
"""
import os
import zipfile
from typing import Optional

RESERVED_SHEETS = ("Element_Types", "Corps_Etat")
SUPPORTED_TYPES = ("string", "int", "float", "number", "bool")
TYPE_LABELS = {bool: "booléen", int: "entier", float: "nombre", str: "texte"}

# Un fichier de règles fait quelques dizaines de Ko. Au-delà de ces plafonds,
# refus net plutôt qu'un chargement interminable dans Blender.
MAX_RULES_BYTES = 20 * 1024 * 1024
MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024
MAX_ARCHIVE_MEMBERS = 2000


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


def check_archive(path: str, label: str = "Fichier de règles") -> None:
    """Refuse un .xlsx qui n'est pas une archive saine : un fichier de
    quelques Ko peut se décompresser en plusieurs Go. `label` nomme le
    fichier dans les messages (ce contrôle sert aussi au relevé terrain)."""
    if not os.path.isfile(path):
        raise RulesFileError(f"{label} introuvable : {path}")
    if os.path.getsize(path) > MAX_RULES_BYTES:
        raise RulesFileError(f"{label} trop volumineux")
    try:
        with zipfile.ZipFile(path) as archive:
            members = archive.infolist()
    except (zipfile.BadZipFile, OSError):
        raise RulesFileError(f"{label} : ce n'est pas un classeur Excel (.xlsx) valide")
    if len(members) > MAX_ARCHIVE_MEMBERS or sum(m.file_size for m in members) > MAX_UNCOMPRESSED_BYTES:
        raise RulesFileError(f"{label} refusé : contenu décompressé anormalement volumineux")


def _open(path: str):
    check_archive(path)
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise RulesFileError("La bibliothèque openpyxl est introuvable : vérifiez que Bonsai est installé et activé")
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
    """{"classes": [IFC_Class...], "psets": {IFC_Class: {PSet: {Paramètre: Type}}},
    "warnings": [...]} filtrés pour le corps d'état `code` (None = toutes les
    lignes)."""
    workbook = _open(path)
    try:
        if "Element_Types" not in workbook.sheetnames:
            raise RulesFileError("Onglet « Element_Types » absent du fichier de règles")
        classes, declared = [], set()
        type_rows = list(_rows(workbook["Element_Types"]))
        if not type_rows or "IFC_Class" not in type_rows[0]:
            raise RulesFileError("Onglet « Element_Types » vide ou sans colonne « IFC_Class »")
        for row in type_rows:
            ifc_class = _text(row.get("IFC_Class"))
            if ifc_class:
                declared.add(ifc_class)
            if ifc_class and applies_to(row.get("Corps_Etat"), code) and ifc_class not in classes:
                classes.append(ifc_class)

        psets, undeclared = {}, set()
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
                if ifc_class not in declared:
                    undeclared.add(ifc_class)
                psets.setdefault(ifc_class, {}).setdefault(sheet_name, {})[param] = param_type
        warnings = []
        if undeclared:
            # Sans cet avertissement, ces règles ne seraient jamais appliquées
            # sans que personne ne le sache.
            warnings.append("Règles Excel ignorées : classes absentes de l'onglet « Element_Types » ("
                            + ", ".join(sorted(undeclared)) + ")")
        return {"classes": classes, "psets": psets, "warnings": warnings}
    finally:
        workbook.close()


def flatten(value) -> list:
    """Valeurs élémentaires d'une propriété : une propriété énumérée ou à
    liste donne plusieurs valeurs, une propriété bornée un dictionnaire."""
    if isinstance(value, dict):
        value = [v for k, v in value.items() if k != "id"]
    if isinstance(value, (list, tuple, set)):
        return [item for v in value for item in flatten(v)]
    return [value]


def is_empty(value) -> bool:
    """Absente, vide ou faite uniquement de blancs."""
    items = [v for v in flatten(value) if v is not None and not (isinstance(v, str) and not v.strip())]
    return not items


def _matches(value, param_type: str) -> bool:
    kind = param_type.lower()
    if kind == "bool":
        return isinstance(value, bool)
    if isinstance(value, bool):
        return False  # True vaut 1 en Python : jamais un nombre ni un texte ici
    if kind == "int":
        return isinstance(value, int) or (isinstance(value, float) and value.is_integer())
    if kind in ("float", "number"):
        return isinstance(value, (int, float))
    return isinstance(value, str)


def value_matches(value, param_type: str) -> bool:
    """Le type de la valeur IFC doit être celui attendu : Bool = booléen,
    Int = entier, Float/Number = nombre, String = texte. Un nombre saisi
    comme texte (« 2.5 ») n'est pas un nombre : il serait inexploitable par
    un outil de calcul ou de gestion."""
    return all(_matches(v, param_type) for v in flatten(value) if v is not None)


def found_type(value) -> str:
    values = [v for v in flatten(value) if v is not None]
    return TYPE_LABELS.get(type(values[0]), type(values[0]).__name__) if values else "vide"
