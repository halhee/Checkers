"""Relevé terrain : confronter la maquette à ce qui est constaté sur site.

Deux temps :
1. `write_survey` produit une fiche de relevé (classeur Excel) à partir des
   équipements de la maquette, classés par niveau et par local. Sur site, on
   y note si chaque équipement est présent, le repère et le numéro de série
   lus sur la plaque, et les équipements trouvés qui ne sont pas modélisés
   (lignes ajoutées en bas, sans GlobalId).
2. `check_survey` relit la fiche remplie et la compare à la maquette ouverte.
   Les écarts deviennent des anomalies comme les autres (source « Terrain ») :
   sélectionnables dans la vue 3D, portées sur la fiche de recette et dans le
   BCF, et suivies d'une livraison à l'autre.

La comparaison se fait avec les valeurs actuelles de la maquette, pas avec
celles recopiées dans la fiche : une maquette corrigée depuis le relevé est
jugée sur son état corrigé.

La fiche remplie est traitée comme une entrée non fiable (voir rules.py).
"""
from collections import defaultdict

import ifcopenshell.util.element

from .excel_safety import excel_safe, header_footer_text, neutralize_formulas
from .rules import check_archive
from .spatial import MODEL_LEVEL

SHEET = "Relevé terrain"
# (clé, intitulé de colonne, largeur, colonne à remplir sur site ?)
COLUMNS = (
    ("level", "Niveau / partie", 18, False), ("space", "Local", 24, False),
    ("tag", "Repère (maquette)", 18, False), ("name", "Désignation", 30, False),
    ("ifc_class", "Classe IFC", 22, False), ("serial", "N° de série (maquette)", 20, False),
    ("present", "Présent sur site (oui / non)", 16, True), ("tag_found", "Repère relevé", 18, True),
    ("serial_found", "N° de série relevé", 20, True), ("note", "Observation", 36, True),
    ("guid", "GlobalId", 26, False),
)
REQUIRED_COLUMNS = ("guid", "present")
INSTRUCTIONS = (
    "Mode d'emploi",
    "1. Sur site, remplir les colonnes jaunes de l'onglet « Relevé terrain ».",
    "2. « Présent sur site » : oui ou non. Laisser vide un équipement qui n'a pas été vérifié.",
    "3. « Repère relevé » et « N° de série relevé » : ce qui est lu sur l'équipement ou sa plaque. "
    "Laisser vide si la valeur n'a pas été relevée.",
    "4. Équipement trouvé sur site mais absent de la liste : ajouter une ligne en bas, sans GlobalId, "
    "avec sa désignation, son local et ce qui a été relevé.",
    "5. Ne pas modifier la colonne « GlobalId » : elle relie chaque ligne à la maquette.",
    "6. Dans Blender, panneau Checkers : indiquer ce fichier dans « Relevé terrain », puis lancer la recette.",
)
MAX_ROWS = 200000

ABSENT = "Terrain : absent sur site"
TAG_DIFFERS = "Terrain : repère différent"
SERIAL_DIFFERS = "Terrain : n° de série différent"
UNMODELLED = "Terrain : équipement non modélisé"
UNMODELLED_CLASS = "Équipement"

_YES = {"oui", "o", "x", "yes", "y", "vrai", "true", "1", "ok", "présent", "present"}
_NO = {"non", "n", "no", "faux", "false", "0", "absent", "manquant"}


class SurveyFileError(ValueError):
    """Fiche de relevé illisible ou mal structurée."""


def _text(value) -> str:
    """Valeur de cellule en texte. Excel transforme volontiers « 00123 » en
    nombre : 123.0 redevient « 123 »."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def _same(a, b) -> bool:
    return _text(a).casefold() == _text(b).casefold()


def parse_presence(value):
    """True, False, ou None si la cellule est vide. Lève ValueError si la
    réponse n'est pas comprise."""
    if isinstance(value, bool):
        return value
    text = _text(value).casefold()
    if not text:
        return None
    if text in _YES:
        return True
    if text in _NO:
        return False
    raise ValueError(text)


def write_survey(path: str, equipment: list, model: dict, marking: str = "", date=None) -> None:
    """`equipment` : lignes de inventory.collect_equipment."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation

    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    model_fill = PatternFill("solid", start_color="E7E6E6")
    site_fill = PatternFill("solid", start_color="FFF2CC")
    workbook = Workbook()
    ws = workbook.active
    ws.title = SHEET
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    if marking:
        for part in (ws.oddHeader.center, ws.oddFooter.center):
            part.text, part.size, part.color = header_footer_text(marking), 11, "C00000"
    ws.oddFooter.right.text = "Page &P / &N"
    ws.print_title_rows = "1:1"

    for col, (_, label, width, on_site) in enumerate(COLUMNS, 1):
        cell = ws.cell(1, col, label)
        cell.font = Font(bold=True)
        cell.fill = site_fill if on_site else model_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = width
    for r, item in enumerate(equipment, 2):
        ws.row_dimensions[r].height = 24          # de la place pour écrire à la main sur la fiche imprimée
        for col, (key, _, _, on_site) in enumerate(COLUMNS, 1):
            cell = ws.cell(r, col, None if on_site else excel_safe(item.get(key, "")))
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if on_site:
                cell.fill = site_fill
    last = max(len(equipment) + 1, 2)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last}"
    present_column = get_column_letter([key for key, *_ in COLUMNS].index("present") + 1)
    choices = DataValidation(type="list", formula1='"oui,non"', allow_blank=True)
    ws.add_data_validation(choices)
    # Quelques lignes de plus, pour les équipements ajoutés sur site.
    choices.add(f"{present_column}2:{present_column}{last + 200}")

    notes = workbook.create_sheet("Mode d'emploi")
    notes.column_dimensions["A"].width = 120
    file_info = model.get("file") or {}
    lines = list(INSTRUCTIONS) + [
        "",
        f"Projet : {model.get('project') or '—'}",
        f"Maquette : {file_info.get('name') or '—'}",
        f"Empreinte SHA-256 : {file_info.get('sha256') or '—'}",
        f"Fiche produite le : {date.strftime('%d/%m/%Y %H:%M') if date else '—'}",
    ]
    if marking:
        lines.insert(0, marking)
    for r, line in enumerate(lines, 1):
        cell = notes.cell(r, 1, excel_safe(line))
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    notes.cell(2 if marking else 1, 1).font = Font(bold=True, size=13)
    if marking:
        notes.cell(1, 1).font = Font(bold=True, color="C00000")

    neutralize_formulas(workbook)
    workbook.save(path)


def load_survey(path: str) -> list:
    """Lignes de la fiche remplie : dicts aux clés de COLUMNS, plus « row »
    (numéro de ligne dans le classeur)."""
    check_archive(path, "Relevé terrain")
    try:
        from openpyxl import load_workbook
    except ImportError:
        raise SurveyFileError("La bibliothèque openpyxl est introuvable : vérifiez que Bonsai est installé et activé")
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise SurveyFileError(f"Relevé terrain illisible ({type(exc).__name__})") from exc
    try:
        if SHEET not in workbook.sheetnames:
            raise SurveyFileError(f"Onglet « {SHEET} » absent du relevé terrain")
        rows = workbook[SHEET].iter_rows(values_only=True)
        header = [_text(h).casefold() for h in next(rows, None) or ()]
        index = {key: header.index(label.casefold()) for key, label, _, _ in COLUMNS if label.casefold() in header}
        missing = [label for key, label, _, _ in COLUMNS if key in REQUIRED_COLUMNS and key not in index]
        if missing:
            raise SurveyFileError("Relevé terrain : colonne(s) absente(s) : " + ", ".join(missing))
        result = []
        for number, values in enumerate(rows, 2):
            if number > MAX_ROWS:
                raise SurveyFileError("Relevé terrain trop volumineux")
            if values is None or all(not _text(v) for v in values):
                continue
            row = {key: (values[i] if i < len(values) else None) for key, i in index.items()}
            row["row"] = number
            result.append(row)
        return result
    finally:
        workbook.close()


def _model_serial(element) -> str:
    try:
        psets = ifcopenshell.util.element.get_psets(element)
    except Exception:
        return ""
    return _text((psets.get("Pset_ManufacturerOccurrence") or {}).get("SerialNumber"))


def check_survey(ifc_file, rows: list, level_of) -> dict:
    """Compare la fiche remplie à la maquette. Retourne :
    - verified : équipements de la maquette vérifiés sur site ; conform : sans écart ;
    - unverified : lignes laissées vides ; unknown : lignes dont le GlobalId
      n'existe plus dans la maquette ; unmodelled : équipements ajoutés sur site ;
    - checks / pass_rate : vérifications (verified + unmodelled) et taux sans écart ;
    - anomalies : {(niveau, classe IFC, écart): ensemble de GlobalId} ;
    - model_anomalies : {(niveau, classe, écart): nombre}, sans élément à sélectionner ;
    - conform_guids, details, warnings."""
    anomalies = defaultdict(set)
    details, warnings, conform_guids = [], [], []
    verified = unverified = unknown = unmodelled = unreadable = 0
    seen = set()

    for row in rows:
        guid = _text(row.get("guid"))
        tag_found, serial_found, note = (_text(row.get(k)) for k in ("tag_found", "serial_found", "note"))
        try:
            present = parse_presence(row.get("present"))
        except ValueError:
            unreadable += 1
            present = None

        if not guid:
            # Ligne ajoutée sur site : un équipement réel que la maquette ignore.
            if present is False or not any((present, tag_found, serial_found, note, _text(row.get("name")))):
                continue
            unmodelled += 1
            details.append({"level": _text(row.get("level")), "space": _text(row.get("space")),
                            "ifc_class": "", "guid": "", "name": _text(row.get("name")),
                            "problem": UNMODELLED, "model": "", "found": tag_found or serial_found, "note": note})
            continue
        if guid in seen:
            warnings.append(f"Relevé terrain : GlobalId en double (ligne {row['row']}), ligne ignorée.")
            continue
        seen.add(guid)
        try:
            element = ifc_file.by_guid(guid)
        except Exception:
            element = None
        if element is None:
            unknown += 1
            continue
        if present is None and not tag_found and not serial_found:
            unverified += 1
            continue

        verified += 1
        level, ifc_class = level_of(element), element.is_a()
        problems = []
        if present is False:
            problems.append((ABSENT, "", ""))
        else:
            tag = _text(getattr(element, "Tag", None))
            if tag_found and not _same(tag_found, tag):
                problems.append((TAG_DIFFERS, tag, tag_found))
            serial = _model_serial(element)
            if serial_found and not _same(serial_found, serial):
                problems.append((SERIAL_DIFFERS, serial, serial_found))
        if not problems:
            conform_guids.append(guid)
        for problem, model_value, found in problems:
            anomalies[(level, ifc_class, problem)].add(guid)
            details.append({"level": level, "space": _text(row.get("space")), "ifc_class": ifc_class, "guid": guid,
                            "name": _text(getattr(element, "Name", None)), "problem": problem,
                            "model": model_value, "found": found, "note": note})

    if unknown:
        warnings.append(f"Relevé terrain : {unknown} ligne(s) désignent un élément absent de la maquette "
                        f"(relevé fait sur une autre version ?).")
    if unreadable:
        warnings.append(f"Relevé terrain : {unreadable} réponse(s) « Présent sur site » non comprise(s) "
                        f"(attendu : oui ou non).")
    if unverified:
        warnings.append(f"Relevé terrain : {unverified} équipement(s) non vérifié(s) sur site.")
    checks = verified + unmodelled
    return {
        "verified": verified, "conform": len(conform_guids), "unverified": unverified, "unknown": unknown,
        "unmodelled": unmodelled, "checks": checks,
        "pass_rate": len(conform_guids) / checks if checks else None,
        "anomalies": dict(anomalies),
        "model_anomalies": {(MODEL_LEVEL, UNMODELLED_CLASS, UNMODELLED): unmodelled} if unmodelled else {},
        "conform_guids": conform_guids, "details": details, "warnings": warnings,
    }
