"""Inventaire d'exploitation : équipements techniques et locaux de la
maquette, dans un classeur Excel prêt à être repris dans un outil de gestion
de maintenance (GMAO) ou de patrimoine.

Trois onglets : la synthèse (taux de renseignement de chaque information,
qui mesure si la maquette est exploitable), les équipements et les locaux.
Pour un export normalisé COBie, voir le module « Facility Management »
(IfcFM) de Bonsai.
"""
import ifcopenshell.util.classification
import ifcopenshell.util.element
import ifcopenshell.util.system

from .excel_safety import excel_safe, neutralize_formulas
from .spatial import NO_LEVEL

# Équipements suivis en exploitation : la distribution technique, hors
# réseaux linéaires (tronçons et raccords, non suivis à l'unité), et les
# appareils de transport (ascenseurs, escaliers mécaniques).
EQUIPMENT_CLASSES = ("IfcDistributionElement", "IfcTransportElement")
EXCLUDED_CLASSES = ("IfcFlowSegment", "IfcFlowFitting", "IfcDistributionPort")

EQUIPMENT_COLUMNS = (
    ("tag", "Repère", 16), ("name", "Désignation", 30), ("ifc_class", "Classe IFC", 24),
    ("predefined_type", "Type prédéfini", 18), ("type_name", "Type (modèle d'objet)", 28),
    ("level", "Niveau / partie", 18), ("space", "Local", 22), ("system", "Système", 24),
    ("manufacturer", "Fabricant", 22), ("model", "Modèle", 22), ("serial", "N° de série", 18),
    ("warranty_end", "Fin de garantie", 16), ("classification", "Classification", 24), ("guid", "GlobalId", 26),
)
SPACE_COLUMNS = (
    ("number", "Numéro", 14), ("name", "Désignation", 32), ("level", "Niveau", 18),
    ("area", "Surface (unité du projet)", 22), ("classification", "Classification", 24), ("guid", "GlobalId", 26),
)
# Informations dont le taux de renseignement est mesuré dans la synthèse.
EQUIPMENT_MEASURED = ("tag", "name", "type_name", "space", "system", "manufacturer", "model", "serial",
                      "warranty_end", "classification")
SPACE_MEASURED = ("number", "name", "area", "classification")


def _by_type(ifc_file, ifc_class) -> list:
    try:
        return list(ifc_file.by_type(ifc_class))
    except Exception:
        return []


def _is_any(entity, classes) -> bool:
    for ifc_class in classes:
        try:
            if entity.is_a(ifc_class):
                return True
        except Exception:
            continue
    return False


def _safe(getter, default=""):
    """Une information illisible sur un élément ne doit pas faire échouer
    tout l'inventaire : la cellule reste vide."""
    try:
        value = getter()
    except Exception:
        return default
    return default if value is None else value


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _pset_value(psets: dict, *candidates) -> str:
    """Première valeur renseignée parmi des couples (PSet, propriété)."""
    for pset_name, prop in candidates:
        value = (psets.get(pset_name) or {}).get(prop)
        if value not in (None, "") and _text(value):
            return _text(value)
    return ""


def _classification(element) -> str:
    references = _safe(lambda: ifcopenshell.util.classification.get_references(element), ())
    labels = []
    for reference in references:
        # Identification en IFC4, ItemReference en IFC2X3.
        code = _text(getattr(reference, "Identification", None) or getattr(reference, "ItemReference", None))
        name = _text(getattr(reference, "Name", None))
        labels.append(" ".join(part for part in (code, name) if part))
    return " ; ".join(sorted(label for label in labels if label))


def _space_label(space) -> str:
    return " – ".join(part for part in (_text(space.Name), _text(getattr(space, "LongName", None))) if part)


def collect_equipment(ifc_file, level_of) -> list:
    seen, rows = set(), []
    for ifc_class in EQUIPMENT_CLASSES:
        for element in _by_type(ifc_file, ifc_class):
            if element in seen or _is_any(element, EXCLUDED_CLASSES):
                continue
            seen.add(element)
            psets = _safe(lambda: ifcopenshell.util.element.get_psets(element), {})
            element_type = _safe(lambda: ifcopenshell.util.element.get_type(element), None)
            container = _safe(lambda: ifcopenshell.util.element.get_container(element), None)
            systems = _safe(lambda: ifcopenshell.util.system.get_element_systems(element), ())
            rows.append({
                "tag": _text(getattr(element, "Tag", None)),
                "name": _text(getattr(element, "Name", None)),
                "ifc_class": element.is_a(),
                "predefined_type": _text(_safe(lambda: ifcopenshell.util.element.get_predefined_type(element))),
                "type_name": _text(getattr(element_type, "Name", None)) if element_type is not None else "",
                "level": "" if level_of(element) == NO_LEVEL else level_of(element),
                "space": _space_label(container) if container is not None and _is_any(container, ("IfcSpace",)) else "",
                "system": " ; ".join(sorted(_text(s.Name) for s in systems if _text(s.Name))),
                "manufacturer": _pset_value(psets, ("Pset_ManufacturerTypeInformation", "Manufacturer")),
                "model": _pset_value(psets, ("Pset_ManufacturerTypeInformation", "ModelReference"),
                                     ("Pset_ManufacturerTypeInformation", "ModelLabel")),
                "serial": _pset_value(psets, ("Pset_ManufacturerOccurrence", "SerialNumber")),
                "warranty_end": _pset_value(psets, ("Pset_Warranty", "WarrantyEndDate")),
                "classification": _classification(element),
                "guid": element.GlobalId,
            })
    rows.sort(key=lambda r: (r["level"], r["space"], r["ifc_class"], r["tag"], r["name"]))
    return rows


def collect_spaces(ifc_file, level_of) -> list:
    rows = []
    for space in _by_type(ifc_file, "IfcSpace"):
        psets = _safe(lambda: ifcopenshell.util.element.get_psets(space), {})
        area = None
        for pset_name in ("Qto_SpaceBaseQuantities", "BaseQuantities"):
            for prop in ("NetFloorArea", "GrossFloorArea"):
                value = (psets.get(pset_name) or {}).get(prop)
                if area is None and isinstance(value, (int, float)) and not isinstance(value, bool):
                    area = round(float(value), 2)
        rows.append({
            "number": _text(space.Name), "name": _text(getattr(space, "LongName", None)),
            "level": "" if level_of(space) == NO_LEVEL else level_of(space),
            "area": "" if area is None else area,
            "classification": _classification(space), "guid": space.GlobalId,
        })
    rows.sort(key=lambda r: (r["level"], r["number"], r["name"]))
    return rows


def fill_rates(rows: list, columns, measured) -> list:
    """[(libellé, nombre renseigné, total, taux)] pour chaque information mesurée."""
    labels = {key: label for key, label, _ in columns}
    total = len(rows)
    rates = []
    for key in measured:
        filled = sum(1 for row in rows if row[key] not in ("", None))
        rates.append((labels[key], filled, total, filled / total if total else None))
    return rates


def collect(ifc_file, level_of) -> dict:
    equipment = collect_equipment(ifc_file, level_of)
    spaces = collect_spaces(ifc_file, level_of)
    return {"equipment": equipment, "spaces": spaces,
            "equipment_rates": fill_rates(equipment, EQUIPMENT_COLUMNS, EQUIPMENT_MEASURED),
            "space_rates": fill_rates(spaces, SPACE_COLUMNS, SPACE_MEASURED)}


def write_inventory(path: str, inventory: dict, model: dict, marking: str = "", date=None) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", start_color="BDD7EE")
    workbook = Workbook()

    def setup(ws):
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        if marking:
            for part in (ws.oddHeader.center, ws.oddFooter.center):
                part.text, part.size, part.color = marking, 11, "C00000"
        ws.oddFooter.right.text = "Page &P / &N"

    def header(ws, row, labels):
        for col, label in enumerate(labels, 1):
            cell = ws.cell(row, col, label)
            cell.font, cell.fill = Font(bold=True), header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    summary = workbook.active
    summary.title = "Synthèse"
    setup(summary)
    for col, width in enumerate((34, 16, 12, 22), 1):
        summary.column_dimensions[get_column_letter(col)].width = width
    row = 1
    if marking:
        summary.cell(row, 1, marking).font = Font(bold=True, color="C00000")
        row += 1
    summary.cell(row, 1, "INVENTAIRE D'EXPLOITATION").font = Font(bold=True, size=14)
    row += 1
    file_info = model.get("file") or {}
    for label, value in (("Projet", model.get("project")), ("Maquette", file_info.get("name")),
                         ("Empreinte SHA-256", file_info.get("sha256")), ("Schéma IFC", model.get("schema")),
                         ("Date", date.strftime("%d/%m/%Y %H:%M") if date else "")):
        summary.cell(row, 1, label).font = Font(bold=True)
        summary.cell(row, 2, excel_safe(value or "—"))
        row += 1
    for title, rates, count in (("Équipements", inventory["equipment_rates"], len(inventory["equipment"])),
                                ("Locaux", inventory["space_rates"], len(inventory["spaces"]))):
        row += 1
        summary.cell(row, 1, f"{title} : {count}").font = Font(bold=True, size=12)
        row += 1
        header(summary, row, ["Information", "Renseignée", "Total", "Taux de renseignement"])
        row += 1
        for label, filled, total, rate in rates:
            summary.cell(row, 1, label)
            summary.cell(row, 2, filled)
            summary.cell(row, 3, total)
            cell = summary.cell(row, 4, "—" if rate is None else rate)
            cell.number_format = "0%"
            row += 1

    for title, columns, rows in (("Équipements", EQUIPMENT_COLUMNS, inventory["equipment"]),
                                 ("Locaux", SPACE_COLUMNS, inventory["spaces"])):
        ws = workbook.create_sheet(title)
        setup(ws)
        header(ws, 1, [label for _, label, _ in columns])
        for col, (_, _, width) in enumerate(columns, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        for r, item in enumerate(rows, 2):
            for col, (key, _, _) in enumerate(columns, 1):
                ws.cell(r, col, excel_safe(item[key]))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(len(rows) + 1, 1)}"

    neutralize_formulas(workbook)
    workbook.save(path)
