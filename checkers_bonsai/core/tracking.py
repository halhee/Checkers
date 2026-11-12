"""Tableau de suivi : consolide les journaux de recette d'un dossier en un
classeur Excel, pour piloter un projet ou un ensemble de lots.

- onglet « Synthèse » : une ligne par maquette suivie (projet et lot), avec
  l'état de sa dernière recette et la tendance depuis la précédente ;
- onglet « Recettes » : toutes les recettes, de la plus ancienne à la plus
  récente.

Ne lit que des journaux (.json) produits par Checkers : aucune maquette n'est
ouverte, le tableau se construit en quelques secondes, sans modèle chargé.
"""
import os
from datetime import datetime

from . import journal
from .excel_safety import excel_safe, header_footer_text, neutralize_formulas

MAX_FILES = 5000
MAX_DEPTH = 4
_VERDICT_COLORS = {"Validé": "C6EFCE", "Validé avec réserves": "FFEB9C", "Rejeté": "FFC7CE"}


def collect_journals(folder: str) -> tuple:
    """(recettes, fichiers refusés). Parcourt `folder` et ses sous-dossiers
    (profondeur limitée). Un .json qui n'est pas un journal Checkers est
    ignoré ; un journal Checkers illisible est listé parmi les refusés."""
    if not os.path.isdir(folder):
        raise ValueError(f"Dossier introuvable : {folder}")
    records, rejected, seen = [], [], 0
    root_depth = os.path.abspath(folder).rstrip(os.sep).count(os.sep)
    for current, dirs, files in os.walk(folder):
        if os.path.abspath(current).count(os.sep) - root_depth >= MAX_DEPTH:
            dirs[:] = []
        dirs.sort()
        for name in sorted(files):
            if not name.lower().endswith(".json"):
                continue
            seen += 1
            if seen > MAX_FILES:
                raise ValueError(f"Trop de fichiers .json dans ce dossier (plus de {MAX_FILES})")
            path = os.path.join(current, name)
            try:
                record = journal.load_record(path)
            except journal.NotAJournalError:
                continue
            except journal.JournalError as exc:
                rejected.append((os.path.relpath(path, folder), str(exc)))
                continue
            record["_file"] = os.path.relpath(path, folder)
            record["_date"] = datetime.fromisoformat(record["date"]).replace(tzinfo=None)
            records.append(record)
    records.sort(key=lambda r: r["_date"])
    return records, rejected


def open_anomalies(record: dict) -> int:
    """Nombre d'anomalies ouvertes, comptées par élément et par exigence
    (comme l'évolution entre deux recettes)."""
    return sum(max(len(a["guids"]), 1) for a in record["anomalies"])


def _key(record: dict) -> tuple:
    model = record["model"]
    return (str(model.get("project_guid") or model.get("project") or ""), record.get("lot", ""))


def summarize(records: list) -> list:
    """Une ligne par (projet, lot) : dernière recette et tendance."""
    groups = {}
    for record in records:  # déjà triées par date
        groups.setdefault(_key(record), []).append(record)
    rows = []
    for history in groups.values():
        last = history[-1]
        previous = history[-2] if len(history) > 1 else None
        opened = open_anomalies(last)
        rows.append({
            "project": str(last["model"].get("project") or ""), "lot": last.get("lot", ""),
            "phase": last.get("phase", ""), "date": last["_date"], "verdict": last.get("verdict", ""),
            "open": opened, "delta": opened - open_anomalies(previous) if previous else None,
            "count": len(history), "first": history[0]["_date"],
            "file": str((last["model"].get("file") or {}).get("name") or ""),
        })
    rows.sort(key=lambda r: (r["project"], r["lot"]))
    return rows


def _rate(record: dict, key: str):
    value = (record.get("rates") or {}).get(key) if isinstance(record.get("rates"), dict) else None
    return value if isinstance(value, (int, float)) and not isinstance(value, bool) else None


def _verdict(record: dict, key: str) -> str:
    value = (record.get("verdicts") or {}).get(key) if isinstance(record.get("verdicts"), dict) else None
    return value if isinstance(value, str) else ""


def write_tracking(path: str, records: list, rejected=(), marking: str = "", date=None) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", start_color="BDD7EE")
    workbook = Workbook()

    def sheet(title, columns, first=False):
        ws = workbook.active if first else workbook.create_sheet()
        ws.title = title
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        if marking:
            for part in (ws.oddHeader.center, ws.oddFooter.center):
                part.text, part.size, part.color = header_footer_text(marking), 11, "C00000"
        ws.oddFooter.right.text = "Page &P / &N"
        for col, (label, width) in enumerate(columns, 1):
            cell = ws.cell(1, col, label)
            cell.font, cell.fill = Font(bold=True), header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(col)].width = width
        ws.freeze_panes = "A2"
        return ws

    def fill_row(ws, r, values, verdict_column, formats=None):
        for col, value in enumerate(values, 1):
            cell = ws.cell(r, col, excel_safe(value))
            if formats and col in formats:
                cell.number_format = formats[col]
        color = _VERDICT_COLORS.get(values[verdict_column - 1])
        if color:
            ws.cell(r, verdict_column).fill = PatternFill("solid", start_color=color)

    summary_columns = (("Projet", 30), ("Lot", 26), ("Phase", 22), ("Dernière recette", 16), ("Verdict", 22),
                       ("Anomalies ouvertes", 14), ("Évolution depuis la recette précédente", 20),
                       ("Nombre de recettes", 12), ("Première recette", 16), ("Maquette", 30))
    ws = sheet("Synthèse", summary_columns, first=True)
    rows = summarize(records)
    for r, row in enumerate(rows, 2):
        # Écart écrit comme un nombre, affiché avec son signe (« +3 », « -2 »).
        fill_row(ws, r, [row["project"], row["lot"], row["phase"], row["date"], row["verdict"], row["open"],
                         row["delta"], row["count"], row["first"], row["file"]], 5,
                 {4: "DD/MM/YYYY", 7: "+0;-0;0", 9: "DD/MM/YYYY"})
    last = len(rows) + 1
    ws.auto_filter.ref = f"A1:{get_column_letter(len(summary_columns))}{max(last, 1)}"
    note_row = last + 2
    notes = [f"Tableau produit le {date.strftime('%d/%m/%Y %H:%M') if date else '—'} à partir de "
             f"{len(records)} journal(aux) de recette.",
             "Anomalies ouvertes : décompte par élément et par exigence. Évolution : écart avec la recette "
             "précédente du même lot (négatif = anomalies en baisse)."]
    notes += [f"Journal refusé : {name} ({reason})" for name, reason in rejected]
    if marking:
        notes.insert(0, marking)
    for offset, text in enumerate(notes):
        ws.cell(note_row + offset, 1, excel_safe(text)).font = Font(italic=True, color="666666")

    recette_columns = (("Date", 16), ("Projet", 30), ("Lot", 26), ("Phase", 22), ("Maquette", 30), ("Verdict", 22),
                       ("Structure", 20), ("IDS", 20), ("Taux IDS", 10), ("Règles Excel", 20), ("Taux PSet", 10),
                       ("Terrain", 20), ("Taux terrain", 10), ("Anomalies ouvertes", 14), ("Réalisée par", 22),
                       ("Journal", 40))
    ws = sheet("Recettes", recette_columns)
    for r, record in enumerate(records, 2):
        model = record["model"]
        fill_row(ws, r, [
            record["_date"], str(model.get("project") or ""), record.get("lot", ""), record.get("phase", ""),
            str((model.get("file") or {}).get("name") or ""), record.get("verdict", ""),
            _verdict(record, "structure"), _verdict(record, "ids"), _rate(record, "ids"),
            _verdict(record, "pset"), _rate(record, "pset"), _verdict(record, "survey"), _rate(record, "survey"),
            open_anomalies(record), record.get("tester", "") if isinstance(record.get("tester", ""), str) else "",
            record["_file"]], 6, {1: "DD/MM/YYYY HH:MM", 9: "0.0%", 11: "0.0%", 13: "0.0%"})
    ws.auto_filter.ref = f"A1:{get_column_letter(len(recette_columns))}{max(len(records) + 1, 1)}"

    neutralize_formulas(workbook)
    workbook.save(path)
