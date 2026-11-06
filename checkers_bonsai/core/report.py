"""Rapport Excel de recette : Fiche de recette (à signer), détail des
non-conformités PSet, synthèse et détail IDS.

La mention de confidentialité éventuelle (ex. « CONFIDENTIEL ») figure
en tête de la fiche et dans l'en-tête et le pied de page de chaque onglet
imprimé.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import ids as ids_check
from .excel_safety import excel_safe, neutralize_formulas
from .recette import INCONCLUSIVE, REJECTED, RESERVED, VALIDATED, RecetteResult, describe_referentials

_THIN = Side(style="thin")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
FILL = {
    "title": PatternFill("solid", start_color="1F3864"),
    "section": PatternFill("solid", start_color="D9E2F3"),
    "header": PatternFill("solid", start_color="BDD7EE"),
    "marking": PatternFill("solid", start_color="C00000"),
    VALIDATED: PatternFill("solid", start_color="C6EFCE"),
    RESERVED: PatternFill("solid", start_color="FFEB9C"),
    REJECTED: PatternFill("solid", start_color="FFC7CE"),
    INCONCLUSIVE: PatternFill("solid", start_color="EDEDED"),
}
WRAP_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
WRAP_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


class _Sheet:
    """Petit utilitaire d'écriture ligne à ligne sur un nombre de colonnes fixe."""

    def __init__(self, workbook, title, widths, marking, first=False):
        self.ws = workbook.active if first else workbook.create_sheet()
        self.ws.title = title
        self.cols = len(widths)
        self.row = 1
        for i, width in enumerate(widths, 1):
            self.ws.column_dimensions[get_column_letter(i)].width = width
        setup = self.ws.page_setup
        setup.orientation = "landscape"
        setup.fitToWidth, setup.fitToHeight = 1, 0
        self.ws.sheet_properties.pageSetUpPr.fitToPage = True
        if marking:
            for part in (self.ws.oddHeader.center, self.ws.oddFooter.center):
                part.text, part.size, part.color = marking, 11, "C00000"
        self.ws.oddFooter.right.text = "Page &P / &N"

    def banner(self, text, font=None, fill=None, height=None, align=WRAP_LEFT):
        self.ws.merge_cells(start_row=self.row, start_column=1, end_row=self.row, end_column=self.cols)
        cell = self.ws.cell(self.row, 1, excel_safe(text))
        cell.font = font or Font(size=11)
        cell.alignment = align
        if fill:
            cell.fill = fill
        if height:
            self.ws.row_dimensions[self.row].height = height
        self.row += 1

    def section(self, text):
        self.row += 1
        self.banner(text, Font(size=12, bold=True), FILL["section"])

    def header(self, labels):
        for col, label in enumerate(labels, 1):
            cell = self.ws.cell(self.row, col, label)
            cell.font, cell.fill, cell.alignment, cell.border = Font(bold=True), FILL["header"], WRAP_CENTER, BORDER
        self.row += 1

    def line(self, values, fills=None, formats=None):
        for col, value in enumerate(values, 1):
            cell = self.ws.cell(self.row, col, excel_safe(value))
            cell.border = BORDER
            cell.alignment = WRAP_LEFT if isinstance(value, str) else WRAP_CENTER
            if fills and fills.get(col):
                cell.fill = fills[col]
            if formats and formats.get(col):
                cell.number_format = formats[col]
        self.row += 1

    def field(self, label, value):
        """Ligne « libellé : valeur » (libellé en colonne A, valeur fusionnée)."""
        label_cell = self.ws.cell(self.row, 1, label)
        label_cell.font, label_cell.alignment = Font(bold=True), WRAP_LEFT
        self.ws.merge_cells(start_row=self.row, start_column=2, end_row=self.row, end_column=self.cols)
        self.ws.cell(self.row, 2, excel_safe(value if value not in (None, "") else "—")).alignment = WRAP_LEFT
        self.row += 1


def _pct(rate):
    """Taux en nombre (format pourcentage Excel, affiché selon la langue du
    poste : « 66,7 % » en français) ou tiret s'il n'y a rien à mesurer."""
    return "—" if rate is None else rate


PCT = {3: "0.0%"}


def _fiche(workbook, result: RecetteResult):
    s = result.settings
    sheet = _Sheet(workbook, "Fiche de recette", [26, 22, 30, 34, 16], s.marking, first=True)
    if s.marking:
        sheet.banner(s.marking, Font(size=12, bold=True, color="FFFFFF"), FILL["marking"], align=WRAP_CENTER)
    sheet.banner("FICHE DE RECETTE DE MAQUETTE NUMÉRIQUE (IFC)", Font(size=14, bold=True, color="FFFFFF"),
                 FILL["title"], height=28, align=WRAP_CENTER)

    sheet.section("1. Identification")
    model = result.model
    sheet.field("Projet", model.get("project"))
    sheet.field("Lot / corps d'état", s.lot)
    sheet.field("Phase / jalon", s.phase)
    sheet.field("Date de la recette", result.date.strftime("%d/%m/%Y %H:%M"))
    sheet.field("Recette réalisée par", s.tester)
    sheet.field("Référentiel(s)", describe_referentials(result))

    sheet.section("2. Résultat")
    sheet.header(["Contrôle", "Éléments contrôlés", "Taux de conformité", "Anomalies (groupes)", "Verdict"])
    if result.ids_result is not None:
        r = result.ids_result
        sheet.line(["IDS", r["applicable_total"], _pct(r["pass_rate"]),
                    sum(1 for a in result.anomalies if a["source"] == "IDS"), result.ids_verdict or INCONCLUSIVE],
                   fills={5: FILL[result.ids_verdict or INCONCLUSIVE]}, formats=PCT)
    if result.pset is not None:
        sheet.line(["Règles Excel (PSet)", result.pset["checked"], _pct(result.pset_rate),
                    sum(1 for a in result.anomalies if a["source"] == "PSet"), result.pset_verdict or INCONCLUSIVE],
                   fills={5: FILL[result.pset_verdict or INCONCLUSIVE]}, formats=PCT)
    sheet.row += 1
    sheet.banner(f"VERDICT : {result.verdict.upper()}", Font(size=13, bold=True), FILL[result.verdict],
                 height=24, align=WRAP_CENTER)
    sheet.banner(f"Seuils : ≥ {s.threshold_ok:.0%} validé ; ≥ {s.threshold_reserve:.0%} validé avec réserves ; "
                 f"rejeté en dessous. Une spécification IDS obligatoire sans aucun élément concerné entraîne le rejet.",
                 Font(size=9, italic=True, color="666666"))
    for warning in result.warnings:
        sheet.banner(f"⚠ {warning}", Font(size=10, color="9C5700"))

    sheet.section(f"3. Anomalies ouvertes ({len(result.anomalies)} groupes, "
                  f"{len(result.invalid_guids)} éléments non conformes)")
    sheet.header(["Niveau / partie", "Classe IFC", "Exigence non respectée", "Source", "Éléments"])
    for a in result.anomalies:
        sheet.line([a["level"], a["ifc_class"], a["requirement"], a["source"], a["count"]])
    if not result.anomalies:
        sheet.banner("Aucune anomalie.")

    sheet.section("4. Traçabilité")
    f = model.get("file") or {}
    sheet.field("Maquette", f.get("name") or "Projet Bonsai non enregistré sur disque")
    sheet.field("Empreinte SHA-256", f.get("sha256"))
    sheet.field("Schéma IFC", model.get("schema"))
    sheet.field("Application auteur", model.get("originating_system"))
    sheet.field("Auteur / organisation", " / ".join(v for v in (model.get("author"), model.get("organization")) if v))
    sheet.field("Date d'export IFC", model.get("timestamp"))
    for label, info in (("IDS", result.ids), ("Règles Excel", result.rules)):
        if info:
            sheet.field(f"{label} (SHA-256)", f"{info.get('name', '')} – {info.get('sha256', '')}")
    sheet.banner("L'analyse porte sur le modèle ouvert dans Bonsai au moment de la recette ; l'empreinte est "
                 "celle du fichier enregistré sur disque. Enregistrer la maquette avant la recette pour qu'ils "
                 "coïncident.", Font(size=9, italic=True, color="666666"))

    sheet.section("5. Validation")
    sheet.banner("Décision :  ☐ Validé     ☐ Validé avec réserves     ☐ Rejeté", height=22)
    sheet.banner("Réserves / observations :", height=60)
    sheet.banner("Nom, fonction :                                        Date :                    Visa :", height=36)
    if s.marking:
        sheet.row += 1
        sheet.banner(s.marking, Font(size=12, bold=True, color="FFFFFF"), FILL["marking"], align=WRAP_CENTER)


def _pset_details(workbook, result: RecetteResult):
    pset = result.pset
    sheet = _Sheet(workbook, "Détail PSet", [20, 22, 24, 26, 14, 28, 24, 18, 30], result.settings.marking)
    sheet.header(["Niveau / partie", "Classe IFC", "GlobalId", "Nom", "Repère (Tag)", "PSet", "Paramètre",
                  "Valeur", "Problème"])
    for f in pset["findings"]:
        sheet.line([f["level"], f["ifc_class"], f["guid"], f["name"], f["tag"], f["pset"], f["param"],
                    "" if f["value"] in (None, "") else str(f["value"]), f["problem"]])
    if not pset["findings"]:
        sheet.banner("Aucune non-conformité PSet.")
    if pset["truncated"]:
        sheet.banner("Liste tronquée (les compteurs de la fiche de recette restent exacts).",
                     Font(italic=True, color="666666"))
    sheet.ws.freeze_panes = "A2"
    sheet.ws.auto_filter.ref = f"A1:I{max(sheet.row - 1, 1)}"


def _ids_sheet(workbook, result: RecetteResult):
    r = result.ids_result
    sheet = _Sheet(workbook, "IDS", [34, 34, 16, 18, 24, 24, 40, 40], result.settings.marking)
    sheet.banner(f"IDS : {result.ids.get('name', '')}" + (f" – {result.ids['title']}" if result.ids.get("title") else ""),
                 Font(bold=True))
    sheet.section("Synthèse par spécification")
    sheet.header(["Spécification", "Identifiant", "Cardinalité", "Statut", "Applicables", "Non conformes",
                  "Conformes", "Exigences non respectées"])
    status_fill = {"conforme": FILL[VALIDATED], "non_conforme": FILL[REJECTED], "ignoree": FILL[INCONCLUSIVE]}
    for spec in r["specifications"]:
        sheet.line([spec["name"], spec["identifier"], ids_check.CARDINALITY_LABELS[spec["cardinality"]],
                    ids_check.STATUS_LABELS[spec["status"]], spec["applicable"], spec["failed"], spec["passed"],
                    spec["failed_requirements"]], fills={4: status_fill[spec["status"]]})
    sheet.section("Détail des éléments non conformes")
    sheet.header(["Spécification", "Exigence", "Niveau / partie", "Classe IFC", "GlobalId", "Nom", "Raison"])
    for d in r["details"]:
        sheet.line([d["specification"], d["requirement"], d["level"], d["ifc_class"], d["guid"], d["name"],
                    d["reason"]])
    if not r["details"]:
        sheet.banner("Aucun élément non conforme.")
    if r["truncated"]:
        sheet.banner("Liste tronquée (les compteurs restent exacts).", Font(italic=True, color="666666"))


def write_report(result: RecetteResult, path: str) -> None:
    workbook = Workbook()
    _fiche(workbook, result)
    if result.ids_result is not None:
        _ids_sheet(workbook, result)
    if result.pset is not None:
        _pset_details(workbook, result)
    neutralize_formulas(workbook)
    workbook.save(path)
