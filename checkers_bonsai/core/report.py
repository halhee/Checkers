"""Rapport Excel de recette : fiche de recette (à signer), puis le détail
des contrôles de structure, de l'IDS, des PSet, du relevé terrain, de la
coordination, des dérogations et de l'évolution depuis la recette précédente.

La mention de confidentialité éventuelle (ex. « CONFIDENTIEL ») figure
en tête de la fiche et dans l'en-tête et le pied de page de chaque onglet
imprimé.
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import ids as ids_check
from .excel_safety import excel_safe, header_footer_text, neutralize_formulas
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
STATUS_LABELS = {"ok": "Conforme", "ko": "Non conforme", "erreur": "Non contrôlé"}
STATUS_FILLS = {"ok": FILL[VALIDATED], "ko": FILL[REJECTED], "erreur": FILL[INCONCLUSIVE]}
WRAP_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
WRAP_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
NOTE = Font(size=9, italic=True, color="666666")
PCT = {3: "0.0%"}


class _Sheet:
    """Petit utilitaire d'écriture ligne à ligne sur un nombre de colonnes fixe."""

    def __init__(self, workbook, title, widths, marking, first=False):
        self.ws = workbook.active if first else workbook.create_sheet()
        self.ws.title = title
        self.cols = len(widths)
        self.row = 1
        self.sections = 0
        for i, width in enumerate(widths, 1):
            self.ws.column_dimensions[get_column_letter(i)].width = width
        setup = self.ws.page_setup
        setup.orientation = "landscape"
        setup.fitToWidth, setup.fitToHeight = 1, 0
        self.ws.sheet_properties.pageSetUpPr.fitToPage = True
        if marking:
            for part in (self.ws.oddHeader.center, self.ws.oddFooter.center):
                part.text, part.size, part.color = header_footer_text(marking), 11, "C00000"
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

    def section(self, text, numbered=True):
        self.row += 1
        if numbered:
            self.sections += 1
            text = f"{self.sections}. {text}"
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


def _groups(result, source) -> int:
    return sum(1 for a in result.anomalies if a["source"] == source)


def _source_label(anomaly) -> str:
    return f"{anomaly['source']} – {anomaly['severity']}" if anomaly.get("severity") else anomaly["source"]


def _fiche(workbook, result: RecetteResult):
    s = result.settings
    sheet = _Sheet(workbook, "Fiche de recette", [26, 22, 30, 34, 18], s.marking, first=True)
    if s.marking:
        sheet.banner(s.marking, Font(size=12, bold=True, color="FFFFFF"), FILL["marking"], align=WRAP_CENTER)
    sheet.banner("FICHE DE RECETTE DE MAQUETTE NUMÉRIQUE (IFC)", Font(size=14, bold=True, color="FFFFFF"),
                 FILL["title"], height=28, align=WRAP_CENTER)

    sheet.section("Identification")
    model = result.model
    sheet.field("Projet", model.get("project"))
    sheet.field("Lot / corps d'état", s.lot)
    sheet.field("Phase / jalon", s.phase)
    sheet.field("Date de la recette", result.date.strftime("%d/%m/%Y %H:%M"))
    sheet.field("Recette réalisée par", s.tester)
    sheet.field("Contrôles réalisés", describe_referentials(result))

    sheet.section("Résultat")
    sheet.header(["Contrôle", "Vérifications", "Taux de conformité", "Anomalies (groupes)", "Verdict"])
    if result.structure is not None:
        checks = result.structure["checks"]
        done = [c for c in checks if c["status"] != "erreur"]
        rate = sum(1 for c in done if c["status"] == "ok") / len(done) if done else None
        sheet.line(["Structure du modèle", len(done), _pct(rate), _groups(result, "Structure"),
                    result.structure_verdict], fills={5: FILL[result.structure_verdict]}, formats=PCT)
    if result.ids_result is not None:
        r = result.ids_result
        sheet.line(["IDS", r["applicable_total"], _pct(result.ids_rate), _groups(result, "IDS"),
                    result.ids_verdict or INCONCLUSIVE],
                   fills={5: FILL[result.ids_verdict or INCONCLUSIVE]}, formats=PCT)
    if result.pset is not None:
        sheet.line(["Règles Excel (PSet)", result.pset["checked"], _pct(result.pset_rate), _groups(result, "PSet"),
                    result.pset_verdict or INCONCLUSIVE],
                   fills={5: FILL[result.pset_verdict or INCONCLUSIVE]}, formats=PCT)
    if result.survey_result is not None:
        r = result.survey_result
        sheet.line(["Relevé terrain", r["checks"], _pct(result.survey_rate), _groups(result, "Terrain"),
                    result.survey_verdict or INCONCLUSIVE],
                   fills={5: FILL[result.survey_verdict or INCONCLUSIVE]}, formats=PCT)
    if result.federation is not None:
        checks = result.federation["checks"]
        done = [c for c in checks if c["status"] != "erreur"]
        rate = sum(1 for c in done if c["status"] == "ok") / len(done) if done else None
        sheet.line([f"Cohérence entre lots ({len(result.federation['lots'])} lot(s))", len(done), _pct(rate),
                    _groups(result, "Cohérence"), result.federation_verdict],
                   fills={5: FILL[result.federation_verdict]}, formats=PCT)
    if result.clashes is not None:
        rules = result.clashes["rules"]
        done = [r for r in rules if r["status"] != "erreur"]
        rate = sum(1 for r in done if r["status"] == "ok") / len(done) if done else None
        sheet.line([f"Conflits géométriques ({result.clashes['total']} conflit(s))", len(done), _pct(rate),
                    _groups(result, "Conflit"), result.clash_verdict],
                   fills={5: FILL[result.clash_verdict]}, formats=PCT)
    if result.admitted:
        sheet.line([f"Dérogations ({result.admitted_elements} élément(s) ou constat(s))", "", "",
                    len(result.admitted), "Admises"], fills={5: FILL[INCONCLUSIVE]})
    sheet.row += 1
    sheet.banner(f"VERDICT : {result.verdict.upper()}", Font(size=13, bold=True), FILL[result.verdict],
                 height=24, align=WRAP_CENTER)
    sheet.banner("Vérifications : nombre de contrôles pour la structure et la cohérence entre lots, de couples "
                 "élément × spécification pour l'IDS, d'éléments pour les règles Excel, d'équipements vérifiés sur "
                 "site pour le relevé terrain, de règles pour les conflits géométriques.", NOTE, height=38)
    sheet.banner(f"Seuils IDS, règles Excel et relevé terrain : ≥ {s.threshold_ok:.0%} validé ; "
                 f"≥ {s.threshold_reserve:.0%} validé avec réserves ; rejeté en dessous. Une spécification IDS "
                 f"obligatoire sans aucun élément concerné entraîne le rejet. Structure, cohérence entre lots et "
                 f"conflits : un constat bloquant entraîne le rejet, un constat majeur ou un contrôle non exécuté "
                 f"des réserves. Le verdict global est le plus sévère. Les anomalies admises par dérogation ne "
                 f"comptent pas (taux et verdicts calculés sans elles ; détail dans l'onglet « Dérogations »).",
                 NOTE, height=62)
    for warning in result.warnings:
        sheet.banner(f"⚠ {warning}", Font(size=10, color="9C5700"))

    if result.evolution is not None:
        e = result.evolution
        sheet.section(f"Évolution depuis la recette du {e['previous_date']:%d/%m/%Y}")
        sheet.field("Recette précédente", " – ".join(v for v in (e["previous_file"], e["previous_verdict"]) if v))
        removed = f" (dont {e['removed_count']} par suppression de l'élément)" if e["removed_count"] else ""
        sheet.field("Anomalies levées", f"{e['resolved_count']}{removed}")
        sheet.field("Anomalies persistantes", str(e["persistent_count"]))
        sheet.field("Anomalies nouvelles", str(e["new_count"]))
        if e["not_rechecked_count"]:
            sheet.field("Anomalies non recontrôlées", f"{e['not_rechecked_count']} (contrôle non relancé)")
        if e.get("admitted_count"):
            sheet.field("Anomalies admises depuis", f"{e['admitted_count']} (dérogation accordée)")
        sheet.banner("Décompte par élément et par exigence ; détail dans l'onglet « Évolution ».", NOTE)

    sheet.section(f"Anomalies ouvertes ({len(result.anomalies)} groupes, "
                  f"{len(result.invalid_guids)} éléments concernés)")
    sheet.header(["Niveau / partie", "Classe IFC", "Exigence non respectée", "Source", "Éléments"])
    for a in result.anomalies:
        sheet.line([a["level"], a["ifc_class"], a["requirement"], _source_label(a), a["count"]])
    if not result.anomalies:
        sheet.banner("Aucune anomalie.")

    sheet.section("Traçabilité")
    f = model.get("file") or {}
    sheet.field("Maquette", f.get("name") or "Projet Bonsai non enregistré sur disque")
    sheet.field("Empreinte SHA-256", f.get("sha256"))
    unsaved = {True: "Oui : l'empreinte ne décrit pas exactement le modèle contrôlé", False: "Non"}
    sheet.field("Modifications non enregistrées", unsaved.get(result.unsaved_changes, "Non déterminé"))
    sheet.field("Schéma IFC", model.get("schema"))
    sheet.field("Application auteur", model.get("originating_system"))
    sheet.field("Auteur / organisation", " / ".join(v for v in (model.get("author"), model.get("organization")) if v))
    sheet.field("Date d'export IFC", model.get("timestamp"))
    previous_journal = (result.evolution or {}).get("journal")
    for label, info in (("IDS", result.ids), ("Règles Excel", result.rules), ("Relevé terrain", result.survey),
                        ("Dérogations", result.derogations), ("Journal précédent", previous_journal)):
        if info:
            sheet.field(f"{label} (SHA-256)", f"{info.get('name', '')} – {info.get('sha256', '')}")
    for lot in result.lots:
        sheet.field(f"Lot « {lot['name']} » (SHA-256)", f"{lot['file_name']} – {lot['sha256']}")
    sheet.banner("L'analyse porte sur le modèle ouvert dans Bonsai au moment de la recette ; l'empreinte est "
                 "celle du fichier enregistré sur disque. Enregistrer la maquette avant la recette pour qu'ils "
                 "coïncident.", NOTE)

    sheet.section("Validation")
    sheet.banner("Décision :  ☐ Validé     ☐ Validé avec réserves     ☐ Rejeté", height=22)
    sheet.banner("Réserves / observations :", height=60)
    sheet.banner("Nom, fonction :                                        Date :                    Visa :", height=36)
    if s.marking:
        sheet.row += 1
        sheet.banner(s.marking, Font(size=12, bold=True, color="FFFFFF"), FILL["marking"], align=WRAP_CENTER)


def _structure_sheet(workbook, result: RecetteResult):
    sheet = _Sheet(workbook, "Structure", [8, 44, 14, 18, 70], result.settings.marking)
    sheet.header(["Code", "Contrôle", "Gravité", "Résultat", "Constat"])
    for check in result.structure["checks"]:
        sheet.line([check["code"], check["title"], check["severity"], STATUS_LABELS[check["status"]],
                    check["message"]], fills={4: STATUS_FILLS[check["status"]]})
    sheet.ws.freeze_panes = "A2"


def _pset_details(workbook, result: RecetteResult):
    pset = result.pset
    sheet = _Sheet(workbook, "Détail PSet", [20, 22, 24, 26, 14, 28, 24, 18, 34], result.settings.marking)
    sheet.header(["Niveau / partie", "Classe IFC", "GlobalId", "Nom", "Repère (Tag)", "PSet", "Paramètre",
                  "Valeur", "Problème"])
    for f in pset["findings"]:
        sheet.line([f["level"], f["ifc_class"], f["guid"], f["name"], f["tag"], f["pset"], f["param"],
                    "" if f["value"] in (None, "") else str(f["value"]), f["problem"]])
    if not pset["findings"]:
        sheet.banner("Aucune non-conformité PSet.")
    if pset["truncated"]:
        sheet.banner("Liste tronquée (les compteurs de la fiche de recette restent exacts).", NOTE)
    sheet.ws.freeze_panes = "A2"
    sheet.ws.auto_filter.ref = f"A1:I{max(sheet.row - 1, 1)}"


def _ids_sheet(workbook, result: RecetteResult):
    r = result.ids_result
    sheet = _Sheet(workbook, "IDS", [34, 34, 16, 18, 24, 24, 40, 40], result.settings.marking)
    sheet.banner(f"IDS : {result.ids.get('name', '')}" + (f" – {result.ids['title']}" if result.ids.get("title") else ""),
                 Font(bold=True))
    sheet.section("Synthèse par spécification", numbered=False)
    sheet.header(["Spécification", "Identifiant", "Cardinalité", "Statut", "Applicables", "Non conformes",
                  "Conformes", "Exigences non respectées"])
    status_fill = {"conforme": FILL[VALIDATED], "non_conforme": FILL[REJECTED], "ignoree": FILL[INCONCLUSIVE]}
    for spec in r["specifications"]:
        sheet.line([spec["name"], spec["identifier"], ids_check.CARDINALITY_LABELS[spec["cardinality"]],
                    ids_check.STATUS_LABELS[spec["status"]], spec["applicable"], spec["failed"], spec["passed"],
                    spec["failed_requirements"]], fills={4: status_fill[spec["status"]]})
    sheet.section("Détail des éléments non conformes", numbered=False)
    sheet.header(["Spécification", "Exigence", "Niveau / partie", "Classe IFC", "GlobalId", "Nom", "Raison"])
    for d in r["details"]:
        sheet.line([d["specification"], d["requirement"], d["level"], d["ifc_class"], d["guid"], d["name"],
                    d["reason"]])
    if not r["details"]:
        sheet.banner("Aucun élément non conforme.")
    if r["truncated"]:
        sheet.banner("Liste tronquée (les compteurs restent exacts).", NOTE)


def _evolution_sheet(workbook, result: RecetteResult):
    e = result.evolution
    sheet = _Sheet(workbook, "Évolution", [16, 20, 22, 44, 14, 12, 18], result.settings.marking)
    sheet.header(["Statut", "Niveau / partie", "Classe IFC", "Exigence", "Source", "Éléments",
                  "Dont éléments supprimés"])
    rows = [("Levée", FILL[VALIDATED], e["resolved"]), ("Nouvelle", FILL[REJECTED], e["new"]),
            ("Persistante", FILL[RESERVED], e["persistent"]),
            ("Non recontrôlée", FILL[INCONCLUSIVE], e["not_rechecked"]),
            ("Admise", FILL[INCONCLUSIVE], e.get("admitted", []))]
    for label, fill, groups in rows:
        for g in groups:
            sheet.line([label, g["level"], g["ifc_class"], g["requirement"], g["source"], g["count"],
                        g["removed"] or ""], fills={1: fill})
    if not any(groups for _, _, groups in rows):
        sheet.banner("Aucune anomalie, ni à la recette précédente ni à celle-ci.")
    sheet.ws.freeze_panes = "A2"
    sheet.ws.auto_filter.ref = f"A1:G{max(sheet.row - 1, 1)}"


def _survey_sheet(workbook, result: RecetteResult):
    r = result.survey_result
    sheet = _Sheet(workbook, "Terrain", [18, 22, 22, 26, 28, 30, 22, 22, 36], result.settings.marking)
    sheet.banner(f"Relevé terrain : {result.survey.get('name', '')} — {r['verified']} équipement(s) vérifié(s) sur "
                 f"site, dont {r['conform']} sans écart ; {r['unmodelled']} non modélisé(s) ; {r['unverified']} non "
                 f"vérifié(s) ; {r['unknown']} ligne(s) sans élément correspondant dans la maquette.", Font(bold=True),
                 height=30)
    sheet.header(["Niveau / partie", "Local", "Classe IFC", "GlobalId", "Désignation", "Écart", "Valeur maquette",
                  "Valeur relevée", "Observation"])
    for d in r["details"]:
        sheet.line([d["level"], d["space"], d["ifc_class"], d["guid"], d["name"], d["problem"], d["model"],
                    d["found"], d["note"]])
    if not r["details"]:
        sheet.banner("Aucun écart entre la maquette et le relevé.")


def _federation_sheet(workbook, result: RecetteResult):
    f = result.federation
    sheet = _Sheet(workbook, "Cohérence", [24, 8, 30, 12, 14, 70], result.settings.marking)
    sheet.banner("Lots comparés à la maquette ouverte : " + "; ".join(
        f"« {lot['name']} » ({lot['schema']}, projet « {lot['project'] or '—'} »)" for lot in f["lots"]),
        Font(bold=True), height=30)
    sheet.header(["Lot", "Code", "Contrôle", "Gravité", "Statut", "Constat"])
    status = {"ok": ("Cohérent", FILL[VALIDATED]), "ko": ("Écart", FILL[REJECTED]),
              "erreur": ("Non exécuté", FILL[INCONCLUSIVE])}
    for c in f["checks"]:
        label, fill = status[c["status"]]
        sheet.line([c["lot"], c["code"], c["title"], c["severity"], label, c["message"]], fills={5: fill})
    sheet.banner("Tolérances : 1 cm en position et en altitude, 0,01° en orientation. Les entités spatiales (site, "
                 "bâtiment, étages) peuvent partager leur GlobalId entre lots ; les éléments, non.", NOTE)
    sheet.ws.freeze_panes = "A2"


def _clash_sheet(workbook, result: RecetteResult):
    c = result.clashes
    sheet = _Sheet(workbook, "Conflits", [30, 16, 20, 24, 26, 16, 20, 24, 26, 18, 10, 24],
                   result.settings.marking)
    sheet.section("Règles de conflits", numbered=False)
    sheet.header(["Règle", "Gravité", "Sélection (cette maquette)", "Sélection (autre lot)", "Résultat"])
    status_fill = {"ok": FILL[VALIDATED], "ko": FILL[REJECTED], "erreur": FILL[INCONCLUSIVE]}
    for rule in c["rules"]:
        sheet.line([rule["label"], rule["severity"], rule["selector_a"] or "tous les éléments",
                    rule["selector_b"] or "tous les éléments", rule["message"]], fills={5: status_fill[rule["status"]]})
    sheet.section("Détail des conflits", numbered=False)
    sheet.header(["Règle", "Niveau / partie", "Classe IFC", "GlobalId", "Nom", "Autre lot", "Classe IFC (autre)",
                  "GlobalId (autre)", "Nom (autre)", "Type", "Distance (m)", "Point (x ; y ; z)"])
    for d in c["details"]:
        sheet.line([d["rule"], d["level"], d["a_class"], d["a_guid"], d["a_name"], d["lot"], d["b_class"],
                    d["b_guid"], d["b_name"], d["type"], round(d["distance"], 4),
                    " ; ".join(f"{v:.3f}" for v in d["point"])])
    if not c["details"]:
        sheet.banner("Aucun conflit.")
    if c["truncated"]:
        sheet.banner("Liste tronquée (les compteurs restent exacts).", NOTE)


def _derogation_sheet(workbook, result: RecetteResult):
    info = result.derogations or {}
    sheet = _Sheet(workbook, "Dérogations", [18, 20, 22, 40, 10, 44, 20, 18, 14], result.settings.marking)
    sheet.banner(f"Fichier des dérogations : {info.get('name', '')} — {info.get('entries', 0)} dérogation(s) "
                 f"applicable(s) à ce lot.", Font(bold=True))
    sheet.section("Anomalies admises", numbered=False)
    sheet.header(["Source", "Niveau / partie", "Classe IFC", "Exigence", "Éléments", "Motif", "Accordée par",
                  "Le", "Identifiant"])
    for a in result.admitted:
        d = a["derogation"]
        sheet.line([_source_label(a), a["level"], a["ifc_class"], a["requirement"], a["count"], d.get("reason", ""),
                    d.get("author", ""), (d.get("date", "") or "")[:10], d.get("id", "")])
    if not result.admitted:
        sheet.banner("Aucune anomalie admise à cette recette.")
    unused = info.get("unused") or []
    sheet.section(f"Dérogations sans objet ({len(unused)})", numbered=False)
    sheet.banner("Plus aucune anomalie ne leur correspond : défaut corrigé, élément supprimé ou contrôle non "
                 "relancé. Elles peuvent être retirées du fichier.", NOTE)
    sheet.header(["Source", "Niveau / partie", "Classe IFC", "Exigence", "Éléments", "Motif", "Accordée par",
                  "Le", "Identifiant"])
    for e in unused:
        sheet.line([e["source"], e.get("level", ""), e.get("ifc_class", ""), e.get("label") or e["requirement"],
                    len(e.get("guids") or []) or "", e["reason"], e.get("author", ""), (e.get("date", "") or "")[:10],
                    e["id"]])


def write_report(result: RecetteResult, path: str) -> None:
    workbook = Workbook()
    _fiche(workbook, result)
    if result.evolution is not None:
        _evolution_sheet(workbook, result)
    if result.structure is not None:
        _structure_sheet(workbook, result)
    if result.ids_result is not None:
        _ids_sheet(workbook, result)
    if result.pset is not None:
        _pset_details(workbook, result)
    if result.survey_result is not None:
        _survey_sheet(workbook, result)
    if result.federation is not None:
        _federation_sheet(workbook, result)
    if result.clashes is not None:
        _clash_sheet(workbook, result)
    if result.derogations is not None:
        _derogation_sheet(workbook, result)
    neutralize_formulas(workbook)
    workbook.save(path)
