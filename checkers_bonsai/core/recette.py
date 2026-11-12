"""Recette d'une maquette IFC : contrôles de structure, IDS, règles Excel
et relevé terrain, verdict, et comparaison avec la recette précédente.

Indépendant de Blender : prend un modèle `ifcopenshell.file` déjà ouvert
(celui de Bonsai) et des chemins de fichiers.
"""
import os
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

from . import ids as ids_check
from . import journal
from . import rules as rules_file
from . import structure as structure_check
from . import survey as survey_check
from .psets import check_psets
from .spatial import LevelCache
from .trace import file_info, model_info

VALIDATED = "Validé"
RESERVED = "Validé avec réserves"
REJECTED = "Rejeté"
INCONCLUSIVE = "Non concluant"
_SEVERITY = {VALIDATED: 0, RESERVED: 1, REJECTED: 2}
# Ordre d'affichage des anomalies : les constats bloquants d'abord, les
# constats mineurs en dernier ; les écarts IDS et PSet entre les deux.
_ANOMALY_RANK = {structure_check.BLOCKING: 0, structure_check.MAJOR: 1, None: 1, structure_check.MINOR: 2}


@dataclass
class RecetteSettings:
    rules_path: str = ""
    corps_etat: Optional[str] = None
    ids_path: str = ""
    check_structure: bool = True
    survey_path: str = ""            # fiche de relevé terrain remplie (.xlsx)
    previous_path: str = ""          # journal JSON de la recette précédente
    lot: str = ""
    phase: str = ""
    tester: str = ""
    marking: str = ""
    threshold_ok: float = 0.98
    threshold_reserve: float = 0.90


@dataclass
class RecetteResult:
    settings: RecetteSettings
    date: datetime
    model: dict
    rules: dict = field(default_factory=dict)        # fichier Excel : infos + corps d'état
    ids: dict = field(default_factory=dict)          # fichier IDS : infos + titre
    survey: dict = field(default_factory=dict)       # relevé terrain : infos du fichier
    pset: Optional[dict] = None
    ids_result: Optional[dict] = None
    structure: Optional[dict] = None
    survey_result: Optional[dict] = None
    evolution: Optional[dict] = None                 # comparaison avec la recette précédente
    anomalies: list = field(default_factory=list)
    valid_guids: set = field(default_factory=set)
    invalid_guids: set = field(default_factory=set)
    verdict: str = INCONCLUSIVE
    pset_verdict: Optional[str] = None
    ids_verdict: Optional[str] = None
    structure_verdict: Optional[str] = None
    survey_verdict: Optional[str] = None
    unsaved_changes: Optional[bool] = None
    warnings: list = field(default_factory=list)
    duration_s: float = 0.0

    @property
    def pset_rate(self) -> Optional[float]:
        if not self.pset or not self.pset["checked"]:
            return None
        return self.pset["valid"] / self.pset["checked"]

    @property
    def ids_rate(self) -> Optional[float]:
        return self.ids_result["pass_rate"] if self.ids_result else None

    @property
    def survey_rate(self) -> Optional[float]:
        return self.survey_result["pass_rate"] if self.survey_result else None

    @property
    def sources(self) -> set:
        """Sources d'anomalies contrôlées par cette recette."""
        done = (("Structure", self.structure), ("IDS", self.ids_result), ("PSet", self.pset),
                ("Terrain", self.survey_result))
        return {source for source, outcome in done if outcome is not None}


def rate_verdict(rate: float, settings: RecetteSettings) -> str:
    if rate >= settings.threshold_ok:
        return VALIDATED
    if rate >= settings.threshold_reserve:
        return RESERVED
    return REJECTED


def ids_verdict(result: dict, settings: RecetteSettings) -> Optional[str]:
    if not result or not result["specs_checked"]:
        return None
    if result["cardinality_failures"]:
        return REJECTED  # spécification obligatoire sans aucun élément concerné
    if result["pass_rate"] is None:
        return VALIDATED if result["specs_failed"] == 0 else REJECTED
    return rate_verdict(result["pass_rate"], settings)


def structure_verdict(result: dict) -> str:
    if result["blocking"]:
        return REJECTED
    # Un contrôle qui n'a pas pu s'exécuter ne prouve rien : pas de « validé ».
    if result["major"] or result.get("errors"):
        return RESERVED
    return VALIDATED


def _check_settings(settings: RecetteSettings) -> None:
    if not (settings.rules_path or settings.ids_path or settings.check_structure or settings.survey_path):
        raise ValueError("Aucun contrôle demandé : choisissez un IDS, des règles Excel, un relevé terrain ou le "
                         "contrôle de structure.")
    if not 0.0 <= settings.threshold_reserve <= settings.threshold_ok <= 1.0:
        raise ValueError("Seuils incohérents : le seuil « avec réserves » doit être inférieur ou égal au seuil "
                         "« validé », tous deux entre 0 et 100 %.")


def run_recette(ifc_file, settings: RecetteSettings, model_path: str = "",
                unsaved_changes: Optional[bool] = None) -> RecetteResult:
    if ifc_file is None:
        raise ValueError("Aucun modèle IFC ouvert dans Bonsai.")
    _check_settings(settings)
    # Lu avant les contrôles : un journal illisible doit arrêter tout de
    # suite, pas après plusieurs minutes de calcul.
    previous = journal.load_record(settings.previous_path) if settings.previous_path else None
    survey_rows = survey_check.load_survey(settings.survey_path) if settings.survey_path else None

    start = time.time()
    result = RecetteResult(settings=settings, date=datetime.now(), model=model_info(ifc_file, model_path),
                           unsaved_changes=unsaved_changes)
    level_of = LevelCache()
    anomalies = []

    def add(key, source, guids, severity=None, count=None):
        level, ifc_class, requirement = key
        anomalies.append({"level": level, "ifc_class": ifc_class, "requirement": requirement, "source": source,
                          "severity": severity, "count": len(guids) if count is None else count,
                          "guids": sorted(guids)})

    if settings.check_structure:
        result.structure = structure_check.check_structure(ifc_file, level_of)
        result.structure_verdict = structure_verdict(result.structure)
        for key, item in result.structure["anomalies"].items():
            add(key, "Structure", item["guids"], item["severity"], item["count"])
        if result.structure.get("errors"):
            result.warnings.append(f"{result.structure['errors']} contrôle(s) de structure n'ont pas pu être "
                                   f"exécutés (voir l'onglet « Structure ») : la structure ne peut pas être "
                                   f"validée sans réserve.")

    if settings.ids_path:
        document = ids_check.load_ids(settings.ids_path)
        result.ids = {**file_info(settings.ids_path), "title": ids_check.ids_title(document)}
        result.ids_result = ids_check.check_ids(document, ifc_file, level_of)
        result.ids_verdict = ids_verdict(result.ids_result, settings)
        result.warnings.extend(result.ids_result["warnings"])
        for key, guids in result.ids_result["anomalies"].items():
            add(key, "IDS", guids, count=max(len(guids), 1))  # 1 = constat sur le modèle entier
        if result.ids_result["specs_checked"] == 0:
            result.warnings.append(f"Aucune spécification de l'IDS ne vise le schéma {result.model['schema']}.")

    if settings.rules_path:
        registry = rules_file.load_registry(settings.rules_path)
        code = settings.corps_etat or None
        if registry and code and code not in registry:
            raise ValueError(f"Corps d'état « {code} » absent de l'onglet Corps_Etat")
        if registry and not code:
            result.warnings.append("Le fichier de règles déclare des corps d'état mais aucun n'a été choisi : "
                                   "toutes les règles sont appliquées.")
        rules = rules_file.load_rules(settings.rules_path, code)
        result.rules = {**file_info(settings.rules_path),
                        "corps_etat": f"{code} – {registry[code]}" if registry and code else "Toutes les règles"}
        result.pset = check_psets(ifc_file, rules, level_of)
        result.warnings.extend(result.pset["warnings"])
        if result.pset["checked"]:
            result.pset_verdict = rate_verdict(result.pset_rate, settings)
        else:
            result.warnings.append("Aucun élément du modèle n'est concerné par les règles Excel.")
        for key, guids in result.pset["anomalies"].items():
            add(key, "PSet", guids)

    if survey_rows is not None:
        result.survey = file_info(settings.survey_path)
        result.survey_result = survey_check.check_survey(ifc_file, survey_rows, level_of)
        result.warnings.extend(result.survey_result["warnings"])
        if result.survey_result["checks"]:
            result.survey_verdict = rate_verdict(result.survey_rate, settings)
        else:
            result.warnings.append("Relevé terrain : aucune vérification exploitable dans la fiche.")
        for key, guids in result.survey_result["anomalies"].items():
            add(key, "Terrain", guids)
        for key, count in result.survey_result["model_anomalies"].items():
            add(key, "Terrain", (), count=count)

    result.anomalies = sorted(anomalies, key=lambda a: (
        _ANOMALY_RANK[a["severity"]], -a["count"], a["level"], a["ifc_class"], a["requirement"]))

    # Rouge : au moins une anomalie. Vert : contrôlé par l'IDS, les règles
    # Excel ou le relevé terrain, sans aucune anomalie.
    for anomaly in result.anomalies:
        result.invalid_guids.update(anomaly["guids"])
    checked = set()
    if result.pset:
        result.invalid_guids.update(result.pset["invalid_guids"])
        checked.update(result.pset["valid_guids"])
    if result.ids_result:
        checked.update(result.ids_result["applicable_guids"])
    if result.survey_result:
        checked.update(result.survey_result["conform_guids"])
    result.valid_guids = checked - result.invalid_guids

    verdicts = [v for v in (result.structure_verdict, result.pset_verdict, result.ids_verdict,
                            result.survey_verdict) if v]
    result.verdict = max(verdicts, key=_SEVERITY.get) if verdicts else INCONCLUSIVE

    if unsaved_changes:
        result.warnings.append("Le modèle contient des modifications non enregistrées : l'empreinte du fichier "
                               "ne décrit pas exactement le modèle contrôlé.")
    if previous is not None:
        def element_exists(guid: str) -> bool:
            try:
                return ifc_file.by_guid(guid) is not None
            except Exception:
                return False

        result.evolution = journal.compare(previous, result, element_exists)
        result.evolution["journal"] = file_info(settings.previous_path)
        result.warnings.extend(result.evolution["warnings"])

    result.duration_s = time.time() - start
    return result


def describe_referentials(result: RecetteResult) -> str:
    parts = []
    if result.ids:
        title = f" – {result.ids['title']}" if result.ids.get("title") else ""
        parts.append(f"IDS « {result.ids.get('name', os.path.basename(result.settings.ids_path))} »{title}")
    if result.rules:
        parts.append(f"Règles Excel « {result.rules.get('name', '')} » ({result.rules['corps_etat']})")
    if result.survey:
        parts.append(f"Relevé terrain « {result.survey.get('name', '')} »")
    if result.structure is not None:
        parts.append("contrôles de structure")
    return " + ".join(parts)
