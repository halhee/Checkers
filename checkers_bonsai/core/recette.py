"""Recette d'une maquette IFC : contrôle IDS et/ou règles Excel, verdict.

Indépendant de Blender : prend un modèle `ifcopenshell.file` déjà ouvert
(celui de Bonsai) et des chemins de fichiers.
"""
import os
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

from . import ids as ids_check
from . import rules as rules_file
from .psets import check_psets
from .spatial import LevelCache
from .trace import file_info, model_info

VALIDATED = "Validé"
RESERVED = "Validé avec réserves"
REJECTED = "Rejeté"
INCONCLUSIVE = "Non concluant"
_SEVERITY = {VALIDATED: 0, RESERVED: 1, REJECTED: 2}


@dataclass
class RecetteSettings:
    rules_path: str = ""
    corps_etat: Optional[str] = None
    ids_path: str = ""
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
    pset: Optional[dict] = None
    ids_result: Optional[dict] = None
    anomalies: list = field(default_factory=list)
    valid_guids: set = field(default_factory=set)
    invalid_guids: set = field(default_factory=set)
    verdict: str = INCONCLUSIVE
    pset_verdict: Optional[str] = None
    ids_verdict: Optional[str] = None
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


def run_recette(ifc_file, settings: RecetteSettings, model_path: str = "") -> RecetteResult:
    if ifc_file is None:
        raise ValueError("Aucun modèle IFC ouvert dans Bonsai.")
    if not settings.rules_path and not settings.ids_path:
        raise ValueError("Choisissez au moins un référentiel : fichier IDS et/ou règles Excel.")
    start = time.time()
    result = RecetteResult(settings=settings, date=datetime.now(), model=model_info(ifc_file, model_path))
    level_of = LevelCache()
    anomalies = {}

    if settings.ids_path:
        document = ids_check.load_ids(settings.ids_path)
        result.ids = {**file_info(settings.ids_path), "title": ids_check.ids_title(document)}
        result.ids_result = ids_check.check_ids(document, ifc_file, level_of)
        result.ids_verdict = ids_verdict(result.ids_result, settings)
        anomalies.update({key: ("IDS", guids) for key, guids in result.ids_result["anomalies"].items()})
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
            anomalies[key] = ("PSet", guids)

    result.anomalies = sorted(
        ({"level": level, "ifc_class": ifc_class, "requirement": requirement, "source": source,
          "count": len(guids), "guids": sorted(guids)}
         for (level, ifc_class, requirement), (source, guids) in anomalies.items()),
        key=lambda a: (-a["count"], a["level"], a["ifc_class"], a["requirement"]))

    for anomaly in result.anomalies:
        result.invalid_guids.update(anomaly["guids"])
    if result.pset:
        result.invalid_guids.update(result.pset["invalid_guids"])
        result.valid_guids = set(result.pset["valid_guids"]) - result.invalid_guids

    verdicts = [v for v in (result.pset_verdict, result.ids_verdict) if v]
    result.verdict = max(verdicts, key=_SEVERITY.get) if verdicts else INCONCLUSIVE
    result.duration_s = time.time() - start
    return result


def describe_referentials(result: RecetteResult) -> str:
    parts = []
    if result.ids:
        title = f" – {result.ids['title']}" if result.ids.get("title") else ""
        parts.append(f"IDS « {result.ids.get('name', os.path.basename(result.settings.ids_path))} »{title}")
    if result.rules:
        parts.append(f"Règles Excel « {result.rules.get('name', '')} » ({result.rules['corps_etat']})")
    return " + ".join(parts)
