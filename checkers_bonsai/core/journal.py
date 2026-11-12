"""Journal de recette : enregistrement d'une recette dans un fichier JSON et
comparaison avec la recette précédente.

Le journal accompagne la fiche de recette. À la livraison suivante, il
permet de dire, élément par élément, quelles anomalies ont été corrigées,
lesquelles persistent et lesquelles sont apparues : c'est le suivi des
reprises entre deux livraisons d'une même maquette.

Un journal relu est traité comme une entrée non fiable : taille limitée,
structure et types vérifiés. Bibliothèque standard uniquement.
"""
import json
import os
from datetime import datetime

FORMAT = "checkers-recette"
VERSION = 1
MAX_JOURNAL_BYTES = 100 * 1024 * 1024
# Les anomalies qui ne portent sur aucun élément précis (ex. unité de
# longueur absente) sont suivies avec cet identifiant vide.
NO_ELEMENT = ""


class JournalError(ValueError):
    """Journal de recette illisible ou d'un format inattendu."""


def to_record(result) -> dict:
    """Recette sous forme de dictionnaire sérialisable en JSON."""
    s = result.settings
    model = result.model
    return {
        "format": FORMAT, "version": VERSION,
        "date": result.date.isoformat(timespec="seconds"),
        "lot": s.lot, "phase": s.phase, "tester": s.tester, "marking": s.marking,
        "verdict": result.verdict,
        "verdicts": {"ids": result.ids_verdict, "pset": result.pset_verdict, "structure": result.structure_verdict},
        "rates": {"ids": result.ids_rate, "pset": result.pset_rate},
        "thresholds": {"ok": s.threshold_ok, "reserve": s.threshold_reserve},
        "model": {
            "project": model.get("project", ""), "project_guid": model.get("project_guid", ""),
            "schema": model.get("schema", ""),
            "file": {k: (model.get("file") or {}).get(k, "") for k in ("name", "sha256")},
        },
        "referentials": {
            "ids": {k: result.ids.get(k, "") for k in ("name", "sha256", "title")} if result.ids else None,
            "rules": {k: result.rules.get(k, "") for k in ("name", "sha256", "corps_etat")} if result.rules else None,
        },
        "anomalies": [
            {"level": a["level"], "ifc_class": a["ifc_class"], "requirement": a["requirement"],
             "source": a["source"], "severity": a.get("severity") or "", "count": a["count"], "guids": a["guids"]}
            for a in result.anomalies
        ],
    }


def write_record(result, path: str) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(to_record(result), fh, ensure_ascii=False, indent=1)


def _require(condition, message):
    if not condition:
        raise JournalError(f"Journal de recette invalide : {message}")


def load_record(path: str) -> dict:
    if not os.path.isfile(path):
        raise JournalError(f"Journal de recette introuvable : {path}")
    if os.path.getsize(path) > MAX_JOURNAL_BYTES:
        raise JournalError("Journal de recette trop volumineux")
    try:
        with open(path, encoding="utf-8") as fh:
            record = json.load(fh)
    except (OSError, ValueError) as exc:
        raise JournalError(f"Journal de recette illisible ({type(exc).__name__})")

    _require(isinstance(record, dict) and record.get("format") == FORMAT, "ce n'est pas un journal Checkers")
    _require(record.get("version") == VERSION, f"version {record.get('version')!r} non prise en charge")
    _require(isinstance(record.get("date"), str), "date absente")
    try:
        datetime.fromisoformat(record["date"])
    except ValueError:
        raise JournalError("Journal de recette invalide : date illisible")
    _require(isinstance(record.get("anomalies"), list), "liste d'anomalies absente")
    for anomaly in record["anomalies"]:
        _require(isinstance(anomaly, dict), "anomalie mal formée")
        for key in ("level", "ifc_class", "requirement", "source"):
            _require(isinstance(anomaly.get(key), str), f"champ « {key} » absent d'une anomalie")
        _require(isinstance(anomaly.get("guids"), list) and all(isinstance(g, str) for g in anomaly["guids"]),
                 "identifiants d'éléments mal formés")
    for key in ("model", "referentials"):
        _require(isinstance(record.get(key), dict), f"bloc « {key} » absent")
    return record


def _pairs(anomalies) -> dict:
    """{(source, exigence, GlobalId): (niveau, classe IFC)} : une anomalie
    est suivie par élément, pas par groupe, pour qu'un élément déplacé d'un
    étage à l'autre ne passe pas pour « corrigé puis réapparu »."""
    pairs = {}
    for anomaly in anomalies:
        for guid in anomaly["guids"] or [NO_ELEMENT]:
            pairs[(anomaly["source"], anomaly["requirement"], guid)] = (anomaly["level"], anomaly["ifc_class"])
    return pairs


def _groups(pairs: dict, keys, removed=()) -> list:
    groups = {}
    for key in keys:
        source, requirement, guid = key
        level, ifc_class = pairs[key]
        group = groups.setdefault((level, ifc_class, requirement, source), {
            "level": level, "ifc_class": ifc_class, "requirement": requirement, "source": source,
            "guids": [], "count": 0, "removed": 0})
        group["count"] += 1
        if guid != NO_ELEMENT:
            group["guids"].append(guid)
        if key in removed:
            group["removed"] += 1
    return sorted(groups.values(), key=lambda g: (-g["count"], g["level"], g["ifc_class"], g["requirement"]))


def compare(previous: dict, result, element_exists) -> dict:
    """Compare la recette courante au journal précédent.

    `element_exists(GlobalId)` dit si l'élément est encore dans la maquette :
    une anomalie qui disparaît parce que l'élément a été supprimé est comptée
    à part, elle n'a pas été corrigée à proprement parler.

    Une exigence qui n'est plus contrôlée (référentiel différent) ferait
    passer ses anomalies pour résolues : la comparaison le signale dans
    « warnings », à l'utilisateur d'en tenir compte."""
    before = _pairs(previous["anomalies"])
    after = _pairs(result.anomalies)
    resolved_keys = set(before) - set(after)
    removed = {key for key in resolved_keys if key[2] != NO_ELEMENT and not element_exists(key[2])}

    warnings = []
    model, previous_model = result.model, previous["model"]
    if previous_model.get("project_guid") and model.get("project_guid") \
            and previous_model["project_guid"] != model["project_guid"]:
        warnings.append("La recette précédente porte sur un autre projet IFC.")
    if previous.get("lot", "") != result.settings.lot:
        warnings.append(f"La recette précédente porte sur un autre lot (« {previous.get('lot', '')} »).")
    current_hash = (model.get("file") or {}).get("sha256")
    if current_hash and current_hash == (previous_model.get("file") or {}).get("sha256"):
        warnings.append("La maquette est identique à celle de la recette précédente (même empreinte).")
    previous_refs, refs = previous["referentials"], {"ids": result.ids, "rules": result.rules}
    for key, label in (("ids", "IDS"), ("rules", "fichier de règles Excel")):
        old_hash = (previous_refs.get(key) or {}).get("sha256", "")
        new_hash = (refs[key] or {}).get("sha256", "")
        if old_hash != new_hash:
            warnings.append(f"Le {label} a changé depuis la recette précédente : les anomalies « résolues » ou "
                            f"« nouvelles » peuvent venir du référentiel et non de la maquette.")

    return {
        "previous_date": datetime.fromisoformat(previous["date"]),
        "previous_verdict": previous.get("verdict", ""),
        "previous_file": (previous_model.get("file") or {}).get("name", ""),
        "resolved": _groups(before, resolved_keys, removed),
        "new": _groups(after, set(after) - set(before)),
        "persistent": _groups(after, set(after) & set(before)),
        "resolved_count": len(resolved_keys), "removed_count": len(removed),
        "new_count": len(set(after) - set(before)), "persistent_count": len(set(after) & set(before)),
        "warnings": warnings,
    }
