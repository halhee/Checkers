"""Dérogations : anomalies admises, avec leur motif.

Un conflit voulu (gaine dans sa réservation), une réserve acceptée par la
maîtrise d'ouvrage ou un écart sans conséquence reviennent à chaque recette.
Une dérogation les admet une fois pour toutes : elles sortent du verdict et
de la liste des anomalies, et figurent à part sur la fiche, avec le motif,
l'auteur et la date.

Une dérogation porte sur des éléments précis (GlobalId) pour une exigence
donnée : un nouvel élément présentant le même défaut n'est pas couvert. Une
anomalie sans élément (constat sur la maquette entière) est admise par
niveau et classe. Une dérogation peut être propre à un lot ou valoir pour
tous (« lot » vide).

Le fichier est un JSON à conserver avec les journaux de recette ; il est
relu comme une entrée non fiable (taille, structure et types vérifiés).
Bibliothèque standard uniquement.
"""
import json
import os
import tempfile
import uuid
from collections import defaultdict
from datetime import datetime

from .clashes import MAX_PAIRS_PER_GROUP, stable_requirement

FORMAT = "checkers-derogations"
VERSION = 1
MAX_BYTES = 20 * 1024 * 1024
MAX_TEXT = 2000
SUFFIX = ".derogations.json"


class DerogationError(ValueError):
    """Fichier de dérogations illisible ou dérogation invalide."""


def new_document() -> dict:
    return {"format": FORMAT, "version": VERSION, "derogations": []}


def default_path(model_path: str) -> str:
    """Fichier de dérogations proposé à côté de la maquette."""
    folder, name = os.path.split(model_path)
    return os.path.join(folder, os.path.splitext(name)[0] + SUFFIX)


def _require(condition, message):
    if not condition:
        raise DerogationError(f"Fichier de dérogations invalide : {message}")


def _text(entry, key, required=False):
    value = entry.get(key, "")
    _require(isinstance(value, str) and len(value) <= MAX_TEXT, f"champ « {key} » mal formé")
    _require(not required or value.strip(), f"champ « {key} » vide")


def load(path: str) -> dict:
    if not os.path.isfile(path):
        raise DerogationError(f"Fichier de dérogations introuvable : {path}")
    if os.path.getsize(path) > MAX_BYTES:
        raise DerogationError("Fichier de dérogations trop volumineux")
    try:
        with open(path, encoding="utf-8") as fh:
            document = json.load(fh)
    except (OSError, ValueError) as exc:
        raise DerogationError(f"Fichier de dérogations illisible ({type(exc).__name__})")
    _require(isinstance(document, dict) and document.get("format") == FORMAT, "ce n'est pas un fichier Checkers")
    _require(document.get("version") == VERSION, f"version {document.get('version')!r} non prise en charge")
    entries = document.get("derogations")
    _require(isinstance(entries, list), "liste des dérogations absente")
    ids = set()
    for entry in entries:
        _require(isinstance(entry, dict), "dérogation mal formée")
        for key in ("id", "source", "requirement", "reason"):
            _text(entry, key, required=True)
        for key in ("label", "level", "ifc_class", "lot", "author", "date"):
            _text(entry, key)
        guids = entry.get("guids", [])
        _require(isinstance(guids, list) and all(isinstance(g, str) and len(g) <= 64 for g in guids),
                 "identifiants d'éléments mal formés")
        _require(entry["id"] not in ids, f"identifiant « {entry['id']} » en double")
        ids.add(entry["id"])
    return document


def load_or_new(path: str) -> dict:
    return load(path) if os.path.isfile(path) else new_document()


def save(document: dict, path: str) -> None:
    """Écriture atomique : un fichier interrompu en cours d'écriture ne doit
    pas faire perdre les dérogations déjà accordées."""
    folder = os.path.dirname(os.path.abspath(path))
    handle, temporary = tempfile.mkstemp(prefix=".derogations-", suffix=".tmp", dir=folder)
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as fh:
            json.dump(document, fh, ensure_ascii=False, indent=1)
        os.replace(temporary, path)
    except BaseException:
        if os.path.exists(temporary):
            os.remove(temporary)
        raise


def make_entry(anomaly: dict, reason: str, author: str = "", lot: str = "", guids=None,
               date: datetime = None) -> dict:
    """Dérogation pour une anomalie de la recette, sur tous ses éléments ou
    sur `guids` seulement."""
    reason = (reason or "").strip()
    if not reason:
        raise DerogationError("Le motif de la dérogation est obligatoire.")
    if len(reason) > MAX_TEXT:
        raise DerogationError(f"Motif trop long ({MAX_TEXT} caractères au plus).")
    chosen = sorted(set(anomaly.get("guids") or ()) if guids is None else set(guids))
    return {
        "id": uuid.uuid4().hex[:12],
        "source": anomaly["source"],
        "requirement": stable_requirement(anomaly["source"], anomaly["requirement"]),
        "label": anomaly["requirement"],
        "level": anomaly["level"], "ifc_class": anomaly["ifc_class"],
        "guids": chosen, "lot": lot or "",
        "reason": reason, "author": (author or "").strip()[:MAX_TEXT],
        "date": (date or datetime.now()).isoformat(timespec="seconds"),
    }


def add(document: dict, entry: dict) -> None:
    document["derogations"].append(entry)


def revoke(document: dict, entry_id: str) -> dict:
    for index, entry in enumerate(document["derogations"]):
        if entry["id"] == entry_id:
            return document["derogations"].pop(index)
    raise DerogationError(f"Aucune dérogation « {entry_id} » dans le fichier.")


def _applicable(document: dict, lot: str) -> list:
    return [e for e in document.get("derogations", []) if not e.get("lot") or e.get("lot") == (lot or "")]


def apply(anomalies, document: dict, lot: str = ""):
    """Sépare les anomalies admises des autres.

    Retourne (restantes, admises, sans_objet) :
    - restantes : anomalies sans les éléments admis ;
    - admises : une entrée par anomalie et par dérogation, avec la clé
      « derogation » (id, reason, author, date) ;
    - sans_objet : dérogations du lot qui ne correspondent plus à aucune
      anomalie (défaut corrigé, élément supprimé, contrôle non relancé)."""
    by_guid, by_group = {}, {}
    entries = _applicable(document, lot) if document else []
    for entry in entries:
        if entry.get("guids"):
            for guid in entry["guids"]:
                by_guid.setdefault((entry["source"], entry["requirement"], guid), entry)
        else:
            by_group.setdefault((entry["source"], entry["requirement"], entry.get("level", ""),
                                 entry.get("ifc_class", "")), entry)

    remaining, admitted, used = [], [], set()
    for anomaly in anomalies:
        stable = stable_requirement(anomaly["source"], anomaly["requirement"])
        guids = list(anomaly.get("guids") or ())
        if not guids:
            entry = by_group.get((anomaly["source"], stable, anomaly["level"], anomaly["ifc_class"]))
            if entry is None:
                remaining.append(anomaly)
            else:
                used.add(entry["id"])
                admitted.append({**anomaly, "derogation": _summary(entry)})
            continue
        covered = defaultdict(list)
        rest = []
        for guid in guids:
            entry = by_guid.get((anomaly["source"], stable, guid))
            if entry is None:
                rest.append(guid)
            else:
                covered[entry["id"]].append((guid, entry))
        if not covered:
            remaining.append(anomaly)
            continue
        for entry_id, items in covered.items():
            used.add(entry_id)
            part = [g for g, _ in items]
            admitted.append(_restrict({**anomaly, "guids": part, "count": len(items),
                                       "derogation": _summary(items[0][1])}, part))
        if rest:
            count = anomaly["count"]
            count = len(rest) if count == len(guids) else max(count - (len(guids) - len(rest)), len(rest))
            remaining.append(_restrict({**anomaly, "guids": rest, "count": count}, rest))
    unused = [entry for entry in entries if entry["id"] not in used]
    return remaining, admitted, unused


def _restrict(anomaly: dict, guids) -> dict:
    """Conflits (paires a / b) limités aux éléments gardés dans cette partie."""
    if anomaly.get("conflicts"):
        kept = set(guids)
        complete = len(anomaly["conflicts"]) < MAX_PAIRS_PER_GROUP     # liste non tronquée
        anomaly["conflicts"] = [c for c in anomaly["conflicts"] if c["a"] in kept]
        if anomaly["conflicts"]:
            anomaly["point"] = anomaly["conflicts"][0]["point"]
            if complete and anomaly.get("related_guids"):
                related = {c["b"] for c in anomaly["conflicts"]}
                anomaly["related_guids"] = [g for g in anomaly["related_guids"] if g in related]
    return anomaly


def _summary(entry: dict) -> dict:
    return {key: entry.get(key, "") for key in ("id", "reason", "author", "date", "lot")}
