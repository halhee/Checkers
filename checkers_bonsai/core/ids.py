"""Contrôle IDS (Information Delivery Specification, buildingSMART) avec
ifctester, fourni par Bonsai.

Le fichier IDS est traité comme une entrée non fiable (il est souvent reçu
d'un tiers) : taille limitée, UTF-8 obligatoire, DOCTYPE/ENTITY refusés
(XXE, expansion d'entités), motifs regex à quantificateurs imbriqués refusés
(heuristique anti-ReDoS), puis validation XSD stricte par ifctester avec les
schémas embarqués, sans téléchargement.
"""
import re
from collections import defaultdict

MAX_IDS_BYTES = 10 * 1024 * 1024
MAX_PATTERN_LENGTH = 256
MAX_DETAILS = 50000

STATUS_LABELS = {"conforme": "Conforme", "non_conforme": "Non conforme", "ignoree": "Ignorée (autre version IFC)"}
CARDINALITY_LABELS = {"required": "Obligatoire", "optional": "Facultative", "prohibited": "Interdite"}


class IdsFileError(ValueError):
    """Fichier IDS refusé (illisible, invalide ou jugé dangereux)."""


_FORBIDDEN = ("<!doctype", "<!entity")
_ENCODING_RE = re.compile(r"""<\?xml[^>]*\bencoding\s*=\s*["']([^"']+)["']""", re.I)
_PATTERN_RE = re.compile(r"""<(?:[\w.-]+:)?pattern\b[^>]*?\bvalue\s*=\s*(?:"([^"]*)"|'([^']*)')""", re.I)
_NESTED_QUANTIFIER_RE = re.compile(r"\((?:[^()\\]|\\.)*[+*](?:[^()\\]|\\.)*\)\s*(?:[+*]|\{\d+,\d*\})")


def check_ids_text(raw: bytes) -> str:
    """Garde-fous avant tout parsing XML. Retourne le texte décodé."""
    if len(raw) > MAX_IDS_BYTES:
        raise IdsFileError(f"Fichier IDS trop volumineux (max {MAX_IDS_BYTES // (1024 * 1024)} Mo)")
    if b"\x00" in raw or raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        raise IdsFileError("Encodage du fichier IDS non supporté (UTF-8 requis)")
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        raise IdsFileError("Encodage du fichier IDS non supporté (UTF-8 requis)")
    declared = _ENCODING_RE.search(text[:200])
    if declared and declared.group(1).lower().replace("_", "-") not in ("utf-8", "utf8", "us-ascii", "ascii"):
        raise IdsFileError("Encodage déclaré du fichier IDS non supporté (UTF-8 requis)")
    if any(token in text.lower() for token in _FORBIDDEN):
        raise IdsFileError("Fichier IDS refusé : déclaration DOCTYPE/ENTITY interdite")
    for match in _PATTERN_RE.finditer(text):
        pattern = match.group(1) if match.group(1) is not None else match.group(2)
        if len(pattern) > MAX_PATTERN_LENGTH:
            raise IdsFileError(f"Fichier IDS refusé : motif regex trop long (max {MAX_PATTERN_LENGTH} caractères)")
        if _NESTED_QUANTIFIER_RE.search(pattern):
            raise IdsFileError("Fichier IDS refusé : motif regex à quantificateurs imbriqués")
    return text


def load_ids(path: str):
    from ifctester import ids

    with open(path, "rb") as fh:
        check_ids_text(fh.read(MAX_IDS_BYTES + 1))
    try:
        document = ids.open(path, validate=True)
    except Exception as exc:
        message = str(exc).strip()
        reason = message.splitlines()[0][:300] if message else type(exc).__name__
        raise IdsFileError(f"Fichier IDS invalide : {reason}")
    if not document.specifications:
        raise IdsFileError("Fichier IDS sans aucune spécification")
    return document


def ids_title(document) -> str:
    return (getattr(document, "info", None) or {}).get("title", "") or ""


def _cardinality(spec) -> str:
    if spec.maxOccurs == 0:
        return "prohibited"
    if spec.minOccurs != 0:
        return "required"
    return "optional"


def anomaly_label(spec_name: str) -> str:
    return f"IDS : {spec_name}"


def check_ids(document, ifc_file, level_of) -> dict:
    """Valide le modèle. Les spécifications visant une autre version IFC
    que celle du modèle sont ignorées (comportement normatif), pas comptées
    en échec. Retourne specifications, compteurs, anomalies {(niveau,
    classe, « IDS : spécification »): GlobalId} et details."""
    document.validate(ifc_file, should_filter_version=True)

    specifications, details = [], []
    anomalies = defaultdict(set)
    applicable_total = failed_total = specs_checked = specs_failed = cardinality_failures = 0
    total_details = 0

    for spec in document.specifications:
        cardinality = _cardinality(spec)
        base = {"name": spec.name, "identifier": getattr(spec, "identifier", None) or "",
                "description": spec.description or "", "instructions": spec.instructions or "",
                "cardinality": cardinality}
        if spec.status is None:
            specifications.append({**base, "status": "ignoree", "applicable": 0, "failed": 0, "passed": 0,
                                   "failed_requirements": ""})
            continue

        applicable_entities = list(spec.applicable_entities)
        if cardinality == "prohibited":
            failed_entities = applicable_entities
            failure_rows = [(e, "Élément interdit par cette spécification", "Interdit") for e in failed_entities]
        else:
            failed_entities = list(spec.failed_entities)
            failure_rows = []
            for requirement in spec.requirements:
                description = requirement.to_string("requirement", spec, requirement)
                for failure in requirement.failures:
                    failure_rows.append((failure["element"], str(failure["reason"]), description))

        for element in failed_entities:
            guid = getattr(element, "GlobalId", None)
            if guid:
                anomalies[(level_of(element), element.is_a(), anomaly_label(spec.name))].add(guid)
        for element, reason, description in failure_rows:
            total_details += 1
            if len(details) < MAX_DETAILS:
                details.append({
                    "specification": spec.name, "requirement": description, "level": level_of(element),
                    "ifc_class": element.is_a(), "guid": getattr(element, "GlobalId", None) or "",
                    "name": getattr(element, "Name", None) or "", "reason": reason,
                })

        failed = len(failed_entities)
        specs_checked += 1
        applicable_total += len(applicable_entities)
        failed_total += failed
        if not spec.status:
            specs_failed += 1
            if cardinality == "required" and not applicable_entities:
                cardinality_failures += 1
        specifications.append({
            **base, "status": "conforme" if spec.status else "non_conforme",
            "applicable": len(applicable_entities), "failed": failed, "passed": len(applicable_entities) - failed,
            "failed_requirements": "; ".join(
                r.to_string("requirement", spec, r) for r in spec.requirements if not r.status),
        })

    return {
        "specifications": specifications, "specs_checked": specs_checked, "specs_failed": specs_failed,
        "cardinality_failures": cardinality_failures, "applicable_total": applicable_total,
        "failed_total": failed_total,
        "pass_rate": (applicable_total - failed_total) / applicable_total if applicable_total else None,
        "anomalies": dict(anomalies), "details": details, "truncated": total_details > len(details),
    }
