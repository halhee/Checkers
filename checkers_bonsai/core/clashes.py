"""Conflits géométriques entre la maquette ouverte et les autres lots, avec
IfcClash (fourni par Bonsai, comme IfcOpenShell).

Chaque règle compare une sélection d'éléments de la maquette ouverte à une
sélection d'un autre lot, ou de la maquette elle-même (conflits internes),
selon un mode :
- intersection : les volumes s'interpénètrent de plus que la tolérance ;
- collision : les volumes se touchent ou s'interpénètrent ;
- dégagement : les éléments sont plus proches que la distance demandée.

Les sélections utilisent la syntaxe de requête d'IfcOpenShell
(ifcopenshell.util.selector), ex. « IfcDuctSegment, IfcPipeSegment » ;
vide = tous les éléments physiques, hors réservations.

Chaque conflit devient une anomalie de source « Conflit », rattachée aux
éléments de la maquette ouverte ; les éléments de l'autre lot sont gardés à
part (`related_guids`), pour le BCF. Le point du conflit sert à placer la
caméra du point de vue BCF.

Les maquettes sont comparées dans leurs coordonnées propres : si les lots
ne partagent pas le même repère (voir federation.py), les conflits trouvés
ne veulent rien dire.
"""
import logging
from collections import Counter, defaultdict
from dataclasses import dataclass

from .federation import CURRENT
from .structure import BLOCKING, MAJOR, MINOR

MODES = {"intersection": "intersection", "collision": "collision", "clearance": "dégagement"}
CLASH_TYPES = {"protrusion": "dépassement", "pierce": "traversée", "collision": "collision",
               "clearance": "dégagement insuffisant"}
SEVERITIES = (BLOCKING, MAJOR, MINOR)
MAX_DETAILS = 50000
# Clé de la maquette ouverte dans le cache d'IfcClash (jamais un chemin réel).
_CURRENT_KEY = "<maquette ouverte>"


class ClashError(ValueError):
    """Règle de conflits invalide ou IfcClash indisponible."""


@dataclass
class ClashRule:
    other: str = CURRENT          # nom d'un lot associé, ou CURRENT (conflits internes)
    selector_a: str = ""          # éléments de la maquette ouverte
    selector_b: str = ""          # éléments de l'autre lot
    mode: str = "intersection"
    distance: float = 0.01        # tolérance (intersection) ou distance minimale (dégagement), en m
    severity: str = MAJOR

    def label(self) -> str:
        target = "Conflit interne" if self.other == CURRENT else f"Conflit avec « {self.other} »"
        mode = MODES.get(self.mode, self.mode)
        detail = f"{mode}, {self.distance * 100:g} cm" if self.mode != "collision" else mode
        return f"{target} ({detail})"


def validate_rules(rules, lot_names) -> None:
    for number, rule in enumerate(rules, 1):
        if rule.mode not in MODES:
            raise ClashError(f"Règle de conflits n° {number} : mode « {rule.mode} » inconnu")
        if rule.severity not in SEVERITIES:
            raise ClashError(f"Règle de conflits n° {number} : gravité « {rule.severity} » inconnue")
        if rule.other != CURRENT and rule.other not in lot_names:
            raise ClashError(f"Règle de conflits n° {number} : aucun lot associé ne s'appelle « {rule.other} »")
        if not 0.0 <= rule.distance <= 10.0:
            raise ClashError(f"Règle de conflits n° {number} : distance hors limites (0 à 10 m)")


def _clasher_module():
    try:
        from ifcclash import ifcclash
    except ImportError:
        raise ClashError("IfcClash est introuvable : vérifiez que Bonsai est installé et activé")
    return ifcclash


def _select(ifc_file, selector: str) -> list:
    """Éléments qu'IfcClash testera, avec la même règle que lui : la requête
    si elle est donnée, sinon tous les éléments physiques hors réservations.
    Seuls ceux qui ont une géométrie comptent."""
    if selector.strip():
        import ifcopenshell.util.selector
        try:
            elements = ifcopenshell.util.selector.filter_elements(ifc_file, selector)
        except Exception as exc:
            raise ClashError(f"Sélection « {selector} » illisible ({type(exc).__name__})")
    else:
        elements = set(ifc_file.by_type("IfcElement")) - set(ifc_file.by_type("IfcFeatureElement"))
    return [e for e in elements if getattr(e, "Representation", None) is not None]


def _clash_set(rule: ClashRule, key_b: str) -> dict:
    def source(key, selector):
        return {"file": key, "mode": "i", "selector": selector} if selector.strip() else {"file": key}

    clash_set = {"name": rule.label(), "a": [source(_CURRENT_KEY, rule.selector_a)],
                 "b": [source(key_b, rule.selector_b)], "mode": rule.mode}
    if rule.mode == "intersection":
        clash_set.update(tolerance=rule.distance, check_all=True)
    elif rule.mode == "collision":
        clash_set.update(allow_touching=False)
    else:
        clash_set.update(clearance=rule.distance, check_all=False)
    return clash_set


def run_clashes(ifc_file, lots: list, rules: list, level_of) -> dict:
    """`lots` : résultat de federation.open_lots ; `rules` : liste de
    ClashRule. Retourne :
    - rules : par règle, label, lot, severity, status (« ok », « ko »,
      « erreur »), message, count (nombre de conflits) ;
    - anomalies : {(niveau, classe IFC, intitulé): {"severity", "guids",
      "count", "related_guids", "point"}} ;
    - details : un dict par conflit (au plus MAX_DETAILS) ; truncated ;
    - total : nombre de conflits ; blocking / major / minor / errors."""
    ifcclash = _clasher_module()
    files = {lot["name"]: (lot["path"], lot["ifc"]) for lot in lots}
    validate_rules(rules, files)
    logger = logging.getLogger("Checkers.clash")

    summary, details = [], []
    anomalies = {}
    groups = defaultdict(lambda: {"guids": set(), "related": set(), "point": None, "pairs": 0})
    failed = Counter()
    total = 0

    for rule in rules:
        internal = rule.other == CURRENT
        key_b, other_file = (_CURRENT_KEY, ifc_file) if internal else files[rule.other]
        base = {"label": rule.label(), "lot": rule.other, "severity": rule.severity, "mode": rule.mode,
                "selector_a": rule.selector_a, "selector_b": rule.selector_b, "count": 0}
        try:
            if not _select(ifc_file, rule.selector_a):
                raise ClashError("aucun élément avec géométrie dans la sélection de la maquette ouverte")
            if not _select(other_file, rule.selector_b):
                raise ClashError(f"aucun élément avec géométrie dans la sélection de « {rule.other} »")
            settings = ifcclash.ClashSettings()
            settings.logger = logger
            clasher = ifcclash.Clasher(settings)
            # Cache d'IfcClash prérempli : la maquette ouverte est utilisée
            # telle quelle, même non enregistrée, et aucun lot n'est relu.
            clasher.ifcs[_CURRENT_KEY] = ifc_file
            if not internal:
                clasher.ifcs[key_b] = other_file
            clasher.clash_sets = [_clash_set(rule, key_b)]
            clasher.clash()
            results = list(clasher.clash_sets[0].get("clashes", {}).values())
        except ClashError as exc:
            summary.append({**base, "status": "erreur", "message": str(exc)})
            failed["erreur"] += 1
            continue
        except AssertionError:
            # IfcClash vérifie par des assertions que la géométrie a pu être
            # calculée : c'est le cas d'échec le plus courant.
            summary.append({**base, "status": "erreur",
                            "message": "géométrie non calculable pour les éléments sélectionnés"})
            failed["erreur"] += 1
            continue
        except Exception as exc:
            summary.append({**base, "status": "erreur", "message": f"IfcClash a échoué ({type(exc).__name__})"})
            failed["erreur"] += 1
            continue

        seen, count = set(), 0
        for clash in results:
            a, b = clash["a_global_id"], clash["b_global_id"]
            if a == b:
                continue
            if internal:
                if frozenset((a, b)) in seen:
                    continue
                seen.add(frozenset((a, b)))
            try:
                element = ifc_file.by_guid(a)
            except Exception:
                continue
            count += 1
            requirement = f"{rule.label()} : {clash['b_ifc_class']}"
            group = groups[(level_of(element), clash["a_ifc_class"], requirement)]
            group["severity"] = rule.severity
            group["guids"].add(a)
            (group["guids"] if internal else group["related"]).add(b)
            group["pairs"] += 1
            if group["point"] is None:
                group["point"] = tuple(float(v) for v in clash.get("p1") or (0.0, 0.0, 0.0))
            if len(details) < MAX_DETAILS:
                details.append({
                    "rule": rule.label(), "level": level_of(element), "a_class": clash["a_ifc_class"], "a_guid": a,
                    "a_name": clash.get("a_name") or "", "lot": rule.other, "b_class": clash["b_ifc_class"],
                    "b_guid": b, "b_name": clash.get("b_name") or "",
                    "type": CLASH_TYPES.get(str(clash.get("type")), str(clash.get("type") or "")),
                    "distance": float(clash.get("distance") or 0.0),
                    "point": tuple(round(float(v), 3) for v in clash.get("p1") or (0.0, 0.0, 0.0)),
                })
        total += count
        if count:
            failed[rule.severity] += 1
            summary.append({**base, "status": "ko", "message": f"{count} conflit(s)", "count": count})
        else:
            summary.append({**base, "status": "ok", "message": "aucun conflit"})

    for key, group in groups.items():
        anomalies[key] = {"severity": group["severity"], "guids": group["guids"], "count": len(group["guids"]),
                          "related_guids": group["related"], "point": group["point"], "pairs": group["pairs"]}
    return {"rules": summary, "anomalies": anomalies, "details": details, "truncated": total > len(details),
            "total": total, "blocking": failed[BLOCKING], "major": failed[MAJOR], "minor": failed[MINOR],
            "errors": failed["erreur"]}
