"""Contrôle des PSet et paramètres requis (fichier Excel de règles)."""
from collections import defaultdict

import ifcopenshell.util.element

from .rules import found_type, is_empty, value_matches

MISSING_PSET = "PSet absent"
MISSING_PARAM = "Paramètre absent ou vide"
# Plafond de lignes de détail conservées pour le rapport (les compteurs et
# la sélection des éléments restent exacts au-delà).
MAX_FINDINGS = 50000


def _elements_by_rule_class(ifc_file, classes, warnings):
    """{élément: classe de règle}. Les sous-classes sont incluses (une règle
    IfcWall s'applique aussi à IfcWallStandardCase en IFC2X3). Si plusieurs
    classes déclarées couvrent le même élément, la plus précise l'emporte :
    c'est celle qui regroupe le moins d'éléments, puisqu'une sous-classe est
    toujours incluse dans sa classe parente."""
    by_class = {}
    for rule_class in classes:
        try:
            by_class[rule_class] = ifc_file.by_type(rule_class)
        except Exception:
            warnings.append(f"Classe « {rule_class} » inconnue du schéma {ifc_file.schema} : ignorée")
    assignment = {}
    for rule_class in sorted(by_class, key=lambda c: len(by_class[c]), reverse=True):
        for element in by_class[rule_class]:
            assignment[element] = rule_class
    return assignment


def check_psets(ifc_file, rules: dict, level_of) -> dict:
    """Retourne :
    - checked / valid : nombre d'éléments contrôlés / conformes ;
    - valid_guids / invalid_guids ;
    - anomalies : {(niveau, classe IFC, PSet): ensemble de GlobalId} ;
    - findings : détail des non-conformités (au plus MAX_FINDINGS) ;
    - truncated, warnings."""
    warnings = []
    assignment = _elements_by_rule_class(ifc_file, rules["classes"], warnings)
    anomalies = defaultdict(set)
    findings = []
    valid_guids, invalid_guids = [], []
    total_findings = 0

    for element, rule_class in assignment.items():
        required = rules["psets"].get(rule_class, {})
        # get_psets inclut les propriétés héritées du type et les quantités.
        psets = ifcopenshell.util.element.get_psets(element) if required else {}
        level = level_of(element)
        problems = []
        for pset_name, params in required.items():
            values = psets.get(pset_name)
            if values is None:
                problems.append((pset_name, "", "", MISSING_PSET))
                continue
            for param, param_type in params.items():
                value = values.get(param)
                if is_empty(value):
                    problems.append((pset_name, param, "", MISSING_PARAM))
                elif not value_matches(value, param_type):
                    problems.append((pset_name, param, value,
                                     f"Type incorrect (attendu {param_type}, trouvé {found_type(value)})"))

        guid = element.GlobalId
        if not problems:
            valid_guids.append(guid)
            continue
        invalid_guids.append(guid)
        for pset_name, param, value, problem in problems:
            anomalies[(level, element.is_a(), pset_name)].add(guid)
            total_findings += 1
            if len(findings) < MAX_FINDINGS:
                findings.append({
                    "level": level, "ifc_class": element.is_a(), "guid": guid,
                    "name": getattr(element, "Name", None) or "", "tag": getattr(element, "Tag", None) or "",
                    "pset": pset_name, "param": param, "value": value, "problem": problem,
                })

    return {
        "checked": len(assignment), "valid": len(valid_guids),
        "valid_guids": valid_guids, "invalid_guids": invalid_guids,
        "anomalies": dict(anomalies), "findings": findings,
        "truncated": total_findings > len(findings),
        "warnings": warnings,
    }
