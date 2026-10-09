"""Registre des corps d'état déclarés dans le fichier Excel de paramètres
requis (onglet "Corps_Etat", optionnel -- voir Checkers.py::process_files) et
mise en correspondance du nom de lot saisi par l'utilisateur avec l'un
d'eux. Purement déterministe (pas d'appel LLM ici) : le recours à un agent
IA pour les cas ambigus vit dans agents/lot_matching.py, appelé seulement
quand cette correspondance déterministe échoue.
"""

import unicodedata
from typing import Optional

import pandas as pd

_FUZZY_THRESHOLD = 85.0
# Écart minimum entre le meilleur et le deuxième meilleur score pour
# considérer la correspondance comme non ambiguë.
_FUZZY_MARGIN = 10.0


def _normalize(text: str) -> str:
    """Minuscules, sans accents, espaces superflus retirés -- pour une
    comparaison exacte tolérante aux variations de saisie courantes."""
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.lower().split())


def load_registry(excel_file_path: str) -> Optional[dict]:
    """Lit l'onglet "Corps_Etat" (colonnes Code, Nom) s'il existe. Retourne
    None si l'onglet est absent -- signal que le fichier est en mode "un
    seul jeu de règles" (comportement historique de l'app, inchangé)."""
    try:
        xls = pd.ExcelFile(excel_file_path, engine="openpyxl")
    except Exception:
        return None
    if "Corps_Etat" not in xls.sheet_names:
        return None

    df = pd.read_excel(excel_file_path, sheet_name="Corps_Etat", engine="openpyxl")
    registry = {}
    for _, row in df.iterrows():
        code = str(row["Code"]).strip()
        nom = str(row["Nom"]).strip()
        # Ignore les lignes incomplètes (note laissée dans l'onglet, ligne
        # partiellement remplie...) plutôt que d'empoisonner le registre
        # avec une entrée invalide qui ne correspondra jamais à rien.
        if code and code.lower() != "nan" and nom and nom.lower() != "nan":
            registry[code] = nom
    return registry or None


def match_lot_name(lot_name: str, registry: dict) -> tuple:
    """Retourne (code, méthode) où méthode ∈ "exact" | "flou" | "ambigu".
    "ambigu" signifie : pas de correspondance déterministe fiable -- à
    l'appelant de décider entre un recours IA (agents/lot_matching.py) et
    l'absence de contrôle PSet pour ce lot."""
    from rapidfuzz import fuzz, process

    normalized_lot = _normalize(lot_name)

    for code, nom in registry.items():
        if _normalize(nom) == normalized_lot or _normalize(code) == normalized_lot:
            return code, "exact"

    choices = {code: _normalize(nom) for code, nom in registry.items()}
    match = process.extractOne(normalized_lot, choices, scorer=fuzz.token_sort_ratio)
    if match is None:
        return None, "ambigu"
    _, best_score, best_code = match

    others = [
        fuzz.token_sort_ratio(normalized_lot, value)
        for code, value in choices.items() if code != best_code
    ]
    runner_up = max(others) if others else 0.0

    if best_score >= _FUZZY_THRESHOLD and (best_score - runner_up) >= _FUZZY_MARGIN:
        return best_code, "flou"
    return None, "ambigu"
