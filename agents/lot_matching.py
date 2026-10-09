"""Recours IA de dernier ressort pour associer le nom d'un lot uploadé à un
corps d'état déclaré dans l'Excel de paramètres requis, quand la
correspondance déterministe (corps_etat.py::match_lot_name, exacte puis
floue) échoue. Même schéma que agents/recommendations.py : import paresseux
de `crewai`, local-first via agents/llm_tiers.py, jamais bloquant -- en cas
d'échec ou d'agents désactivés, retourne simplement None (le lot est alors
analysé sans contrôle PSet, voir Checkers.py::process_files).

La tâche est volontairement une classification fermée (choisir un code
parmi une liste donnée, ou "AUCUN") : l'agent ne peut pas inventer de
nouveau corps d'état, seulement se tromper de choix ou ne pas savoir.
"""

import os


def _ai_enabled() -> bool:
    return os.environ.get("ENABLE_AI_RECOMMENDATIONS", "true").strip().lower() not in ("false", "0", "no")


def match_lot_with_ai(lot_name: str, registry: dict):
    """Retourne un code de `registry`, ou None si les agents sont
    désactivés, indisponibles, ou si aucun choix fiable n'a été fait."""
    if not _ai_enabled() or not registry:
        return None

    try:
        from .llm_tiers import run_with_fallback
    except Exception:
        return None

    os.environ.setdefault("OPENAI_API_KEY", "NA")  # crewai/litellm en ont besoin même sans l'utiliser

    options_text = "\n".join(f"- {code} : {nom}" for code, nom in registry.items())
    valid_codes_lower = {code.lower(): code for code in registry}

    def build(llm):
        from crewai import Agent, Task

        agent = Agent(
            role="Classificateur de lots BIM",
            goal="Associer le nom d'un lot au corps d'état correspondant, strictement parmi une liste fermée.",
            backstory=(
                "Un assistant qui choisit exactement parmi les options fournies -- il n'invente "
                "jamais de nouvelle catégorie et dit clairement quand rien ne correspond."
            ),
            llm=llm,
            verbose=False,
        )
        task = Task(
            description=(
                f"Voici les corps d'état disponibles (code : nom complet) :\n{options_text}\n\n"
                f"Le lot uploadé s'appelle : \"{lot_name}\".\n\n"
                "Réponds uniquement par le code (et rien d'autre : pas de phrase, pas de ponctuation) "
                "du corps d'état qui correspond le mieux à ce nom de lot. Si aucun ne correspond "
                "raisonnablement, réponds exactement par le mot AUCUN."
            ),
            expected_output="Un seul mot : soit un code de la liste, soit AUCUN.",
            agent=agent,
        )
        return agent, task

    def validator(text: str) -> bool:
        candidate = text.strip().splitlines()[-1].strip().strip(".:,;\"'")
        return candidate.lower() in valid_codes_lower or candidate.upper() == "AUCUN"

    try:
        result_text, _tier = run_with_fallback(build, validator)
    except Exception:
        return None

    candidate = result_text.strip().splitlines()[-1].strip().strip(".:,;\"'")
    return valid_codes_lower.get(candidate.lower())
