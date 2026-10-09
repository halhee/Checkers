"""Grounding checks run against an agent's text output before it's trusted.
Targets the two failure modes actually observed with small local models
(see ~/Desktop/mon-projet, same pattern): inventing entities that don't
exist in the data, and omitting a real, significant finding.

Unlike mon-projet's validators.py, `known_vocabulary()` isn't owned here --
it depends on the specific IFC model/report being analyzed, so callers in
recommendations.py build it from the precomputed facts they already have.
"""

import re

from rapidfuzz import fuzz, process

_NAME_PATTERN = re.compile(r"\*\*(?:[^*']*?['‘’])?([^*'‘’]+?)['‘’]?\*\*")


def _extract_named_entities(text: str) -> list[str]:
    return [m.strip() for m in _NAME_PATTERN.findall(text) if m.strip()]


def grounded(text: str, must_include: list[str], vocabulary: set[str], min_score: float = 80.0) -> bool:
    """True if `text` mentions every string in `must_include` and doesn't
    name any bolded entity outside `vocabulary` (fuzzy-matched)."""
    text_l = text.lower()
    if not all(term.lower() in text_l for term in must_include):
        return False

    if not vocabulary:
        return True

    choices = list(vocabulary)
    for name in _extract_named_entities(text):
        match = process.extractOne(name.lower(), choices, scorer=fuzz.token_sort_ratio)
        if match is None or match[1] < min_score:
            return False
    return True


def has_sections(text: str, headers: list[str]) -> bool:
    """Case-insensitive, markdown-syntax-agnostic: a small model won't
    reliably reproduce exact '## Heading' casing/punctuation, so this
    checks the header wording appears somewhere, not the literal syntax."""
    text_l = text.lower()
    return all(h.lower() in text_l for h in headers)
