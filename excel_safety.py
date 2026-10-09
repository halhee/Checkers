"""Neutralise l'injection de formule Excel (CSV/Formula Injection -- voir
OWASP "CSV Injection"). openpyxl interprète automatiquement toute valeur de
cellule de type `str` commençant par '=', '+', '-' ou '@' comme une formule
vivante (`cell.data_type` passe à `'f'`), y compris quand cette valeur vient
entièrement de données utilisateur : un nom de lot saisi au formulaire, ou un
nom d'élément/d'étage/de matériau lu dans un fichier IFC uploadé (format
texte librement éditable par quiconque fournit le fichier).

Les rapports générés par Checkers (rapport Excel, Fiche de Recette) sont
faits pour être ouverts par quelqu'un d'autre que la personne qui a fourni
les fichiers d'entrée -- exactement le scénario où une formule injectée
devient exploitable (ex. une formule DDE ou HYPERLINK/WEBSERVICE exécutée à
l'ouverture par le relecteur de la recette)."""

_DANGEROUS_PREFIXES = ('=', '+', '-', '@', '\t', '\r')


def excel_safe(value):
    """Si `value` est une chaîne commençant par un caractère qu'Excel/openpyxl
    peut interpréter comme le début d'une formule, la préfixe d'une
    apostrophe pour la forcer en texte brut. Ne touche à rien d'autre
    (nombres, None, bool...), qu'openpyxl n'interprète jamais comme formule."""
    if isinstance(value, str) and value.startswith(_DANGEROUS_PREFIXES):
        return "'" + value
    return value
