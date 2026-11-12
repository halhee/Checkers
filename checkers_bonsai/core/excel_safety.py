"""Neutralise l'injection de formule Excel (CSV/Formula Injection -- voir
OWASP "CSV Injection"). openpyxl interprète automatiquement toute valeur de
cellule de type `str` commençant par '=', '+', '-' ou '@' comme une formule
vivante (`cell.data_type` passe à `'f'`), y compris quand cette valeur vient
entièrement de données extérieures : un nom de lot saisi dans le panneau, ou
un nom d'élément, d'étage ou de local lu dans la maquette IFC (format texte
librement éditable par quiconque fournit le fichier).

Les classeurs produits par Checkers (fiche de recette, inventaire) sont
faits pour être ouverts par quelqu'un d'autre que la personne qui a fourni
la maquette -- exactement le scénario où une formule injectée devient
exploitable (ex. une formule DDE ou HYPERLINK/WEBSERVICE exécutée à
l'ouverture par le relecteur de la recette)."""

import re

_DANGEROUS_PREFIXES = ('=', '+', '-', '@', '\t', '\r')
# Caractères de contrôle interdits en XML 1.0 : openpyxl refuse d'écrire une
# cellule qui en contient, et un BCF qui en contient est illisible. Un nom
# d'élément IFC peut en porter (séquence d'échappement mal encodée).
_CONTROL_CHARACTERS = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


def clean_text(value):
    """Retire d'une chaîne les caractères de contrôle interdits en XML.
    Ne touche à rien d'autre."""
    return _CONTROL_CHARACTERS.sub("", value) if isinstance(value, str) else value


def header_footer_text(text: str) -> str:
    """Texte sûr pour un en-tête ou un pied de page Excel, où « & » introduit
    un code de mise en forme (« &D » = date) : il est doublé."""
    return clean_text(text).replace("&", "&&")


def excel_safe(value):
    """Si `value` est une chaîne commençant par un caractère qu'Excel/openpyxl
    peut interpréter comme le début d'une formule, la préfixe d'une
    apostrophe pour la forcer en texte brut ; en retire aussi les caractères
    de contrôle. Ne touche à rien d'autre (nombres, None, bool...),
    qu'openpyxl n'interprète jamais comme formule."""
    value = clean_text(value)
    if isinstance(value, str) and value.startswith(_DANGEROUS_PREFIXES):
        return "'" + value
    return value


def neutralize_formulas(workbook) -> int:
    """Repasse sur tout le classeur et force en texte toute cellule restée
    « formule » (`data_type == 'f'`). À appeler juste avant `save()` : les
    rapports Checkers ne contiennent aucune formule volontaire. Retourne le
    nombre de cellules neutralisées."""
    count = 0
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.data_type == 'f':
                    cell.data_type = 's'
                    count += 1
    return count
