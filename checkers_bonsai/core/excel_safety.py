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

_DANGEROUS_PREFIXES = ('=', '+', '-', '@', '\t', '\r')


def excel_safe(value):
    """Si `value` est une chaîne commençant par un caractère qu'Excel/openpyxl
    peut interpréter comme le début d'une formule, la préfixe d'une
    apostrophe pour la forcer en texte brut. Ne touche à rien d'autre
    (nombres, None, bool...), qu'openpyxl n'interprète jamais comme formule."""
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
