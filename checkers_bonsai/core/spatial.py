"""Niveau d'un élément pour regrouper les anomalies : étage pour un
bâtiment, partie d'ouvrage pour une infrastructure (IFC 4.3 : IfcFacilityPart,
ce qui couvre IfcMarinePart des ouvrages portuaires, IfcBridgePart,
IfcRoadPart, IfcRailwayPart).

Lecture par les attributs inverses standard (ContainedInStructure,
Decomposes), valable en IFC2X3, IFC4 et IFC4X3.
"""

NO_LEVEL = "Sans niveau"
LEVEL_CLASSES = ("IfcBuildingStorey", "IfcFacilityPart")
# Conteneurs de haut niveau : l'élément y est rangé sans étage ni partie.
TOP_CLASSES = ("IfcBuilding", "IfcFacility", "IfcSite", "IfcProject")
MAX_DEPTH = 30


def _first(relations):
    return relations[0] if relations else None


def _parent(entity):
    """Conteneur spatial direct, sinon objet parent par agrégation (ex. un
    limon dans un escalier, un espace dans un étage)."""
    rel = _first(getattr(entity, "ContainedInStructure", None) or ())
    if rel is not None:
        return rel.RelatingStructure
    for inverse in ("Decomposes", "Nests"):  # agrégation, puis imbrication (IFC4)
        rel = _first(getattr(entity, inverse, None) or ())
        if rel is not None:
            return rel.RelatingObject
    return None


def _is_a(entity, ifc_class) -> bool:
    # Prudence : une classe absente du schéma du modèle (IfcFacilityPart en
    # IFC2X3) ne doit jamais faire échouer la recherche.
    try:
        return bool(entity.is_a(ifc_class))
    except Exception:
        return False


def _is_any(entity, classes) -> bool:
    return any(_is_a(entity, c) for c in classes)


def level_of(element) -> str:
    current = element
    for _ in range(MAX_DEPTH):
        parent = _parent(current)
        if parent is None:
            return NO_LEVEL
        if _is_any(parent, LEVEL_CLASSES):
            return parent.Name or getattr(parent, "LongName", None) or parent.is_a()
        if _is_any(parent, TOP_CLASSES):
            return NO_LEVEL
        current = parent
    return NO_LEVEL


class LevelCache:
    """level_of() mémorisé par élément (sollicité par le contrôle PSet puis
    par l'IDS sur les mêmes éléments)."""

    def __init__(self):
        self._cache = {}

    def __call__(self, element) -> str:
        key = element.id() if hasattr(element, "id") else id(element)
        if key not in self._cache:
            try:
                self._cache[key] = level_of(element)
            except Exception:
                self._cache[key] = NO_LEVEL
        return self._cache[key]
