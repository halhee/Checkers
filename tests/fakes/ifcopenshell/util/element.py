"""Doublure d'ifcopenshell.util.element (mêmes signatures que la bibliothèque)."""


def get_psets(element, psets_only=False, qtos_only=False, should_inherit=True, verbose=False):
    return getattr(element, "_psets", {}) or {}


def get_pset(element, name, prop=None, psets_only=False, qtos_only=False, should_inherit=True, verbose=False):
    pset = get_psets(element).get(name)
    if pset is None:
        return None
    return pset.get(prop) if prop else dict(pset)


def get_aggregate(element):
    relations = getattr(element, "Decomposes", None) or ()
    return relations[0].RelatingObject if relations else None


def get_container(element, should_get_direct=False, ifc_class=None):
    """Conteneur spatial direct, ou celui de l'assemblage dont l'élément
    fait partie."""
    relations = getattr(element, "ContainedInStructure", None) or ()
    if relations:
        return relations[0].RelatingStructure
    parent = get_aggregate(element)
    if parent is not None and not should_get_direct:
        return get_container(parent)
    return None


def get_type(element):
    return getattr(element, "_type", None)


def get_predefined_type(element):
    return getattr(element, "PredefinedType", None)
