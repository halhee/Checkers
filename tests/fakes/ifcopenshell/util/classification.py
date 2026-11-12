"""Doublure d'ifcopenshell.util.classification."""


def get_references(element, should_inherit=True):
    return set(getattr(element, "_classifications", None) or ())
