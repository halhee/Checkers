"""Doublure d'ifcopenshell.util.system."""


def get_element_systems(element):
    return list(getattr(element, "_systems", None) or ())
