"""Doublure d'ifcopenshell.util.unit."""


def get_project_unit(ifc_file, unit_type, use_cache=False):
    return ifc_file.length_unit if unit_type == "LENGTHUNIT" else None
