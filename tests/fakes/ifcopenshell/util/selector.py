"""Doublure d'ifcopenshell.util.selector : requête limitée à une liste de
classes séparées par des virgules (« IfcBeam, IfcColumn »)."""


def filter_elements(ifc_file, query, elements=None, edit_in_place=False):
    found = set()
    for name in (part.strip() for part in query.split(",")):
        if not name.startswith("Ifc"):
            raise ValueError(f"Requête non comprise : {query}")
        try:
            found.update(ifc_file.by_type(name))
        except RuntimeError:
            raise ValueError(f"Classe inconnue : {name}")
    return found
