"""Doublure d'ifcopenshell.util.placement (IfcOpenShell 0.9)."""


def get_placement(placement=None, should_return_si=False):
    if placement is None:
        return [[1.0, 0.0, 0.0, 0.0], [0.0, 1.0, 0.0, 0.0], [0.0, 0.0, 1.0, 0.0], [0.0, 0.0, 0.0, 1.0]]
    return placement.matrix
