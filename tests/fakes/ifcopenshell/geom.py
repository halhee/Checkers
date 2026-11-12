"""Doublure d'ifcopenshell.geom : la géométrie d'un élément est la boîte
`_box` ((xmin, ymin, zmin), (xmax, ymax, zmax)) donnée dans les tests, en
coordonnées du projet."""
import types


class settings:
    USE_WORLD_COORDS = "use-world-coords"

    def __init__(self):
        self.values = {}

    def set(self, name, value):
        self.values[name] = value


def create_shape(settings, element):
    box = getattr(element, "_box", None)
    if box is None:
        raise RuntimeError("Representation not found")
    (x0, y0, z0), (x1, y1, z1) = box
    verts = []
    for x, y, z in ((x0, y0, z0), (x1, y0, z0), (x1, y1, z0), (x0, y1, z0),
                    (x0, y0, z1), (x1, y0, z1), (x1, y1, z1), (x0, y1, z1)):
        verts += [float(x), float(y), float(z)]
    quads = ((0, 1, 2, 3), (4, 5, 6, 7), (0, 1, 5, 4), (1, 2, 6, 5), (2, 3, 7, 6), (3, 0, 4, 7))
    faces = []
    for a, b, c, d in quads:
        faces += [a, b, c, a, c, d]
    return types.SimpleNamespace(geometry=types.SimpleNamespace(verts=verts, faces=faces))
