"""Passage des coordonnées IFC (mètres, repère du projet) à celles de la
scène Blender.

Bonsai peut décaler la scène (« false origin ») pour les maquettes aux
grandes coordonnées : il retranche alors un point et tourne les axes. Les
points de conflit d'IfcClash et les géométries des autres lots sont en
coordonnées IFC : ils doivent subir la même transformation pour tomber au
bon endroit dans la vue 3D. Même calcul que
ifcopenshell.util.geolocation.global2local, sans numpy.
"""
import math

IDENTITY = ((1.0, 0.0, 0.0, 0.0), (0.0, 1.0, 0.0, 0.0), (0.0, 0.0, 1.0, 0.0), (0.0, 0.0, 0.0, 1.0))
# Point de vue autour d'un conflit : décalage de l'œil, comme la caméra BCF.
VIEW_DISTANCE = 6.0


def world_to_scene(offset=None) -> tuple:
    """Matrice 4×4 (lignes) qui envoie un point IFC dans la scène.
    `offset` : ((x, y, z) en mètres, abscisse, ordonnée de l'axe X), ou
    None si la scène n'est pas décalée."""
    if not offset:
        return IDENTITY
    (x, y, z), abscissa, ordinate = offset
    theta = math.atan2(ordinate, abscissa)
    c, s = math.cos(theta), math.sin(theta)
    # local = R(-θ) · (global - décalage)
    return ((c, s, 0.0, -(c * x + s * y)),
            (-s, c, 0.0, -(-s * x + c * y)),
            (0.0, 0.0, 1.0, -z),
            (0.0, 0.0, 0.0, 1.0))


def transform(matrix, point) -> tuple:
    x, y, z = (float(v) for v in point)
    return tuple(row[0] * x + row[1] * y + row[2] * z + row[3] for row in matrix[:3])


def geometry_settings():
    """Réglages IfcOpenShell : géométrie en coordonnées du projet (mètres),
    placements compris."""
    import ifcopenshell.geom
    settings = ifcopenshell.geom.settings()
    try:
        settings.set("use-world-coords", True)
    except Exception:
        settings.set(settings.USE_WORLD_COORDS, True)
    return settings


def element_mesh(element, settings=None) -> tuple:
    """(sommets, triangles) d'un élément, en coordonnées IFC (mètres)."""
    import ifcopenshell.geom
    shape = ifcopenshell.geom.create_shape(settings or geometry_settings(), element)
    verts, faces = shape.geometry.verts, shape.geometry.faces
    return ([tuple(verts[i:i + 3]) for i in range(0, len(verts), 3)],
            [tuple(faces[i:i + 3]) for i in range(0, len(faces), 3)])
