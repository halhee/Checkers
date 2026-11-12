"""Affichage d'un conflit dans la vue 3D : repère au point de conflit, vue
centrée dessus, et élément de l'autre lot affiché en filaire par-dessus la
maquette.

L'élément de l'autre lot n'existe pas dans la scène : sa géométrie est
calculée à la demande avec IfcOpenShell, à partir du fichier du lot (relu
seulement s'il a changé). Les objets créés sont temporaires : rangés dans
une collection à part, ils sont effacés avant chaque enregistrement du
fichier .blend, à l'ouverture d'un autre fichier, ou par le bouton
« Effacer ».
"""
import os

import bpy

from . import bonsai_link
from .core import scene as scene_math

COLLECTION = "Checkers – conflit"
MARKER = "Checkers – point de conflit"
GHOST_COLOR = (1.0, 0.45, 0.0, 1.0)
MAX_GHOSTS = 20
MAX_CACHED_LOTS = 3
_lots = {}   # chemin → ((date, taille), modèle)


def clear() -> int:
    """Efface les objets temporaires ; retourne leur nombre."""
    collection = bpy.data.collections.get(COLLECTION)
    if collection is None:
        return 0
    removed = 0
    for obj in list(collection.objects):
        mesh = obj.data
        bpy.data.objects.remove(obj, do_unlink=True)
        if mesh is not None and getattr(mesh, "users", 0) == 0:
            bpy.data.meshes.remove(mesh)
        removed += 1
    bpy.data.collections.remove(collection)
    return removed


def forget_lots() -> None:
    _lots.clear()


def open_lot(path: str):
    """Maquette d'un autre lot, gardée en mémoire tant que le fichier ne
    change pas."""
    stat = os.stat(path)
    stamp = (stat.st_mtime, stat.st_size)
    cached = _lots.get(path)
    if cached is not None and cached[0] == stamp:
        return cached[1]
    import ifcopenshell
    model = ifcopenshell.open(path)
    if len(_lots) >= MAX_CACHED_LOTS:
        _lots.clear()
    _lots[path] = (stamp, model)
    return model


def _collection(scene):
    collection = bpy.data.collections.new(COLLECTION)
    scene.collection.children.link(collection)
    return collection


def _ghost(collection, element, geometry, matrix):
    name = f"{element.is_a()} « {getattr(element, 'Name', None) or element.GlobalId} » (autre lot)"
    vertices, triangles = geometry
    mesh = bpy.data.meshes.new(name)
    # Sommets passés directement dans le repère de la scène : pas de
    # matrice d'objet à convertir.
    mesh.from_pydata([scene_math.transform(matrix, v) for v in vertices], [], triangles)
    mesh.update()
    obj = bpy.data.objects.new(name, mesh)
    obj.display_type = "WIRE"
    obj.show_in_front = True
    obj.color = GHOST_COLOR
    obj.hide_select = True
    collection.objects.link(obj)
    return obj


def _marker(collection, location):
    obj = bpy.data.objects.new(MARKER, None)
    obj.empty_display_type = "SPHERE"
    obj.empty_display_size = 0.15
    obj.show_in_front = True
    obj.location = location
    obj.hide_select = True
    collection.objects.link(obj)
    return obj


def _look_at(context, location) -> None:
    for window in context.window_manager.windows:
        for area in window.screen.areas:
            if area.type != "VIEW_3D":
                continue
            for space in area.spaces:
                region = getattr(space, "region_3d", None)
                if space.type == "VIEW_3D" and region is not None:
                    region.view_location = location
                    region.view_distance = scene_math.VIEW_DISTANCE


def show_conflict(context, point, lot_path: str = "", guids=()) -> dict:
    """Place le repère au point de conflit (coordonnées IFC, en mètres),
    centre la vue dessus, et affiche en filaire les éléments `guids` du lot
    `lot_path`. Retourne {"shown", "missing", "error"}."""
    clear()
    matrix = scene_math.world_to_scene(bonsai_link.scene_offset())
    outcome = {"shown": 0, "missing": 0, "error": ""}
    collection = _collection(context.scene)
    if lot_path and guids:
        try:
            model = open_lot(lot_path)
            settings = scene_math.geometry_settings()
        except Exception as exc:
            outcome["error"] = f"maquette du lot illisible ({type(exc).__name__})"
            model = None
        for guid in list(guids)[:MAX_GHOSTS] if model is not None else ():
            try:
                element = model.by_guid(guid)
                geometry = scene_math.element_mesh(element, settings)
            except Exception:
                outcome["missing"] += 1
                continue
            _ghost(collection, element, geometry, matrix)
            outcome["shown"] += 1
    if point is not None:
        location = scene_math.transform(matrix, point)
        _marker(collection, location)
        context.scene.cursor.location = location
        _look_at(context, location)
    return outcome
