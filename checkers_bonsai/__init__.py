"""Checkers – recette de maquettes IFC dans Bonsai.

Contrôle le modèle ouvert dans Bonsai (structure, IDS, PSet requis), suit
les anomalies d'une livraison à l'autre, et produit la fiche de recette à
signer, l'export BCF des anomalies et l'inventaire d'exploitation.
Fonctionne entièrement hors ligne.

Panneau : vue 3D > barre latérale (N) > onglet « Checkers ».
"""


def _classes():
    from . import operators, props, ui
    return props.classes + operators.classes + ui.classes


def register():
    import bpy
    from .operators import on_load_post
    from .props import CheckersProperties
    for cls in _classes():
        bpy.utils.register_class(cls)
    bpy.types.Scene.checkers = bpy.props.PointerProperty(type=CheckersProperties)
    if on_load_post not in bpy.app.handlers.load_post:
        bpy.app.handlers.load_post.append(on_load_post)


def unregister():
    import bpy
    from .operators import on_load_post
    if on_load_post in bpy.app.handlers.load_post:
        bpy.app.handlers.load_post.remove(on_load_post)
    del bpy.types.Scene.checkers
    for cls in reversed(_classes()):
        bpy.utils.unregister_class(cls)
