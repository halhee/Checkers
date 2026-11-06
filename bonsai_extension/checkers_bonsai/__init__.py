"""Checkers – recette de maquettes IFC dans Bonsai.

Contrôle le modèle ouvert dans Bonsai avec un IDS (cahier des charges BIM)
et/ou un fichier Excel de PSet requis, produit une fiche de recette à signer
et un export BCF des anomalies. Fonctionne entièrement hors ligne.

Panneau : vue 3D > barre latérale (N) > onglet « Checkers ».
"""


def _classes():
    from . import operators, props, ui
    return props.classes + operators.classes + ui.classes


def register():
    import bpy
    from .props import CheckersProperties
    for cls in _classes():
        bpy.utils.register_class(cls)
    bpy.types.Scene.checkers = bpy.props.PointerProperty(type=CheckersProperties)


def unregister():
    import bpy
    del bpy.types.Scene.checkers
    for cls in reversed(_classes()):
        bpy.utils.unregister_class(cls)
