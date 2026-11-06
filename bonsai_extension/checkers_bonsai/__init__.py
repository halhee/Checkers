"""Checkers pour Bonsai : contrôle du modèle IFC ouvert (règles Excel,
IDS, bilan carbone) avec le même moteur que l'application web Checkers.

Panneau : vue 3D > barre latérale (N) > onglet « Checkers ».
"""
import bpy

from . import engine, operators, props, ui

_classes = props.classes + operators.classes + ui.classes


def register():
    engine.ensure_lib_on_path()
    for cls in _classes:
        bpy.utils.register_class(cls)
    bpy.types.Scene.checkers = bpy.props.PointerProperty(type=props.CheckersProperties)


def unregister():
    del bpy.types.Scene.checkers
    for cls in reversed(_classes):
        bpy.utils.unregister_class(cls)
    engine.remove_lib_from_path()
