"""Accès au modèle ouvert dans Bonsai et aux objets Blender correspondants."""
import importlib
import sys


def bonsai_tool():
    """Module `bonsai.tool`, ou None si Bonsai n'est pas chargé. Bonsai est
    une extension (`bl_ext.<dépôt>.bonsai`) qui s'enregistre aussi sous le
    nom `bonsai` ; sinon on le cherche dans sys.modules."""
    try:
        return importlib.import_module("bonsai.tool")
    except ImportError:
        pass
    for name in list(sys.modules):
        if name.endswith(".bonsai"):
            try:
                return importlib.import_module(name + ".tool")
            except ImportError:
                continue
    return None


def current_ifc():
    tool = bonsai_tool()
    if tool is None:
        return None
    try:
        return tool.Ifc.get()
    except Exception:
        return None


def current_ifc_path() -> str:
    tool = bonsai_tool()
    try:
        return (tool.Ifc.get_path() or "") if tool else ""
    except Exception:
        return ""


def objects_for_guids(guids):
    """(objets Blender, nombre d'éléments sans objet dans la scène)."""
    tool, ifc_file = bonsai_tool(), current_ifc()
    if tool is None or ifc_file is None:
        return [], len(guids)
    objects, missing = [], 0
    for guid in guids:
        try:
            obj = tool.Ifc.get_object(ifc_file.by_guid(guid))
        except Exception:
            obj = None
        if obj is None:
            missing += 1
        else:
            objects.append(obj)
    return objects, missing
