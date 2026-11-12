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


def has_unsaved_changes(scene):
    """True si Bonsai signale des modifications IFC non enregistrées, False
    s'il n'y en a pas, None si l'information n'est pas disponible. Sert à
    avertir que l'empreinte du fichier sur disque ne décrit plus exactement
    le modèle contrôlé."""
    props = getattr(scene, "BIMProperties", None)
    dirty = getattr(props, "is_dirty", None)
    return dirty if isinstance(dirty, bool) else None


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


def guid_of(obj):
    """GlobalId de l'élément IFC représenté par un objet Blender, ou None."""
    tool = bonsai_tool()
    try:
        entity = tool.Ifc.get_entity(obj) if tool else None
        return getattr(entity, "GlobalId", None) if entity is not None else None
    except Exception:
        return None


def _float(value, default: float) -> float:
    try:
        return float(value) if value not in (None, "") else default
    except (TypeError, ValueError):
        return default


def scene_offset():
    """Décalage de la scène appliqué par Bonsai aux grandes coordonnées
    (« false origin ») : ((x, y, z) en mètres, abscisse, ordonnée de l'axe
    X), ou None s'il n'y en a pas. Voir core/scene.py."""
    tool, ifc_file = bonsai_tool(), current_ifc()
    if tool is None or ifc_file is None:
        return None
    try:
        props = tool.Georeference.get_georeference_props()
        if not getattr(props, "has_blender_offset", False):
            return None
        import ifcopenshell.util.unit
        scale = ifcopenshell.util.unit.calculate_unit_scale(ifc_file)
        position = tuple(_float(getattr(props, f"blender_offset_{axis}", 0.0), 0.0) * scale for axis in "xyz")
        return (position, _float(getattr(props, "blender_x_axis_abscissa", 1.0), 1.0),
                _float(getattr(props, "blender_x_axis_ordinate", 0.0), 0.0))
    except Exception:
        return None
