"""Doublure minimale de bpy pour tester l'extension hors de Blender."""
import types as _t


class _Base:
    def report(self, level, message):
        self.reports = getattr(self, "reports", []) + [(set(level), message)]


types = _t.SimpleNamespace(Operator=_Base, Panel=_Base, UIList=_Base, PropertyGroup=object,
                           Scene=type("Scene", (), {}))


def _prop(*args, **kwargs):
    return ("prop", kwargs)


def _ensure_ext(filepath, ext):
    return filepath if filepath.lower().endswith(ext.lower()) else filepath + ext


props = _t.SimpleNamespace(**{name: _prop for name in (
    "StringProperty", "IntProperty", "FloatProperty", "BoolProperty", "EnumProperty",
    "CollectionProperty", "PointerProperty")})
registered = []
utils = _t.SimpleNamespace(register_class=registered.append, unregister_class=registered.remove)
path = _t.SimpleNamespace(abspath=lambda p: p, ensure_ext=_ensure_ext)
class _ID:
    """Bloc de données (objet, maillage, collection) créé par l'extension."""

    def __init__(self, name, data=None):
        self.name, self.data, self.users = name, data, 0
        self.objects = _Linked()          # collections
        self.children = _Linked()
        self.vertices, self.polygons = [], []

    def from_pydata(self, vertices, edges, faces):  # maillages
        self.vertices, self.polygons = list(vertices), list(faces)

    def update(self):
        pass


class _Linked(list):
    def link(self, item):
        self.append(item)


class IDCollection(dict):
    """bpy.data.objects / meshes / collections : par nom, avec new et remove."""

    def new(self, name, data=None):
        item = _ID(name, data)
        if data is not None:
            data.users += 1           # un objet utilise son maillage
        self[name] = item
        return item

    def remove(self, item, do_unlink=True):
        self.pop(item.name, None)
        if getattr(item, "data", None) is not None and item.data.users:
            item.data.users -= 1


data = _t.SimpleNamespace(objects=IDCollection(), scenes=[], meshes=IDCollection(), collections=IDCollection())
app = _t.SimpleNamespace(handlers=_t.SimpleNamespace(load_post=[], save_pre=[], persistent=lambda function: function))
view_selected_calls = []
ops = _t.SimpleNamespace(view3d=_t.SimpleNamespace(view_selected=lambda: view_selected_calls.append(1)))
