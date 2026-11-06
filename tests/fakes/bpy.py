"""Doublure minimale de bpy pour tester l'extension hors de Blender."""
import types as _t


class _Base:
    def report(self, level, message):
        self.reports = getattr(self, "reports", []) + [(set(level), message)]


types = _t.SimpleNamespace(Operator=_Base, Panel=_Base, UIList=_Base, PropertyGroup=object,
                           Scene=type("Scene", (), {}))


def _prop(*args, **kwargs):
    return ("prop", kwargs)


props = _t.SimpleNamespace(**{name: _prop for name in (
    "StringProperty", "IntProperty", "FloatProperty", "BoolProperty", "EnumProperty",
    "CollectionProperty", "PointerProperty")})
registered = []
utils = _t.SimpleNamespace(register_class=registered.append, unregister_class=registered.remove)
path = _t.SimpleNamespace(abspath=lambda p: p)
data = _t.SimpleNamespace(objects={})
view_selected_calls = []
ops = _t.SimpleNamespace(view3d=_t.SimpleNamespace(view_selected=lambda: view_selected_calls.append(1)))
