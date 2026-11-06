import os
import re

import bpy
from bpy_extras.io_utils import ExportHelper

from . import bonsai_link
from .core import bcf, recette, report
from .props import ALL_RULES

EXAMPLES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples")

# Dernière recette (en mémoire : volumineuse et recalculable) et couleurs
# d'origine des objets, pour la réinitialisation.
_state = {"result": None, "colors": {}, "color_types": {}}

COLOR_VALID = (0.30, 0.75, 0.35, 1.0)
COLOR_INVALID = (0.90, 0.20, 0.20, 1.0)


def last_result():
    return _state["result"]


def _settings(props) -> recette.RecetteSettings:
    return recette.RecetteSettings(
        ids_path=bpy.path.abspath(props.ids_path) if props.ids_path else "",
        rules_path=bpy.path.abspath(props.rules_path) if props.rules_path else "",
        corps_etat=None if props.corps_etat in ("", ALL_RULES) else props.corps_etat,
        lot=props.lot.strip(), phase=props.phase.strip(), tester=props.tester.strip(),
        marking=props.marking_text(), threshold_ok=props.threshold_ok, threshold_reserve=props.threshold_reserve,
    )


def _default_name(extension: str) -> str:
    result = last_result()
    parts = ["Recette", result.settings.lot if result else "", result.date.strftime("%Y%m%d") if result else ""]
    name = "_".join(p for p in parts if p)
    return re.sub(r"[^\w\-]+", "_", name).strip("_") + extension


def _view3d_spaces(context):
    for window in context.window_manager.windows:
        for area in window.screen.areas:
            if area.type == "VIEW_3D":
                for space in area.spaces:
                    if space.type == "VIEW_3D":
                        yield space


class CHECKERS_OT_run_recette(bpy.types.Operator):
    bl_idname = "checkers.run_recette"
    bl_label = "Lancer la recette"
    bl_description = "Contrôle le modèle IFC ouvert avec l'IDS et/ou les règles Excel choisis"

    @classmethod
    def poll(cls, context):
        props = context.scene.checkers
        return bonsai_link.current_ifc() is not None and bool(props.ids_path or props.rules_path)

    def execute(self, context):
        props = context.scene.checkers
        wm = context.window_manager
        wm.progress_begin(0, 1)
        try:
            result = recette.run_recette(bonsai_link.current_ifc(), _settings(props),
                                         model_path=bonsai_link.current_ifc_path())
        except Exception as exc:
            self.report({"ERROR"}, f"Recette impossible : {exc}")
            return {"CANCELLED"}
        finally:
            wm.progress_end()

        _state["result"] = result
        props.has_result = True
        props.verdict = result.verdict
        props.pset_rate = -1.0 if result.pset_rate is None else result.pset_rate
        props.pset_checked = result.pset["checked"] if result.pset else 0
        props.ids_rate = -1.0 if result.ids_rate is None else result.ids_rate
        props.ids_applicable = result.ids_result["applicable_total"] if result.ids_result else 0
        props.invalid_count = len(result.invalid_guids)
        props.warnings = "\n".join(result.warnings)
        props.anomalies.clear()
        for anomaly in result.anomalies:
            item = props.anomalies.add()
            item.level, item.ifc_class = anomaly["level"], anomaly["ifc_class"]
            item.requirement, item.source = anomaly["requirement"], anomaly["source"]
            item.count, item.guids = anomaly["count"], ";".join(anomaly["guids"])
        props.anomaly_index = 0
        self.report({"INFO"}, f"Recette terminée en {result.duration_s:.1f} s : {result.verdict}, "
                              f"{len(result.anomalies)} anomalie(s)")
        return {"FINISHED"}


class CHECKERS_OT_select_anomaly(bpy.types.Operator):
    bl_idname = "checkers.select_anomaly"
    bl_label = "Sélectionner"
    bl_description = "Sélectionne et cadre dans la vue 3D les éléments concernés par cette anomalie"
    bl_options = {"REGISTER", "UNDO"}

    index: bpy.props.IntProperty(default=-1)

    def execute(self, context):
        props = context.scene.checkers
        index = self.index if self.index >= 0 else props.anomaly_index
        if not 0 <= index < len(props.anomalies):
            return {"CANCELLED"}
        props.anomaly_index = index
        objects, missing = bonsai_link.objects_for_guids([g for g in props.anomalies[index].guids.split(";") if g])
        for obj in context.selected_objects:
            obj.select_set(False)
        visible = [o for o in objects if o.visible_get()]
        for obj in visible:
            obj.select_set(True)
        if visible:
            context.view_layer.objects.active = visible[0]
            region = next((r for r in context.area.regions if r.type == "WINDOW"), None) if context.area else None
            if region is not None and context.area.type == "VIEW_3D":
                with context.temp_override(area=context.area, region=region):
                    bpy.ops.view3d.view_selected()
        message = f"{len(visible)} élément(s) sélectionné(s)"
        if len(objects) > len(visible):
            message += f", {len(objects) - len(visible)} masqué(s)"
        if missing:
            message += f", {missing} sans objet dans la scène"
        self.report({"INFO"} if visible else {"WARNING"}, message)
        return {"FINISHED"}


class CHECKERS_OT_color_results(bpy.types.Operator):
    bl_idname = "checkers.color_results"
    bl_label = "Colorer"
    bl_description = "Vert = conforme, rouge = non conforme ; les éléments non contrôlés ne changent pas"

    @classmethod
    def poll(cls, context):
        return last_result() is not None

    def execute(self, context):
        result = last_result()
        for guids, color in ((result.valid_guids, COLOR_VALID), (result.invalid_guids, COLOR_INVALID)):
            objects, _ = bonsai_link.objects_for_guids(sorted(guids))
            for obj in objects:
                _state["colors"].setdefault(obj.name, tuple(obj.color))
                obj.color = color
        for space in _view3d_spaces(context):
            _state["color_types"].setdefault(space.as_pointer(), space.shading.color_type)
            space.shading.color_type = "OBJECT"
        return {"FINISHED"}


class CHECKERS_OT_reset_colors(bpy.types.Operator):
    bl_idname = "checkers.reset_colors"
    bl_label = "Réinitialiser"
    bl_description = "Restaure les couleurs et le mode d'affichage d'origine"

    def execute(self, context):
        for name, color in _state["colors"].items():
            obj = bpy.data.objects.get(name)
            if obj is not None:
                obj.color = color
        _state["colors"].clear()
        for space in _view3d_spaces(context):
            original = _state["color_types"].get(space.as_pointer())
            if original is not None:
                space.shading.color_type = original
        _state["color_types"].clear()
        return {"FINISHED"}


class CHECKERS_OT_use_example_ids(bpy.types.Operator):
    bl_idname = "checkers.use_example_ids"
    bl_label = "IDS d'exemple"
    bl_description = "Utiliser un IDS d'exemple fourni avec l'extension (à adapter au cahier des charges du projet)"

    example: bpy.props.EnumProperty(name="Exemple", items=[
        ("exploitation_maintenance.ids", "Exploitation – maintenance (DOE / GMAO)", ""),
        ("conception_minimale.ids", "Conception – exigences minimales", ""),
    ])

    def execute(self, context):
        context.scene.checkers.ids_path = os.path.join(EXAMPLES_DIR, self.example)
        return {"FINISHED"}


class CHECKERS_OT_export_report(bpy.types.Operator, ExportHelper):
    bl_idname = "checkers.export_report"
    bl_label = "Fiche de recette (Excel)"
    bl_description = "Enregistre la fiche de recette à signer, avec le détail des non-conformités"

    filename_ext = ".xlsx"
    filter_glob: bpy.props.StringProperty(default="*.xlsx", options={"HIDDEN"})

    @classmethod
    def poll(cls, context):
        return last_result() is not None

    def invoke(self, context, event):
        self.filepath = _default_name(".xlsx")
        return ExportHelper.invoke(self, context, event)

    def execute(self, context):
        try:
            report.write_report(last_result(), self.filepath)
        except OSError as exc:
            self.report({"ERROR"}, f"Enregistrement impossible : {exc}")
            return {"CANCELLED"}
        self.report({"INFO"}, f"Fiche de recette enregistrée : {self.filepath}")
        return {"FINISHED"}


class CHECKERS_OT_export_bcf(bpy.types.Operator, ExportHelper):
    bl_idname = "checkers.export_bcf"
    bl_label = "Anomalies (BCF)"
    bl_description = ("Un sujet BCF 2.1 par anomalie, éléments concernés sélectionnés : à transmettre à la "
                      "maîtrise d'œuvre et aux entreprises")

    filename_ext = ".bcf"
    filter_glob: bpy.props.StringProperty(default="*.bcf", options={"HIDDEN"})

    @classmethod
    def poll(cls, context):
        result = last_result()
        return result is not None and bool(result.anomalies)

    def invoke(self, context, event):
        self.filepath = _default_name(".bcf")
        return ExportHelper.invoke(self, context, event)

    def execute(self, context):
        result = last_result()
        s = result.settings
        try:
            stats = bcf.write_bcf(self.filepath, result.anomalies, author=context.scene.checkers.bcf_author,
                                  project_name=result.model.get("project") or None, lot=s.lot, phase=s.phase,
                                  referentials=recette.describe_referentials(result), marking=s.marking)
        except OSError as exc:
            self.report({"ERROR"}, f"Enregistrement impossible : {exc}")
            return {"CANCELLED"}
        message = f"BCF enregistré : {stats['topics']} sujet(s), {stats['components']} élément(s)"
        if stats["skipped_guids"]:
            message += f" ({stats['skipped_guids']} identifiant(s) invalide(s) ignoré(s))"
        self.report({"INFO"}, message)
        return {"FINISHED"}


classes = (
    CHECKERS_OT_run_recette,
    CHECKERS_OT_select_anomaly,
    CHECKERS_OT_color_results,
    CHECKERS_OT_reset_colors,
    CHECKERS_OT_use_example_ids,
    CHECKERS_OT_export_report,
    CHECKERS_OT_export_bcf,
)
