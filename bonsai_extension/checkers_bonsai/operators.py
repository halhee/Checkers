import os
import shutil

import bpy
from bpy_extras.io_utils import ExportHelper

from . import engine

# Résultats non persistés dans le .blend (gros volumes, recalculables) :
# statut de chaque élément pour la coloration, et couleurs d'origine des
# objets pour pouvoir les restaurer.
_last_status = {"valid": set(), "invalid": set()}
_original_colors = {}
_original_color_types = {}

COLOR_VALID = (0.30, 0.75, 0.35, 1.0)
COLOR_INVALID = (0.90, 0.20, 0.20, 1.0)


def _objects_for_guids(guids):
    """Objets Blender des GlobalId donnés. Retourne (objets, nb introuvables) :
    un élément peut ne pas avoir d'objet (non chargé, filtré à l'ouverture)."""
    tool = engine.bonsai_tool()
    ifc_file = engine.current_ifc()
    objects, missing = [], 0
    if tool is None or ifc_file is None:
        return objects, len(guids)
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


def _view3d_spaces(context):
    for window in context.window_manager.windows:
        for area in window.screen.areas:
            if area.type == "VIEW_3D":
                for space in area.spaces:
                    if space.type == "VIEW_3D":
                        yield space


class CHECKERS_OT_run_analysis(bpy.types.Operator):
    bl_idname = "checkers.run_analysis"
    bl_label = "Lancer l'analyse"
    bl_description = "Contrôle le modèle IFC ouvert avec les règles Excel et/ou IDS choisies"

    @classmethod
    def poll(cls, context):
        return engine.current_ifc() is not None

    def execute(self, context):
        props = context.scene.checkers
        wm = context.window_manager
        wm.progress_begin(0, 1)
        try:
            result = engine.run_analysis(
                engine.current_ifc(),
                lot_name=props.lot_name,
                rules_path=bpy.path.abspath(props.rules_path) if props.rules_path else "",
                ids_path=bpy.path.abspath(props.ids_path) if props.ids_path else "",
                corps_etat_code=props.corps_etat,
            )
        except Exception as exc:
            self.report({"ERROR"}, f"Analyse impossible : {exc}")
            return {"CANCELLED"}
        finally:
            wm.progress_end()

        valid, invalid = engine.element_status(result)
        _last_status["valid"], _last_status["invalid"] = valid, invalid

        props.has_result = True
        props.total_elements = result["total_elements"]
        props.valid_elements = result["valid_elements"]
        props.report_path = result["report_path"]
        props.carbon_total = result["carbon_footprint"]["total"]
        lot = result["lots"][0]
        props.lot_description = lot["corps_etat"]
        props.has_rules = lot["has_rules"]
        ids = result.get("ids")
        props.ids_pass_rate = ids["global_pass_rate"] if ids and ids["global_pass_rate"] is not None else -1.0

        props.anomalies.clear()
        for anomaly in sorted(result["anomalies"], key=lambda a: -a["count"]):
            item = props.anomalies.add()
            item.floor = str(anomaly["floor"])
            item.ifc_class = anomaly["ifc_class"]
            item.pset = anomaly["pset"]
            item.count = anomaly["count"]
            item.guids = ";".join(anomaly.get("guids", []))
        props.anomaly_index = 0

        self.report({"INFO"}, f"Analyse terminée en {result['duration_s']:.1f} s : "
                              f"{len(result['anomalies'])} anomalie(s)")
        return {"FINISHED"}


class CHECKERS_OT_select_anomaly(bpy.types.Operator):
    bl_idname = "checkers.select_anomaly"
    bl_label = "Sélectionner"
    bl_description = "Sélectionne dans la vue 3D les éléments concernés par cette anomalie"
    bl_options = {"REGISTER", "UNDO"}

    index: bpy.props.IntProperty(default=-1)
    frame: bpy.props.BoolProperty(name="Cadrer", default=True)

    def execute(self, context):
        props = context.scene.checkers
        index = self.index if self.index >= 0 else props.anomaly_index
        if not (0 <= index < len(props.anomalies)):
            return {"CANCELLED"}
        guids = [g for g in props.anomalies[index].guids.split(";") if g]
        objects, missing = _objects_for_guids(guids)

        for obj in context.selected_objects:
            obj.select_set(False)
        selectable = [o for o in objects if o.visible_get()]
        for obj in selectable:
            obj.select_set(True)
        if selectable:
            context.view_layer.objects.active = selectable[0]
            if self.frame and context.area and context.area.type == "VIEW_3D":
                # Le bouton est dans la barre latérale : le cadrage doit viser
                # la région principale de la vue 3D, pas la barre elle-même.
                region = next((r for r in context.area.regions if r.type == "WINDOW"), None)
                if region is not None:
                    with context.temp_override(area=context.area, region=region):
                        bpy.ops.view3d.view_selected()

        hidden = len(objects) - len(selectable)
        message = f"{len(selectable)} élément(s) sélectionné(s)"
        if hidden:
            message += f", {hidden} masqué(s)"
        if missing:
            message += f", {missing} sans objet dans la scène"
        self.report({"INFO"} if selectable else {"WARNING"}, message)
        return {"FINISHED"}


class CHECKERS_OT_color_results(bpy.types.Operator):
    bl_idname = "checkers.color_results"
    bl_label = "Colorer"
    bl_description = "Colore le modèle : vert = conforme, rouge = non conforme (affichage par couleur d'objet)"

    @classmethod
    def poll(cls, context):
        return bool(_last_status["valid"] or _last_status["invalid"])

    def execute(self, context):
        for status, color in (("valid", COLOR_VALID), ("invalid", COLOR_INVALID)):
            objects, _ = _objects_for_guids(_last_status[status])
            for obj in objects:
                _original_colors.setdefault(obj.name, tuple(obj.color))
                obj.color = color
        for space in _view3d_spaces(context):
            _original_color_types.setdefault(space.as_pointer(), space.shading.color_type)
            space.shading.color_type = "OBJECT"
        return {"FINISHED"}


class CHECKERS_OT_reset_colors(bpy.types.Operator):
    bl_idname = "checkers.reset_colors"
    bl_label = "Réinitialiser"
    bl_description = "Restaure les couleurs et le mode d'affichage d'origine"

    def execute(self, context):
        for name, color in _original_colors.items():
            obj = bpy.data.objects.get(name)
            if obj is not None:
                obj.color = color
        _original_colors.clear()
        for space in _view3d_spaces(context):
            original = _original_color_types.get(space.as_pointer())
            if original is not None:
                space.shading.color_type = original
        _original_color_types.clear()
        return {"FINISHED"}


class CHECKERS_OT_export_report(bpy.types.Operator, ExportHelper):
    bl_idname = "checkers.export_report"
    bl_label = "Enregistrer le rapport Excel"
    bl_description = "Enregistre le rapport complet (Résumé, Détails, Carbone, IDS, Fiche de Recette)"

    filename_ext = ".xlsx"
    filter_glob: bpy.props.StringProperty(default="*.xlsx", options={"HIDDEN"})

    @classmethod
    def poll(cls, context):
        path = context.scene.checkers.report_path
        return bool(path) and os.path.isfile(path)

    def execute(self, context):
        shutil.copyfile(context.scene.checkers.report_path, self.filepath)
        self.report({"INFO"}, f"Rapport enregistré : {self.filepath}")
        return {"FINISHED"}


classes = (
    CHECKERS_OT_run_analysis,
    CHECKERS_OT_select_anomaly,
    CHECKERS_OT_color_results,
    CHECKERS_OT_reset_colors,
    CHECKERS_OT_export_report,
)
