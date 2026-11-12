import bpy

from . import bonsai_link, operators
from .props import registry_for

VERDICT_ICONS = {"Validé": "CHECKMARK", "Validé avec réserves": "ERROR", "Rejeté": "CANCEL"}
SOURCE_ICONS = {"IDS": "FILE_TEXT", "PSet": "PROPERTIES", "Structure": "OUTLINER"}


class CHECKERS_UL_anomalies(bpy.types.UIList):
    def draw_item(self, context, layout, data, item, icon, active_data, active_propname, index):
        row = layout.row(align=True)
        row.label(text=str(item.count))
        row.label(text=f"{item.ifc_class} · {item.level}")
        requirement = f"{item.requirement} ({item.severity.lower()})" if item.severity else item.requirement
        row.label(text=requirement, icon=SOURCE_ICONS.get(item.source, "DOT"))
        row.operator("checkers.select_anomaly", text="", icon="RESTRICT_SELECT_OFF").index = index


class _Panel:
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "Checkers"


class CHECKERS_PT_main(_Panel, bpy.types.Panel):
    bl_label = "Recette IFC"
    bl_idname = "CHECKERS_PT_main"

    def draw(self, context):
        if bonsai_link.current_ifc() is None:
            self.layout.label(text="Ouvrez un projet IFC dans Bonsai.", icon="INFO")


class _SubPanel(_Panel):
    bl_parent_id = "CHECKERS_PT_main"

    @classmethod
    def poll(cls, context):
        return bonsai_link.current_ifc() is not None


class CHECKERS_PT_referentials(_SubPanel, bpy.types.Panel):
    bl_label = "Contrôles"
    bl_idname = "CHECKERS_PT_referentials"

    def draw(self, context):
        props = context.scene.checkers
        col = self.layout.column(align=True)
        col.prop(props, "check_structure")
        col.separator()
        row = col.row(align=True)
        row.prop(props, "ids_path")
        row.operator_menu_enum("checkers.use_example_ids", "example", text="", icon="PRESET")
        row = col.row(align=True)
        row.prop(props, "rules_path")
        row.operator("checkers.use_example_rules", text="", icon="PRESET")
        if props.rules_path and registry_for(props.rules_path):
            col.prop(props, "corps_etat")


class CHECKERS_PT_recette(_SubPanel, bpy.types.Panel):
    bl_label = "Recette"
    bl_idname = "CHECKERS_PT_recette"

    def draw(self, context):
        props = context.scene.checkers
        col = self.layout.column(align=True)
        col.prop(props, "lot")
        col.prop(props, "phase")
        col.prop(props, "tester")
        col.prop(props, "marking")
        col.separator()
        col.prop(props, "previous_path")
        col.separator()
        row = col.row(align=True)
        row.prop(props, "threshold_ok", text="Validé ≥")
        row.prop(props, "threshold_reserve", text="Réserves ≥")
        self.layout.operator("checkers.run_recette", icon="CHECKMARK")


class _ResultPanel(_SubPanel):
    @classmethod
    def poll(cls, context):
        return super().poll(context) and context.scene.checkers.has_result


class CHECKERS_PT_results(_ResultPanel, bpy.types.Panel):
    bl_label = "Résultats"
    bl_idname = "CHECKERS_PT_results"

    def draw(self, context):
        props = context.scene.checkers
        layout = self.layout
        if operators.is_stale():
            layout.label(text="Le modèle a changé : relancez la recette.", icon="ERROR")
        box = layout.box()
        box.label(text=f"Verdict : {props.verdict}", icon=VERDICT_ICONS.get(props.verdict, "QUESTION"))
        for line in filter(None, props.summary.split("\n")):
            box.label(text=line)
        box.label(text=f"{props.invalid_count} élément(s) avec au moins une anomalie")
        if props.evolution:
            box.label(text=props.evolution, icon="TIME")
        for warning in filter(None, props.warnings.split("\n")):
            box.label(text=warning, icon="ERROR")
        if props.anomalies:
            layout.template_list("CHECKERS_UL_anomalies", "", props, "anomalies", props, "anomaly_index", rows=6)
        row = layout.row(align=True)
        row.operator("checkers.color_results", icon="COLOR")
        row.operator("checkers.reset_colors", icon="LOOP_BACK")


class CHECKERS_PT_deliverables(_ResultPanel, bpy.types.Panel):
    bl_label = "Livrables de recette"
    bl_idname = "CHECKERS_PT_deliverables"

    def draw(self, context):
        props = context.scene.checkers
        col = self.layout.column(align=True)
        col.operator("checkers.export_report", icon="FILE_TICK")
        col.separator()
        col.prop(props, "bcf_author")
        col.operator("checkers.export_bcf", icon="EXPORT")


class CHECKERS_PT_operation(_SubPanel, bpy.types.Panel):
    bl_label = "Exploitation"
    bl_idname = "CHECKERS_PT_operation"
    bl_options = {"DEFAULT_CLOSED"}

    def draw(self, context):
        self.layout.operator("checkers.export_inventory", icon="SPREADSHEET")


classes = (
    CHECKERS_UL_anomalies,
    CHECKERS_PT_main,
    CHECKERS_PT_referentials,
    CHECKERS_PT_recette,
    CHECKERS_PT_results,
    CHECKERS_PT_deliverables,
    CHECKERS_PT_operation,
)
