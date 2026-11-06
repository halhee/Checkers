import bpy

from . import bonsai_link
from .props import registry_for

def _pct(rate: float) -> str:
    return f"{rate * 100:.1f} %".replace(".", ",")


VERDICT_ICONS = {"Validé": "CHECKMARK", "Validé avec réserves": "ERROR", "Rejeté": "CANCEL"}


class CHECKERS_UL_anomalies(bpy.types.UIList):
    def draw_item(self, context, layout, data, item, icon, active_data, active_propname, index):
        row = layout.row(align=True)
        row.label(text=str(item.count))
        row.label(text=f"{item.ifc_class} · {item.level}")
        row.label(text=item.requirement, icon="FILE_TEXT" if item.source == "IDS" else "PROPERTIES")
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
    bl_label = "Référentiels"
    bl_idname = "CHECKERS_PT_referentials"

    def draw(self, context):
        props = context.scene.checkers
        col = self.layout.column(align=True)
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
        col.separator()
        col.prop(props, "marking")
        col.separator()
        row = col.row(align=True)
        row.prop(props, "threshold_ok", text="Validé ≥")
        row.prop(props, "threshold_reserve", text="Réserves ≥")
        self.layout.operator("checkers.run_recette", icon="CHECKMARK")


class CHECKERS_PT_results(_SubPanel, bpy.types.Panel):
    bl_label = "Résultats"
    bl_idname = "CHECKERS_PT_results"

    @classmethod
    def poll(cls, context):
        return super().poll(context) and context.scene.checkers.has_result

    def draw(self, context):
        props = context.scene.checkers
        layout = self.layout
        box = layout.box()
        box.label(text=f"Verdict : {props.verdict}", icon=VERDICT_ICONS.get(props.verdict, "QUESTION"))
        if props.ids_rate >= 0:
            box.label(text=f"IDS : {_pct(props.ids_rate)} conformes ({props.ids_applicable} applicables)")
        if props.pset_rate >= 0:
            box.label(text=f"Règles Excel : {_pct(props.pset_rate)} conformes ({props.pset_checked} éléments)")
        box.label(text=f"{props.invalid_count} élément(s) non conforme(s)")
        for warning in filter(None, props.warnings.split("\n")):
            box.label(text=warning, icon="ERROR")
        if props.anomalies:
            layout.template_list("CHECKERS_UL_anomalies", "", props, "anomalies", props, "anomaly_index", rows=6)
        row = layout.row(align=True)
        row.operator("checkers.color_results", icon="COLOR")
        row.operator("checkers.reset_colors", icon="LOOP_BACK")


class CHECKERS_PT_deliverables(_SubPanel, bpy.types.Panel):
    bl_label = "Livrables"
    bl_idname = "CHECKERS_PT_deliverables"

    @classmethod
    def poll(cls, context):
        return super().poll(context) and context.scene.checkers.has_result

    def draw(self, context):
        props = context.scene.checkers
        col = self.layout.column(align=True)
        col.operator("checkers.export_report", icon="FILE_TICK")
        col.separator()
        col.prop(props, "bcf_author")
        col.operator("checkers.export_bcf", icon="EXPORT")


classes = (
    CHECKERS_UL_anomalies,
    CHECKERS_PT_main,
    CHECKERS_PT_referentials,
    CHECKERS_PT_recette,
    CHECKERS_PT_results,
    CHECKERS_PT_deliverables,
)
