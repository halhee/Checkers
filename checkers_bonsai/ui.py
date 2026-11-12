import bpy

from . import bonsai_link, operators
from .props import registry_for

_known_icons = None


def _icon(name: str) -> str:
    """`name` si cette icône existe dans la version de Blender utilisée,
    sinon « NONE ». Les noms d'icônes changent parfois d'une version à
    l'autre, et un nom inconnu empêcherait tout le panneau de s'afficher."""
    global _known_icons
    if _known_icons is None:
        try:
            _known_icons = set(bpy.types.UILayout.bl_rna.functions["label"].parameters["icon"].enum_items.keys())
        except Exception:
            _known_icons = set()
    return name if not _known_icons or name in _known_icons else "NONE"


VERDICT_ICONS = {"Validé": "CHECKMARK", "Validé avec réserves": "ERROR", "Rejeté": "CANCEL"}
SOURCE_ICONS = {"IDS": "FILE_TEXT", "PSet": "PROPERTIES", "Structure": "OUTLINER", "Terrain": "VIEWZOOM",
                "Cohérence": "LINKED", "Conflit": "MOD_BOOLEAN"}


class CHECKERS_UL_anomalies(bpy.types.UIList):
    def draw_item(self, context, layout, data, item, icon, active_data, active_propname, index):
        row = layout.row(align=True)
        row.label(text=str(item.count))
        row.label(text=f"{item.ifc_class} · {item.level}")
        requirement = f"{item.requirement} ({item.severity.lower()})" if item.severity else item.requirement
        row.label(text=requirement, icon=_icon(SOURCE_ICONS.get(item.source, "DOT")))
        row.operator("checkers.select_anomaly", text="", icon=_icon("RESTRICT_SELECT_OFF")).index = index
        row.operator("checkers.admit_anomaly", text="", icon=_icon("CHECKBOX_HLT")).index = index


class CHECKERS_UL_admitted(bpy.types.UIList):
    def draw_item(self, context, layout, data, item, icon, active_data, active_propname, index):
        row = layout.row(align=True)
        row.label(text=str(item.count))
        row.label(text=f"{item.requirement} · {item.ifc_class} · {item.level}",
                  icon=_icon(SOURCE_ICONS.get(item.source, "DOT")))
        row.label(text=item.reason)
        row.operator("checkers.revoke_derogation", text="", icon=_icon("X")).index = index


class CHECKERS_UL_lots(bpy.types.UIList):
    def draw_item(self, context, layout, data, item, icon, active_data, active_propname, index):
        row = layout.row(align=True)
        row.prop(item, "name", text="", emboss=False, placeholder="Nom du lot")
        row.prop(item, "path", text="")


class CHECKERS_UL_clash_rules(bpy.types.UIList):
    def draw_item(self, context, layout, data, item, icon, active_data, active_propname, index):
        layout.label(text=_rule_label(item), icon=_icon("MOD_BOOLEAN"))


def _rule_label(rule) -> str:
    from .core.clashes import ClashRule
    return ClashRule(other=rule.other or "Ce modèle", mode=rule.mode, distance=rule.distance).label()


class _Panel:
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "Checkers"


class CHECKERS_PT_main(_Panel, bpy.types.Panel):
    bl_label = "Recette IFC"
    bl_idname = "CHECKERS_PT_main"

    def draw(self, context):
        if bonsai_link.current_ifc() is None:
            self.layout.label(text="Ouvrez un projet IFC dans Bonsai.", icon=_icon("INFO"))


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
        row.operator_menu_enum("checkers.use_example_ids", "example", text="", icon=_icon("PRESET"))
        row = col.row(align=True)
        row.prop(props, "rules_path")
        row.operator("checkers.use_example_rules", text="", icon=_icon("PRESET"))
        if props.rules_path and registry_for(props.rules_path):
            col.prop(props, "corps_etat")
        col.prop(props, "survey_path")


class CHECKERS_PT_coordination(_SubPanel, bpy.types.Panel):
    bl_label = "Coordination entre lots"
    bl_idname = "CHECKERS_PT_coordination"
    bl_options = {"DEFAULT_CLOSED"}

    def draw(self, context):
        props = context.scene.checkers
        layout = self.layout
        layout.label(text="Maquettes des autres lots :")
        row = layout.row()
        row.template_list("CHECKERS_UL_lots", "", props, "lots", props, "lot_index", rows=3)
        col = row.column(align=True)
        col.operator("checkers.add_lot", text="", icon=_icon("ADD"))
        col.operator("checkers.remove_lot", text="", icon=_icon("REMOVE"))

        col = layout.column(align=True)
        col.prop(props, "check_federation")
        col.prop(props, "check_clashes")
        if not props.check_clashes:
            return
        row = layout.row()
        row.template_list("CHECKERS_UL_clash_rules", "", props, "clash_rules", props, "clash_rule_index", rows=3)
        col = row.column(align=True)
        col.operator("checkers.add_clash_rule", text="", icon=_icon("ADD"))
        col.operator("checkers.remove_clash_rule", text="", icon=_icon("REMOVE"))
        layout.operator("checkers.default_clash_rules", icon=_icon("LINKED"))
        if 0 <= props.clash_rule_index < len(props.clash_rules):
            rule = props.clash_rules[props.clash_rule_index]
            box = layout.box()
            col = box.column(align=True)
            col.prop(rule, "other")
            col.prop(rule, "selector_a")
            col.prop(rule, "selector_b")
            col.prop(rule, "mode")
            if rule.mode != "collision":
                col.prop(rule, "distance")
            col.prop(rule, "severity")


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
        col.prop(props, "derogations_path")
        col.separator()
        row = col.row(align=True)
        row.prop(props, "threshold_ok", text="Validé ≥")
        row.prop(props, "threshold_reserve", text="Réserves ≥")
        self.layout.operator("checkers.run_recette", icon=_icon("CHECKMARK"))


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
            layout.label(text="Le modèle a changé : relancez la recette.", icon=_icon("ERROR"))
        box = layout.box()
        box.label(text=f"Verdict : {props.verdict}", icon=_icon(VERDICT_ICONS.get(props.verdict, "QUESTION")))
        for line in filter(None, props.summary.split("\n")):
            box.label(text=line)
        box.label(text=f"{props.invalid_count} élément(s) avec au moins une anomalie")
        if props.evolution:
            box.label(text=props.evolution, icon=_icon("TIME"))
        for warning in filter(None, props.warnings.split("\n")):
            box.label(text=warning, icon=_icon("ERROR"))
        if props.anomalies:
            layout.template_list("CHECKERS_UL_anomalies", "", props, "anomalies", props, "anomaly_index", rows=6)
        if props.conflict_position:
            row = layout.row(align=True)
            row.operator("checkers.step_conflict", text="", icon=_icon("TRIA_LEFT")).direction = -1
            row.label(text=props.conflict_position, icon=_icon("MOD_BOOLEAN"))
            row.operator("checkers.step_conflict", text="", icon=_icon("TRIA_RIGHT")).direction = 1
            row.operator("checkers.clear_conflict_view", icon=_icon("TRASH"))
        if props.admitted:
            layout.label(text=f"Admises par dérogation ({len(props.admitted)}) : hors verdict",
                         icon=_icon("CHECKBOX_HLT"))
            layout.template_list("CHECKERS_UL_admitted", "", props, "admitted", props, "admitted_index", rows=3)
        row = layout.row(align=True)
        row.operator("checkers.color_results", icon=_icon("COLOR"))
        row.operator("checkers.reset_colors", icon=_icon("LOOP_BACK"))


class CHECKERS_PT_deliverables(_ResultPanel, bpy.types.Panel):
    bl_label = "Livrables de recette"
    bl_idname = "CHECKERS_PT_deliverables"

    def draw(self, context):
        props = context.scene.checkers
        col = self.layout.column(align=True)
        col.operator("checkers.export_report", icon=_icon("FILE_TICK"))
        col.separator()
        col.prop(props, "bcf_author")
        col.operator("checkers.export_bcf", icon=_icon("EXPORT"))


class CHECKERS_PT_operation(_SubPanel, bpy.types.Panel):
    bl_label = "Exploitation et terrain"
    bl_idname = "CHECKERS_PT_operation"
    bl_options = {"DEFAULT_CLOSED"}

    def draw(self, context):
        col = self.layout.column(align=True)
        col.operator("checkers.export_inventory", icon=_icon("SPREADSHEET"))
        col.operator("checkers.export_survey", icon=_icon("VIEWZOOM"))


class CHECKERS_PT_tracking(_Panel, bpy.types.Panel):
    """Disponible sans modèle ouvert : ne lit que des journaux de recette."""
    bl_label = "Suivi de projet"
    bl_idname = "CHECKERS_PT_tracking"
    bl_parent_id = "CHECKERS_PT_main"
    bl_options = {"DEFAULT_CLOSED"}

    def draw(self, context):
        col = self.layout.column(align=True)
        col.prop(context.scene.checkers, "journals_dir")
        col.operator("checkers.export_tracking", icon=_icon("LINENUMBERS_ON"))


class CHECKERS_PT_diagnostic(_Panel, bpy.types.Panel):
    bl_label = "Installation"
    bl_idname = "CHECKERS_PT_diagnostic"
    bl_parent_id = "CHECKERS_PT_main"
    bl_options = {"DEFAULT_CLOSED"}

    def draw(self, context):
        layout = self.layout
        layout.operator("checkers.diagnostic", icon=_icon("TOOL_SETTINGS"))
        for line in filter(None, context.scene.checkers.diagnostic.split("\n")):
            layout.label(text=line[2:], icon=_icon("CHECKMARK") if line.startswith("✓") else "ERROR")


classes = (
    CHECKERS_UL_anomalies,
    CHECKERS_UL_admitted,
    CHECKERS_UL_lots,
    CHECKERS_UL_clash_rules,
    CHECKERS_PT_main,
    CHECKERS_PT_referentials,
    CHECKERS_PT_coordination,
    CHECKERS_PT_recette,
    CHECKERS_PT_results,
    CHECKERS_PT_deliverables,
    CHECKERS_PT_operation,
    CHECKERS_PT_tracking,
    CHECKERS_PT_diagnostic,
)
