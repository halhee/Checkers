import bpy

from . import engine


class CHECKERS_UL_anomalies(bpy.types.UIList):
    def draw_item(self, context, layout, data, item, icon, active_data, active_propname, index):
        row = layout.row(align=True)
        row.label(text=f"{item.ifc_class} · {item.floor}", icon="ERROR")
        row.label(text=item.pset)
        row.label(text=str(item.count))
        op = row.operator("checkers.select_anomaly", text="", icon="RESTRICT_SELECT_OFF")
        op.index = index


class CHECKERS_PT_main(bpy.types.Panel):
    bl_label = "Checkers"
    bl_idname = "CHECKERS_PT_main"
    bl_space_type = "VIEW_3D"
    bl_region_type = "UI"
    bl_category = "Checkers"

    def draw(self, context):
        layout = self.layout
        props = context.scene.checkers

        if engine.current_ifc() is None:
            layout.label(text="Ouvrez un projet IFC dans Bonsai.", icon="INFO")
            return

        col = layout.column(align=True)
        col.prop(props, "rules_path")
        col.prop(props, "ids_path")
        col.separator()
        col.prop(props, "lot_name", placeholder=engine.default_lot_name())
        # Liste déroulante seulement si le fichier Excel déclare des corps d'état.
        if props.rules_path and engine.load_corps_etat(bpy.path.abspath(props.rules_path)):
            col.prop(props, "corps_etat")
        layout.operator("checkers.run_analysis", icon="CHECKMARK")

        if not props.has_result:
            return

        box = layout.box()
        if props.total_elements:
            rate = props.valid_elements / props.total_elements
            box.label(text=f"Conformité PSet : {rate:.0%} ({props.valid_elements}/{props.total_elements})",
                      icon="CHECKMARK" if rate >= 0.98 else "ERROR")
        elif props.rules_path:
            box.label(text="Aucun élément contrôlé par les règles Excel", icon="INFO")
        if props.rules_path and not props.has_rules:
            box.label(text=f"Lot sans règles : {props.lot_description}", icon="ERROR")
            box.label(text="→ choisissez le corps d'état ci-dessus")
        elif props.lot_description:
            box.label(text=f"Règles : {props.lot_description}")
        if props.ids_pass_rate >= 0:
            box.label(text=f"Conformité IDS : {props.ids_pass_rate:.0%}")
        box.label(text=f"Bilan carbone : {props.carbon_total:,.0f} kg CO₂e".replace(",", " "))

        layout.label(text=f"Anomalies ({len(props.anomalies)})")
        layout.template_list("CHECKERS_UL_anomalies", "", props, "anomalies", props, "anomaly_index", rows=6)

        row = layout.row(align=True)
        row.operator("checkers.color_results", icon="COLOR")
        row.operator("checkers.reset_colors", icon="LOOP_BACK")
        layout.operator("checkers.export_report", icon="EXPORT")


classes = (CHECKERS_UL_anomalies, CHECKERS_PT_main)
