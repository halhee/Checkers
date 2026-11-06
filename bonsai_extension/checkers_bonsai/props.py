import bpy

from . import engine

# Blender exige que les chaînes renvoyées par un callback d'EnumProperty
# restent référencées côté Python : sans cette liste, l'affichage de la
# liste déroulante peut montrer des caractères aléatoires.
_corps_etat_items = []


def corps_etat_items(self, context):
    registry = engine.load_corps_etat(bpy.path.abspath(self.rules_path)) if self.rules_path else None
    _corps_etat_items[:] = [(engine.AUTO, "Auto", "Correspondance automatique avec le nom du lot")]
    for code, nom in (registry or {}).items():
        _corps_etat_items.append((code, f"{code} – {nom}", nom))
    return _corps_etat_items


class CheckersAnomaly(bpy.types.PropertyGroup):
    lot: bpy.props.StringProperty(name="Lot")
    floor: bpy.props.StringProperty(name="Étage")
    ifc_class: bpy.props.StringProperty(name="Classe IFC")
    pset: bpy.props.StringProperty(name="PSet / spécification")
    count: bpy.props.IntProperty(name="Éléments")
    guids: bpy.props.StringProperty(name="GlobalId", options={"HIDDEN"})


class CheckersProperties(bpy.types.PropertyGroup):
    rules_path: bpy.props.StringProperty(
        name="Règles Excel", subtype="FILE_PATH",
        description="Fichier Excel de paramètres requis (Element_Types + un onglet par PSet)")
    ids_path: bpy.props.StringProperty(
        name="IDS", subtype="FILE_PATH",
        description="Information Delivery Specification (.ids), optionnel")
    lot_name: bpy.props.StringProperty(
        name="Lot", description="Nom du lot (par défaut : nom du fichier IFC)")
    corps_etat: bpy.props.EnumProperty(
        name="Corps d'état", items=corps_etat_items,
        description="Corps d'état dont les règles s'appliquent (onglet Corps_Etat du fichier Excel)")

    bcf_author: bpy.props.StringProperty(
        name="Auteur BCF", description="Auteur des sujets BCF (adresse e-mail en général)")

    has_result: bpy.props.BoolProperty(default=False)
    total_elements: bpy.props.IntProperty()
    valid_elements: bpy.props.IntProperty()
    carbon_total: bpy.props.FloatProperty()
    ids_pass_rate: bpy.props.FloatProperty(default=-1.0)
    lot_description: bpy.props.StringProperty()
    has_rules: bpy.props.BoolProperty(default=True)
    report_path: bpy.props.StringProperty(subtype="FILE_PATH")
    rules_description: bpy.props.StringProperty()

    anomalies: bpy.props.CollectionProperty(type=CheckersAnomaly)
    anomaly_index: bpy.props.IntProperty()


classes = (CheckersAnomaly, CheckersProperties)
