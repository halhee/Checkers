import os

import bpy

from .core import rules as rules_file

PHASES = [
    "ESQ – Esquisse", "APS – Avant-projet sommaire", "APD – Avant-projet définitif", "PRO – Projet",
    "DCE – Dossier de consultation des entreprises", "EXE – Exécution", "DOE – Dossier des ouvrages exécutés",
    "Exploitation – maintenance",
]

ALL_RULES = "ALL"

# Les chaînes d'un callback d'EnumProperty doivent rester référencées côté
# Python, sinon Blender peut afficher des caractères aléatoires.
_corps_etat_items = []
_registry_cache = {}


def registry_for(path: str):
    """Onglet Corps_Etat du fichier de règles, mis en cache par (chemin,
    date de modification) : appelé à chaque rafraîchissement du panneau."""
    path = bpy.path.abspath(path) if path else ""
    if not path or not os.path.isfile(path):
        return None
    key = (path, os.path.getmtime(path))
    if key not in _registry_cache:
        _registry_cache.clear()
        try:
            _registry_cache[key] = rules_file.load_registry(path)
        except Exception:
            _registry_cache[key] = None
    return _registry_cache[key]


def _corps_etat_enum(self, context):
    _corps_etat_items[:] = [(ALL_RULES, "Toutes les règles", "Appliquer toutes les lignes du fichier")]
    for code, nom in (registry_for(self.rules_path) or {}).items():
        _corps_etat_items.append((code, f"{code} – {nom}", nom))
    return _corps_etat_items


def _phase_search(self, context, edit_text):
    return [p for p in PHASES if edit_text.lower() in p.lower()] or PHASES


class CheckersAnomaly(bpy.types.PropertyGroup):
    """Ligne affichée dans le panneau. Les identifiants des éléments restent
    en mémoire (voir operators.py) : aucun n'est écrit dans le fichier .blend."""
    level: bpy.props.StringProperty(name="Niveau / partie")
    ifc_class: bpy.props.StringProperty(name="Classe IFC")
    requirement: bpy.props.StringProperty(name="Exigence")
    source: bpy.props.StringProperty(name="Source")
    severity: bpy.props.StringProperty(name="Gravité")
    count: bpy.props.IntProperty(name="Éléments")


class CheckersProperties(bpy.types.PropertyGroup):
    # Référentiels
    ids_path: bpy.props.StringProperty(
        name="IDS", subtype="FILE_PATH",
        description="Information Delivery Specification (.ids) : exigences du cahier des charges BIM")
    rules_path: bpy.props.StringProperty(
        name="Règles Excel", subtype="FILE_PATH",
        description="Fichier Excel de PSet requis, facultatif")
    corps_etat: bpy.props.EnumProperty(name="Corps d'état", items=_corps_etat_enum,
                                       description="Lignes du fichier Excel à appliquer")
    check_structure: bpy.props.BoolProperty(
        name="Contrôler la structure du modèle", default=True,
        description="Projet unique, unités, arborescence spatiale, éléments rattachés à un niveau, "
                    "identifiants uniques… sans référentiel")
    survey_path: bpy.props.StringProperty(
        name="Relevé terrain", subtype="FILE_PATH",
        description="Fiche de relevé terrain remplie sur site (.xlsx, produite par le panneau « Exploitation et "
                    "terrain ») : les écarts avec la maquette deviennent des anomalies")
    previous_path: bpy.props.StringProperty(
        name="Recette précédente", subtype="FILE_PATH",
        description="Journal (.json) enregistré avec la fiche de la recette précédente : permet de suivre les "
                    "anomalies levées, persistantes et nouvelles")

    # Informations de recette
    lot: bpy.props.StringProperty(name="Lot", description="Lot ou corps d'état contrôlé (ex. Lot 02 – Gros œuvre)")
    phase: bpy.props.StringProperty(name="Phase", search=_phase_search,
                                    description="Phase ou jalon de livraison (ex. DOE)")
    tester: bpy.props.StringProperty(name="Réalisée par", description="Nom et fonction de la personne qui recette")
    marking: bpy.props.StringProperty(
        name="Mention", description="Mention de confidentialité facultative, portée sur chaque page des classeurs "
                                    "et dans le BCF (ex. CONFIDENTIEL)")
    threshold_ok: bpy.props.FloatProperty(name="Seuil « validé »", default=0.98, min=0.0, max=1.0,
                                          subtype="FACTOR")
    threshold_reserve: bpy.props.FloatProperty(name="Seuil « avec réserves »", default=0.90, min=0.0, max=1.0,
                                               subtype="FACTOR")
    bcf_author: bpy.props.StringProperty(name="Auteur BCF",
                                         description="Auteur des sujets BCF (adresse électronique en général)")
    journals_dir: bpy.props.StringProperty(
        name="Dossier des journaux", subtype="DIR_PATH",
        description="Dossier où sont rangés les journaux de recette (.json), sous-dossiers compris")
    diagnostic: bpy.props.StringProperty()   # une ligne par vérification

    # Résumé de la dernière recette, pour l'affichage. Remis à zéro à
    # l'ouverture d'un fichier (voir operators.clear_results).
    has_result: bpy.props.BoolProperty(default=False)
    verdict: bpy.props.StringProperty()
    summary: bpy.props.StringProperty()      # une ligne par contrôle
    warnings: bpy.props.StringProperty()     # une ligne par avertissement
    evolution: bpy.props.StringProperty()
    invalid_count: bpy.props.IntProperty()
    anomalies: bpy.props.CollectionProperty(type=CheckersAnomaly)
    anomaly_index: bpy.props.IntProperty()

    def marking_text(self) -> str:
        return self.marking.strip().upper()


classes = (CheckersAnomaly, CheckersProperties)
