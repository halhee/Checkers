import os
import re
import weakref
from datetime import datetime

import bpy
from bpy_extras.io_utils import ExportHelper

from . import bonsai_link
from .core import bcf, journal, recette
from .core.spatial import LevelCache
from .core.trace import model_info
from .props import ALL_RULES

EXAMPLES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples")


def _missing_library(exc: ImportError) -> str:
    return f"bibliothèque {exc.name or 'requise'} introuvable (vérifiez que Bonsai est installé et activé)"


# Dernière recette, gardée en mémoire seulement : elle contient les
# identifiants de tous les éléments en anomalie, qui n'ont pas à être écrits
# dans le fichier .blend. `model` désigne le modèle contrôlé, pour détecter
# qu'un autre projet a été ouvert depuis. `colors` et `color_types` gardent
# l'affichage d'origine, pour la réinitialisation.
_state = {"result": None, "model": None, "colors": {}, "color_types": {}}

COLOR_VALID = (0.30, 0.75, 0.35, 1.0)
COLOR_INVALID = (0.90, 0.20, 0.20, 1.0)


def last_result():
    return _state["result"]


def _model_token(ifc_file):
    """Référence faible au modèle contrôlé : elle ne le retient pas en
    mémoire quand l'utilisateur ouvre un autre projet."""
    try:
        return weakref.ref(ifc_file)
    except TypeError:
        return id(ifc_file)


def is_stale() -> bool:
    """La dernière recette porte-t-elle sur un autre modèle que celui ouvert ?"""
    if _state["result"] is None:
        return False
    current, token = bonsai_link.current_ifc(), _state["model"]
    if current is None:
        return True
    return token() is not current if isinstance(token, weakref.ref) else token != id(current)


def usable_result():
    """Dernière recette, si elle porte bien sur le modèle ouvert."""
    return None if is_stale() else _state["result"]


def clear_results(props) -> None:
    props.has_result = False
    props.verdict = props.summary = props.warnings = props.evolution = ""
    props.invalid_count = 0
    props.anomalies.clear()
    props.anomaly_index = 0


@bpy.app.handlers.persistent
def on_load_post(*_):
    """À l'ouverture d'un fichier, la recette en mémoire ne correspond plus à
    rien : le panneau repart sans résultat plutôt que d'afficher un ancien
    verdict."""
    _state.update(result=None, model=None)
    _state["colors"].clear()
    _state["color_types"].clear()
    for scene in bpy.data.scenes:
        props = getattr(scene, "checkers", None)
        if props is not None:
            clear_results(props)


def _pct(rate: float) -> str:
    return f"{rate * 100:.1f} %".replace(".", ",")


def _settings(props) -> recette.RecetteSettings:
    def path(value):
        return bpy.path.abspath(value) if value else ""

    return recette.RecetteSettings(
        ids_path=path(props.ids_path), rules_path=path(props.rules_path),
        corps_etat=None if props.corps_etat in ("", ALL_RULES) else props.corps_etat,
        check_structure=props.check_structure, previous_path=path(props.previous_path),
        lot=props.lot.strip(), phase=props.phase.strip(), tester=props.tester.strip(),
        marking=props.marking_text(), threshold_ok=props.threshold_ok, threshold_reserve=props.threshold_reserve,
    )


def _summary(result) -> str:
    lines = []
    if result.structure is not None:
        s = result.structure
        found = ", ".join(f"{s[key]} {label}" for key, label in
                          (("blocking", "bloquant(s)"), ("major", "majeur(s)"), ("minor", "mineur(s)")) if s[key])
        lines.append(f"Structure : {found or 'aucun constat'}")
    if result.ids_rate is not None:
        lines.append(f"IDS : {_pct(result.ids_rate)} conformes "
                     f"({result.ids_result['applicable_total']} vérifications)")
    if result.pset_rate is not None:
        lines.append(f"Règles Excel : {_pct(result.pset_rate)} conformes ({result.pset['checked']} éléments)")
    return "\n".join(lines)


def _default_name(prefix: str, label: str, date: datetime, extension: str) -> str:
    name = "_".join(part for part in (prefix, label, date.strftime("%Y%m%d")) if part)
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
    bl_description = "Contrôle le modèle IFC ouvert : structure, IDS et règles Excel selon les choix ci-dessus"

    @classmethod
    def poll(cls, context):
        props = context.scene.checkers
        return bonsai_link.current_ifc() is not None and bool(
            props.ids_path or props.rules_path or props.check_structure)

    def execute(self, context):
        props = context.scene.checkers
        ifc_file = bonsai_link.current_ifc()
        wm = context.window_manager
        wm.progress_begin(0, 1)
        try:
            result = recette.run_recette(ifc_file, _settings(props), model_path=bonsai_link.current_ifc_path(),
                                         unsaved_changes=bonsai_link.has_unsaved_changes(context.scene))
        except Exception as exc:
            self.report({"ERROR"}, f"Recette impossible : {exc}")
            return {"CANCELLED"}
        finally:
            wm.progress_end()

        _state.update(result=result, model=_model_token(ifc_file))
        clear_results(props)
        props.has_result = True
        props.verdict = result.verdict
        props.summary = _summary(result)
        props.warnings = "\n".join(result.warnings)
        props.invalid_count = len(result.invalid_guids)
        if result.evolution is not None:
            e = result.evolution
            props.evolution = (f"Depuis le {e['previous_date']:%d/%m/%Y} : {e['resolved_count']} levée(s), "
                               f"{e['persistent_count']} persistante(s), {e['new_count']} nouvelle(s)")
        for anomaly in result.anomalies:
            item = props.anomalies.add()
            item.level, item.ifc_class = anomaly["level"], anomaly["ifc_class"]
            item.requirement, item.source = anomaly["requirement"], anomaly["source"]
            item.severity, item.count = anomaly["severity"] or "", anomaly["count"]
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
        result = usable_result()
        if result is None:
            self.report({"WARNING"}, "Le modèle a changé depuis la recette : relancez-la.")
            return {"CANCELLED"}
        index = self.index if self.index >= 0 else props.anomaly_index
        if not 0 <= index < len(result.anomalies):
            return {"CANCELLED"}
        props.anomaly_index = index
        guids = result.anomalies[index]["guids"]
        if not guids:
            self.report({"INFO"}, "Ce constat porte sur le modèle entier, pas sur des éléments précis.")
            return {"FINISHED"}
        objects, missing = bonsai_link.objects_for_guids(guids)
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
    bl_description = ("Rouge = au moins une anomalie, vert = contrôlé sans anomalie ; "
                      "les éléments non contrôlés ne changent pas")

    @classmethod
    def poll(cls, context):
        return usable_result() is not None

    def execute(self, context):
        result = usable_result()
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


class CHECKERS_OT_use_example_rules(bpy.types.Operator):
    bl_idname = "checkers.use_example_rules"
    bl_label = "Règles d'exemple"
    bl_description = "Utiliser le fichier de règles Excel d'exemple fourni avec l'extension (multi-corps d'état)"

    def execute(self, context):
        context.scene.checkers.rules_path = os.path.join(EXAMPLES_DIR, "regles_exemple.xlsx")
        return {"FINISHED"}


class CHECKERS_OT_export_report(bpy.types.Operator, ExportHelper):
    bl_idname = "checkers.export_report"
    bl_label = "Fiche de recette (Excel) + journal"
    bl_description = ("Enregistre la fiche de recette à signer, avec le détail des non-conformités, et le journal "
                      "(.json, même nom) à fournir comme « recette précédente » à la livraison suivante")

    filename_ext = ".xlsx"
    filter_glob: bpy.props.StringProperty(default="*.xlsx", options={"HIDDEN"})

    @classmethod
    def poll(cls, context):
        return usable_result() is not None

    def invoke(self, context, event):
        result = last_result()
        self.filepath = _default_name("Recette", result.settings.lot, result.date, ".xlsx")
        return ExportHelper.invoke(self, context, event)

    def execute(self, context):
        result = usable_result()
        if result is None:
            self.report({"ERROR"}, "Le modèle a changé depuis la recette : relancez-la.")
            return {"CANCELLED"}
        path = bpy.path.ensure_ext(self.filepath, ".xlsx")
        journal_path = os.path.splitext(path)[0] + ".json"
        try:
            # Le journal d'abord : il ne dépend d'aucune bibliothèque.
            journal.write_record(result, journal_path)
            from .core import report
            report.write_report(result, path)
        except ImportError as exc:
            self.report({"ERROR"}, f"Fiche Excel impossible : {_missing_library(exc)}. "
                                   f"Journal enregistré : {journal_path}")
            return {"CANCELLED"}
        except OSError as exc:
            self.report({"ERROR"}, f"Enregistrement impossible : {exc}")
            return {"CANCELLED"}
        self.report({"INFO"}, f"Fiche de recette et journal enregistrés : {path}")
        return {"FINISHED"}


class CHECKERS_OT_export_bcf(bpy.types.Operator, ExportHelper):
    bl_idname = "checkers.export_bcf"
    bl_label = "Anomalies (BCF)"
    bl_description = ("Un sujet BCF 2.1 par anomalie, éléments concernés sélectionnés : à transmettre à la "
                      "maîtrise d'œuvre et aux entreprises. Les anomalies levées depuis la recette précédente "
                      "sont transmises comme sujets clos")

    filename_ext = ".bcf"
    filter_glob: bpy.props.StringProperty(default="*.bcf", options={"HIDDEN"})

    @classmethod
    def poll(cls, context):
        result = usable_result()
        return result is not None and bool(result.anomalies or (result.evolution and result.evolution["resolved"]))

    def invoke(self, context, event):
        result = last_result()
        self.filepath = _default_name("Recette", result.settings.lot, result.date, ".bcf")
        return ExportHelper.invoke(self, context, event)

    def execute(self, context):
        result = usable_result()
        if result is None:
            self.report({"ERROR"}, "Le modèle a changé depuis la recette : relancez-la.")
            return {"CANCELLED"}
        s = result.settings
        try:
            stats = bcf.write_bcf(
                bpy.path.ensure_ext(self.filepath, ".bcf"), result.anomalies,
                author=context.scene.checkers.bcf_author, project_name=result.model.get("project") or None,
                project_key=result.model.get("project_guid") or "", lot=s.lot, phase=s.phase,
                referentials=recette.describe_referentials(result), marking=s.marking,
                resolved=result.evolution["resolved"] if result.evolution else ())
        except OSError as exc:
            self.report({"ERROR"}, f"Enregistrement impossible : {exc}")
            return {"CANCELLED"}
        message = f"BCF enregistré : {stats['topics'] - stats['closed']} sujet(s) ouvert(s)"
        if stats["closed"]:
            message += f", {stats['closed']} clos"
        if stats["skipped_guids"]:
            message += f" ({stats['skipped_guids']} identifiant(s) invalide(s) ignoré(s))"
        self.report({"INFO"}, message)
        return {"FINISHED"}


class CHECKERS_OT_export_inventory(bpy.types.Operator, ExportHelper):
    bl_idname = "checkers.export_inventory"
    bl_label = "Inventaire d'exploitation (Excel)"
    bl_description = ("Liste les équipements techniques et les locaux de la maquette, avec le taux de "
                      "renseignement de chaque information : à reprendre dans un outil de maintenance (GMAO)")

    filename_ext = ".xlsx"
    filter_glob: bpy.props.StringProperty(default="*.xlsx", options={"HIDDEN"})

    @classmethod
    def poll(cls, context):
        return bonsai_link.current_ifc() is not None

    def invoke(self, context, event):
        self.filepath = _default_name("Inventaire", context.scene.checkers.lot.strip(), datetime.now(), ".xlsx")
        return ExportHelper.invoke(self, context, event)

    def execute(self, context):
        ifc_file = bonsai_link.current_ifc()
        if ifc_file is None:
            self.report({"ERROR"}, "Aucun modèle IFC ouvert dans Bonsai.")
            return {"CANCELLED"}
        try:
            from .core import inventory
            data = inventory.collect(ifc_file, LevelCache())
            inventory.write_inventory(
                bpy.path.ensure_ext(self.filepath, ".xlsx"), data,
                model_info(ifc_file, bonsai_link.current_ifc_path()),
                marking=context.scene.checkers.marking_text(), date=datetime.now())
        except ImportError as exc:
            self.report({"ERROR"}, f"Inventaire impossible : {_missing_library(exc)}")
            return {"CANCELLED"}
        except OSError as exc:
            self.report({"ERROR"}, f"Enregistrement impossible : {exc}")
            return {"CANCELLED"}
        self.report({"INFO"}, f"Inventaire enregistré : {len(data['equipment'])} équipement(s), "
                              f"{len(data['spaces'])} local(aux)")
        return {"FINISHED"}


classes = (
    CHECKERS_OT_run_recette,
    CHECKERS_OT_select_anomaly,
    CHECKERS_OT_color_results,
    CHECKERS_OT_reset_colors,
    CHECKERS_OT_use_example_ids,
    CHECKERS_OT_use_example_rules,
    CHECKERS_OT_export_report,
    CHECKERS_OT_export_bcf,
    CHECKERS_OT_export_inventory,
)
