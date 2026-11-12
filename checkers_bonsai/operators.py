import os
import re
import weakref
from datetime import datetime

import bpy
from bpy_extras.io_utils import ExportHelper

from . import bonsai_link, viewer
from .core import bcf, clashes, derogations, diagnostic, journal, recette
from .core.federation import CURRENT, lot_name
from .core.spatial import LevelCache
from .core.trace import file_info, model_info
from .props import ALL_RULES

EXAMPLES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "examples")


def _missing_library(exc: ImportError) -> str:
    return f"bibliothèque {exc.name or 'requise'} introuvable (vérifiez que Bonsai est installé et activé)"


# Dernière recette, gardée en mémoire seulement : elle contient les
# identifiants de tous les éléments en anomalie, qui n'ont pas à être écrits
# dans le fichier .blend. `model` désigne le modèle contrôlé, pour détecter
# qu'un autre projet a été ouvert depuis. `colors` et `color_types` gardent
# l'affichage d'origine, pour la réinitialisation.
# `conflict` : (n° de l'anomalie, n° du conflit affiché, -1 pour la vue
# d'ensemble) pendant la navigation entre conflits.
_state = {"result": None, "model": None, "colors": {}, "shading": {}, "conflict": None}

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
    props.admitted.clear()
    props.admitted_index = 0
    props.conflict_position = ""


@bpy.app.handlers.persistent
def on_load_post(*_):
    """À l'ouverture d'un fichier, la recette en mémoire ne correspond plus à
    rien : le panneau repart sans résultat plutôt que d'afficher un ancien
    verdict."""
    _state.update(result=None, model=None, conflict=None)
    _state["colors"].clear()
    _state["shading"].clear()
    viewer.forget_lots()
    for scene in bpy.data.scenes:
        props = getattr(scene, "checkers", None)
        if props is not None:
            clear_results(props)


@bpy.app.handlers.persistent
def on_save_pre(*_):
    """Les objets d'affichage d'un conflit sont temporaires : ils ne doivent
    pas être enregistrés dans le fichier .blend."""
    viewer.clear()
    _state["conflict"] = None
    for scene in bpy.data.scenes:
        props = getattr(scene, "checkers", None)
        if props is not None:
            props.conflict_position = ""


def _pct(rate: float) -> str:
    return f"{rate * 100:.1f} %".replace(".", ",")


def _settings(props) -> recette.RecetteSettings:
    def path(value):
        return bpy.path.abspath(value) if value else ""

    lots = [(lot.name.strip(), path(lot.path)) for lot in props.lots if lot.path]
    rules = [clashes.ClashRule(other=rule.other.strip() or CURRENT, selector_a=rule.selector_a.strip(),
                               selector_b=rule.selector_b.strip(), mode=rule.mode, distance=rule.distance,
                               severity=rule.severity)
             for rule in props.clash_rules] if props.check_clashes else []
    return recette.RecetteSettings(
        lots=lots, check_federation=props.check_federation and bool(lots), clash_rules=rules,
        ids_path=path(props.ids_path), rules_path=path(props.rules_path),
        corps_etat=None if props.corps_etat in ("", ALL_RULES) else props.corps_etat,
        check_structure=props.check_structure, survey_path=path(props.survey_path),
        previous_path=path(props.previous_path), derogations_path=path(props.derogations_path),
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
    if result.survey_rate is not None:
        lines.append(f"Terrain : {_pct(result.survey_rate)} sans écart "
                     f"({result.survey_result['checks']} vérifications)")
    for label, outcome in (("Cohérence entre lots", result.federation), ("Conflits", result.clashes)):
        if outcome is not None:
            found = ", ".join(f"{outcome[key]} {name}" for key, name in
                              (("blocking", "bloquant(s)"), ("major", "majeur(s)"), ("minor", "mineur(s)"),
                               ("errors", "non exécuté(s)")) if outcome[key])
            if outcome is result.clashes:
                found = f"{outcome['total']} conflit(s)" + (f" — {found}" if found else "")
            lines.append(f"{label} : {found or 'aucun constat'}")
    return "\n".join(lines)


def _show_result(props, result) -> None:
    """Recopie dans le panneau le résumé d'une recette (après la recette, ou
    quand une dérogation est accordée ou retirée)."""
    index = props.anomaly_index
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
        if e.get("admitted_count"):
            props.evolution += f", {e['admitted_count']} admise(s)"
    for anomaly in result.anomalies:
        item = props.anomalies.add()
        item.level, item.ifc_class = anomaly["level"], anomaly["ifc_class"]
        item.requirement, item.source = anomaly["requirement"], anomaly["source"]
        item.severity, item.count = anomaly["severity"] or "", anomaly["count"]
    for anomaly in result.admitted:
        item = props.admitted.add()
        item.level, item.ifc_class = anomaly["level"], anomaly["ifc_class"]
        item.requirement, item.source = anomaly["requirement"], anomaly["source"]
        item.severity, item.count = anomaly["severity"] or "", anomaly["count"]
        item.reason, item.derogation_id = anomaly["derogation"]["reason"], anomaly["derogation"]["id"]
    props.anomaly_index = min(index, max(len(result.anomalies) - 1, 0))


def _select_guids(context, guids) -> tuple:
    """Sélectionne les objets des éléments ; (visibles, objets, sans objet)."""
    objects, missing = bonsai_link.objects_for_guids(guids)
    for obj in context.selected_objects:
        obj.select_set(False)
    visible = []
    for obj in objects:
        # Un objet hors du calque de vue courant ne peut pas être
        # sélectionné : Blender lève alors une erreur.
        try:
            if obj.visible_get():
                obj.select_set(True)
                visible.append(obj)
        except RuntimeError:
            continue
    if visible:
        context.view_layer.objects.active = visible[0]
    return visible, objects, missing


def _lot_path(result, name: str) -> str:
    return dict(result.settings.lots).get(name, "") if name and name != CURRENT else ""


def _show_conflict(context, result, anomaly, position: int) -> str:
    """Affiche un conflit de l'anomalie (position ≥ 0) ou tous ses éléments
    (position -1) ; retourne le complément du message."""
    conflicts = anomaly.get("conflicts") or []
    external = anomaly.get("lot") not in (None, "", CURRENT)
    if position < 0:
        point, others = anomaly.get("point"), anomaly.get("related_guids") or []
    else:
        conflict = conflicts[position]
        point, others = conflict["point"], [conflict["b"]] if external else []
    outcome = viewer.show_conflict(context, point, _lot_path(result, anomaly.get("lot")) if external else "", others)
    message = ""
    if outcome["shown"]:
        message += f", {outcome['shown']} élément(s) du lot « {anomaly['lot']} » affiché(s) en filaire"
    if outcome["missing"]:
        message += f", {outcome['missing']} introuvable(s) dans ce lot"
    if outcome["error"]:
        message += f" ({outcome['error']})"
    return message


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
    bl_description = ("Contrôle le modèle IFC ouvert : structure, IDS, règles Excel, relevé terrain et "
                      "coordination avec les autres lots, selon les choix ci-dessus")

    @classmethod
    def poll(cls, context):
        props = context.scene.checkers
        coordination = any(lot.path for lot in props.lots) and props.check_federation or \
            props.check_clashes and len(props.clash_rules) > 0
        return bonsai_link.current_ifc() is not None and bool(
            props.ids_path or props.rules_path or props.check_structure or props.survey_path or coordination)

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

        viewer.clear()
        _state.update(result=result, model=_model_token(ifc_file), conflict=None)
        props.anomaly_index = 0
        _show_result(props, result)
        message = ""
        if result.admitted:
            message = f" ({len(result.admitted)} admise(s) par dérogation)"
        self.report({"INFO"}, f"Recette terminée en {result.duration_s:.1f} s : {result.verdict}, "
                              f"{len(result.anomalies)} anomalie(s){message}")
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
        anomaly = result.anomalies[index]
        guids = anomaly["guids"]
        viewer.clear()
        _state["conflict"], props.conflict_position = None, ""
        if not guids:
            self.report({"INFO"}, "Ce constat porte sur le modèle entier, pas sur des éléments précis.")
            return {"FINISHED"}
        visible, objects, missing = _select_guids(context, guids)
        message = f"{len(visible)} élément(s) sélectionné(s)"
        if len(objects) > len(visible):
            message += f", {len(objects) - len(visible)} masqué(s)"
        if missing:
            message += f", {missing} sans objet dans la scène"
        if anomaly.get("conflicts"):
            # Vue centrée sur le premier conflit plutôt que cadrage de tous
            # les éléments, qui peuvent être longs (gaines, réseaux).
            message += _show_conflict(context, result, anomaly, -1)
            _state["conflict"] = (index, -1)
            count = len(anomaly["conflicts"])
            props.conflict_position = f"{count} conflit(s) : flèches pour les parcourir un à un"
        elif visible:
            region = next((r for r in context.area.regions if r.type == "WINDOW"), None) if context.area else None
            if region is not None and context.area.type == "VIEW_3D":
                with context.temp_override(area=context.area, region=region):
                    bpy.ops.view3d.view_selected()
        self.report({"INFO"} if visible else {"WARNING"}, message)
        return {"FINISHED"}


class CHECKERS_OT_step_conflict(bpy.types.Operator):
    bl_idname = "checkers.step_conflict"
    bl_label = "Conflit suivant"
    bl_description = ("Passe au conflit suivant (ou précédent) de l'anomalie : sélectionne l'élément, centre la vue "
                      "sur le point de conflit et affiche l'élément de l'autre lot")
    bl_options = {"REGISTER", "UNDO"}

    direction: bpy.props.IntProperty(default=1)

    @classmethod
    def poll(cls, context):
        result = usable_result()
        index = context.scene.checkers.anomaly_index
        return result is not None and 0 <= index < len(result.anomalies) \
            and bool(result.anomalies[index].get("conflicts"))

    def execute(self, context):
        props = context.scene.checkers
        result = usable_result()
        if result is None:
            self.report({"WARNING"}, "Le modèle a changé depuis la recette : relancez-la.")
            return {"CANCELLED"}
        index = props.anomaly_index
        if not 0 <= index < len(result.anomalies) or not result.anomalies[index].get("conflicts"):
            return {"CANCELLED"}
        anomaly = result.anomalies[index]
        conflicts = anomaly["conflicts"]
        current = _state["conflict"][1] if _state["conflict"] and _state["conflict"][0] == index else -1
        if current < 0:
            position = 0 if self.direction >= 0 else len(conflicts) - 1
        else:
            position = (current + (1 if self.direction >= 0 else -1)) % len(conflicts)
        conflict = conflicts[position]
        internal = anomaly.get("lot") in (None, "", CURRENT)
        visible, _, _ = _select_guids(context, [conflict["a"], conflict["b"]] if internal else [conflict["a"]])
        message = _show_conflict(context, result, anomaly, position)
        _state["conflict"] = (index, position)
        props.conflict_position = f"Conflit {position + 1} / {len(conflicts)}"
        if anomaly.get("pairs", 0) > len(conflicts):
            props.conflict_position += f" (sur {anomaly['pairs']})"
        self.report({"INFO"}, f"{props.conflict_position} : {len(visible)} élément(s) sélectionné(s){message}")
        return {"FINISHED"}


class CHECKERS_OT_clear_conflict_view(bpy.types.Operator):
    bl_idname = "checkers.clear_conflict_view"
    bl_label = "Effacer"
    bl_description = "Efface le repère du point de conflit et les éléments de l'autre lot affichés en filaire"

    def execute(self, context):
        viewer.clear()
        _state["conflict"] = None
        context.scene.checkers.conflict_position = ""
        return {"FINISHED"}


def _derogations_target(props) -> str:
    """Fichier des dérogations : celui choisi, sinon un fichier à côté de la
    maquette (le champ est alors renseigné)."""
    if props.derogations_path:
        return bpy.path.abspath(props.derogations_path)
    model_path = bonsai_link.current_ifc_path()
    if not model_path:
        return ""
    path = derogations.default_path(model_path)
    props.derogations_path = path
    return path


def _apply_derogations(result, document, path) -> None:
    """Recalcule verdicts, couleurs et évolution sans relancer les contrôles."""
    result.settings.derogations_path = path
    result.derogation_doc = document
    result.derogations = file_info(path)
    recette.finalize(result)


class CHECKERS_OT_admit_anomaly(bpy.types.Operator):
    bl_idname = "checkers.admit_anomaly"
    bl_label = "Admettre par dérogation"
    bl_description = ("Admet cette anomalie, avec un motif : elle sort du verdict et figure sur la fiche dans "
                      "l'onglet « Dérogations ». Enregistrée dans le fichier des dérogations, elle vaut aussi pour "
                      "les recettes suivantes, pour ces éléments seulement")
    bl_options = {"REGISTER", "UNDO"}

    index: bpy.props.IntProperty(default=-1)
    reason: bpy.props.StringProperty(name="Motif", description="Pourquoi cette anomalie est admise, par qui elle a "
                                                               "été décidée (obligatoire)")
    scope: bpy.props.EnumProperty(name="Portée", items=[
        ("all", "Tous les éléments de l'anomalie", ""),
        ("selected", "Seulement les éléments sélectionnés dans la vue 3D", "")])
    all_lots: bpy.props.BoolProperty(name="Valable pour tous les lots", default=False,
                                     description="Sinon, la dérogation ne vaut que pour le lot de cette recette")

    @classmethod
    def poll(cls, context):
        return usable_result() is not None

    def invoke(self, context, event):
        if self.index >= 0:
            context.scene.checkers.anomaly_index = self.index
        return context.window_manager.invoke_props_dialog(self, width=460)

    def draw(self, context):
        layout = self.layout
        layout.prop(self, "reason")
        layout.prop(self, "scope")
        layout.prop(self, "all_lots")

    def execute(self, context):
        props = context.scene.checkers
        result = usable_result()
        if result is None:
            self.report({"WARNING"}, "Le modèle a changé depuis la recette : relancez-la.")
            return {"CANCELLED"}
        index = self.index if self.index >= 0 else props.anomaly_index
        if not 0 <= index < len(result.anomalies):
            return {"CANCELLED"}
        anomaly = result.anomalies[index]
        guids = None
        if self.scope == "selected":
            selected = {bonsai_link.guid_of(obj) for obj in context.selected_objects}
            guids = [g for g in anomaly["guids"] if g in selected]
            if not guids:
                self.report({"ERROR"}, "Aucun élément sélectionné ne fait partie de cette anomalie.")
                return {"CANCELLED"}
        path = _derogations_target(props)
        if not path:
            self.report({"ERROR"}, "Choisissez le fichier des dérogations (section Recette) : la maquette n'est pas "
                                   "enregistrée sur disque.")
            return {"CANCELLED"}
        try:
            document = derogations.load_or_new(path)
            entry = derogations.make_entry(anomaly, self.reason, author=result.settings.tester,
                                           lot="" if self.all_lots else result.settings.lot, guids=guids)
            derogations.add(document, entry)
            derogations.save(document, path)
        except (derogations.DerogationError, OSError) as exc:
            self.report({"ERROR"}, f"Dérogation impossible : {exc}")
            return {"CANCELLED"}
        _apply_derogations(result, document, path)
        _show_result(props, result)
        viewer.clear()
        _state["conflict"] = None
        count = len(entry["guids"]) or anomaly["count"]
        self.report({"INFO"}, f"{count} élément(s) ou constat(s) admis. Verdict : {result.verdict}. "
                              f"Dérogation enregistrée dans {os.path.basename(path)}")
        return {"FINISHED"}


class CHECKERS_OT_revoke_derogation(bpy.types.Operator):
    bl_idname = "checkers.revoke_derogation"
    bl_label = "Retirer la dérogation"
    bl_description = "Retire cette dérogation du fichier : l'anomalie redevient ouverte et compte dans le verdict"
    bl_options = {"REGISTER", "UNDO"}

    index: bpy.props.IntProperty(default=-1)

    @classmethod
    def poll(cls, context):
        result = usable_result()
        return result is not None and bool(result.admitted)

    def invoke(self, context, event):
        return context.window_manager.invoke_confirm(self, event)

    def execute(self, context):
        props = context.scene.checkers
        result = usable_result()
        index = self.index if self.index >= 0 else props.admitted_index
        if result is None or not 0 <= index < len(result.admitted):
            return {"CANCELLED"}
        entry_id = result.admitted[index]["derogation"]["id"]
        path = result.settings.derogations_path
        try:
            document = derogations.load(path)
            derogations.revoke(document, entry_id)
            derogations.save(document, path)
        except (derogations.DerogationError, OSError) as exc:
            self.report({"ERROR"}, f"Retrait impossible : {exc}")
            return {"CANCELLED"}
        _apply_derogations(result, document, path)
        _show_result(props, result)
        self.report({"INFO"}, f"Dérogation retirée. Verdict : {result.verdict}")
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
            shading = space.shading
            _state["shading"].setdefault(space.as_pointer(), (shading.type, shading.color_type))
            # La couleur d'objet ne s'affiche qu'en mode solide (ou filaire).
            if shading.type not in ("SOLID", "WIREFRAME"):
                shading.type = "SOLID"
            shading.color_type = "OBJECT"
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
            original = _state["shading"].get(space.as_pointer())
            if original is not None:
                space.shading.type, space.shading.color_type = original
        _state["shading"].clear()
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
        except Exception as exc:
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
        return result is not None and bool(result.anomalies or result.admitted
                                           or (result.evolution and result.evolution["resolved"]))

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
                resolved=result.evolution["resolved"] if result.evolution else (), admitted=result.admitted)
        except Exception as exc:
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
        except Exception as exc:
            self.report({"ERROR"}, f"Enregistrement impossible : {exc}")
            return {"CANCELLED"}
        self.report({"INFO"}, f"Inventaire enregistré : {len(data['equipment'])} équipement(s), "
                              f"{len(data['spaces'])} local(aux)")
        return {"FINISHED"}


class CHECKERS_OT_export_survey(bpy.types.Operator, ExportHelper):
    bl_idname = "checkers.export_survey"
    bl_label = "Fiche de relevé terrain (Excel)"
    bl_description = ("Liste les équipements de la maquette par niveau et par local, avec des colonnes à remplir "
                      "sur site (présent, repère et n° de série relevés). La fiche remplie se contrôle ensuite "
                      "dans « Contrôles > Relevé terrain »")

    filename_ext = ".xlsx"
    filter_glob: bpy.props.StringProperty(default="*.xlsx", options={"HIDDEN"})

    @classmethod
    def poll(cls, context):
        return bonsai_link.current_ifc() is not None

    def invoke(self, context, event):
        self.filepath = _default_name("Releve_terrain", context.scene.checkers.lot.strip(), datetime.now(), ".xlsx")
        return ExportHelper.invoke(self, context, event)

    def execute(self, context):
        ifc_file = bonsai_link.current_ifc()
        if ifc_file is None:
            self.report({"ERROR"}, "Aucun modèle IFC ouvert dans Bonsai.")
            return {"CANCELLED"}
        try:
            from .core import inventory, survey
            equipment = inventory.collect_equipment(ifc_file, LevelCache())
            survey.write_survey(
                bpy.path.ensure_ext(self.filepath, ".xlsx"), equipment,
                model_info(ifc_file, bonsai_link.current_ifc_path()),
                marking=context.scene.checkers.marking_text(), date=datetime.now())
        except ImportError as exc:
            self.report({"ERROR"}, f"Fiche de relevé impossible : {_missing_library(exc)}")
            return {"CANCELLED"}
        except Exception as exc:
            self.report({"ERROR"}, f"Enregistrement impossible : {exc}")
            return {"CANCELLED"}
        self.report({"INFO"}, f"Fiche de relevé enregistrée : {len(equipment)} équipement(s) à vérifier")
        return {"FINISHED"}


class CHECKERS_OT_add_lot(bpy.types.Operator):
    bl_idname = "checkers.add_lot"
    bl_label = "Ajouter un lot"
    bl_description = "Ajoute la maquette IFC d'un autre lot, à comparer à la maquette ouverte"
    bl_options = {"REGISTER", "UNDO"}

    def execute(self, context):
        props = context.scene.checkers
        props.lots.add()
        props.lot_index = len(props.lots) - 1
        return {"FINISHED"}


class CHECKERS_OT_remove_lot(bpy.types.Operator):
    bl_idname = "checkers.remove_lot"
    bl_label = "Retirer le lot"
    bl_description = "Retire ce lot de la liste (le fichier n'est pas supprimé)"
    bl_options = {"REGISTER", "UNDO"}

    @classmethod
    def poll(cls, context):
        return len(context.scene.checkers.lots) > 0

    def execute(self, context):
        props = context.scene.checkers
        if 0 <= props.lot_index < len(props.lots):
            props.lots.remove(props.lot_index)
            props.lot_index = max(0, min(props.lot_index, len(props.lots) - 1))
        return {"FINISHED"}


class CHECKERS_OT_add_clash_rule(bpy.types.Operator):
    bl_idname = "checkers.add_clash_rule"
    bl_label = "Ajouter une règle"
    bl_description = "Ajoute une règle de conflits (par défaut : conflits internes, intersection de plus de 1 cm)"
    bl_options = {"REGISTER", "UNDO"}

    def execute(self, context):
        props = context.scene.checkers
        props.clash_rules.add()
        props.clash_rule_index = len(props.clash_rules) - 1
        props.check_clashes = True
        return {"FINISHED"}


class CHECKERS_OT_remove_clash_rule(bpy.types.Operator):
    bl_idname = "checkers.remove_clash_rule"
    bl_label = "Retirer la règle"
    bl_options = {"REGISTER", "UNDO"}

    @classmethod
    def poll(cls, context):
        return len(context.scene.checkers.clash_rules) > 0

    def execute(self, context):
        props = context.scene.checkers
        if 0 <= props.clash_rule_index < len(props.clash_rules):
            props.clash_rules.remove(props.clash_rule_index)
            props.clash_rule_index = max(0, min(props.clash_rule_index, len(props.clash_rules) - 1))
        return {"FINISHED"}


class CHECKERS_OT_default_clash_rules(bpy.types.Operator):
    bl_idname = "checkers.default_clash_rules"
    bl_label = "Une règle par lot"
    bl_description = ("Ajoute, pour chaque lot associé qui n'en a pas encore, une règle « intersection de plus de "
                      "1 cm, tous éléments » avec la maquette ouverte")
    bl_options = {"REGISTER", "UNDO"}

    @classmethod
    def poll(cls, context):
        return any(lot.path for lot in context.scene.checkers.lots)

    def execute(self, context):
        props = context.scene.checkers
        covered = {rule.other for rule in props.clash_rules}
        added = 0
        for lot in props.lots:
            if not lot.path:
                continue
            name = lot_name(lot.name, bpy.path.abspath(lot.path))
            if name in covered:
                continue
            rule = props.clash_rules.add()
            rule.other = name
            added += 1
        props.check_clashes = props.check_clashes or added > 0
        props.clash_rule_index = max(0, len(props.clash_rules) - 1)
        self.report({"INFO"}, f"{added} règle(s) ajoutée(s)")
        return {"FINISHED"}


class CHECKERS_OT_export_tracking(bpy.types.Operator, ExportHelper):
    bl_idname = "checkers.export_tracking"
    bl_label = "Tableau de suivi (Excel)"
    bl_description = ("Consolide tous les journaux de recette du dossier choisi : état de la dernière recette de "
                      "chaque lot, tendance, et historique complet. Aucune maquette n'a besoin d'être ouverte")

    filename_ext = ".xlsx"
    filter_glob: bpy.props.StringProperty(default="*.xlsx", options={"HIDDEN"})

    @classmethod
    def poll(cls, context):
        return bool(context.scene.checkers.journals_dir)

    def invoke(self, context, event):
        self.filepath = _default_name("Suivi_recettes", "", datetime.now(), ".xlsx")
        return ExportHelper.invoke(self, context, event)

    def execute(self, context):
        props = context.scene.checkers
        try:
            from .core import tracking
            records, rejected = tracking.collect_journals(bpy.path.abspath(props.journals_dir))
            if not records:
                self.report({"WARNING"}, "Aucun journal de recette (.json) dans ce dossier.")
                return {"CANCELLED"}
            tracking.write_tracking(bpy.path.ensure_ext(self.filepath, ".xlsx"), records, rejected,
                                    marking=props.marking_text(), date=datetime.now())
        except ImportError as exc:
            self.report({"ERROR"}, f"Tableau de suivi impossible : {_missing_library(exc)}")
            return {"CANCELLED"}
        except Exception as exc:
            self.report({"ERROR"}, f"Tableau de suivi impossible : {exc}")
            return {"CANCELLED"}
        message = f"Tableau de suivi enregistré : {len(records)} recette(s)"
        if rejected:
            message += f", {len(rejected)} journal(aux) refusé(s)"
        self.report({"INFO"}, message)
        return {"FINISHED"}


class CHECKERS_OT_diagnostic(bpy.types.Operator):
    bl_idname = "checkers.diagnostic"
    bl_label = "Vérifier l'installation"
    bl_description = ("Vérifie que Bonsai et ses bibliothèques (IfcOpenShell, ifctester, openpyxl) sont présents "
                      "et que les IDS d'exemple sont valides sur ce poste")

    def execute(self, context):
        lines = diagnostic.run(bonsai_link.current_ifc(), bonsai_found=bonsai_link.bonsai_tool() is not None)
        context.scene.checkers.diagnostic = diagnostic.to_text(lines)
        failed = [line for line in lines if not line["ok"]]
        # Sans modèle ouvert, le reste peut être en ordre : ce n'est pas un défaut d'installation.
        blocking = [line for line in failed if line["label"] != "Modèle IFC"]
        if blocking:
            self.report({"ERROR"}, f"Installation incomplète : {len(blocking)} point(s) à corriger (voir le panneau)")
        else:
            self.report({"INFO"}, "Installation en ordre")
        return {"FINISHED"}


classes = (
    CHECKERS_OT_run_recette,
    CHECKERS_OT_select_anomaly,
    CHECKERS_OT_step_conflict,
    CHECKERS_OT_clear_conflict_view,
    CHECKERS_OT_admit_anomaly,
    CHECKERS_OT_revoke_derogation,
    CHECKERS_OT_color_results,
    CHECKERS_OT_reset_colors,
    CHECKERS_OT_use_example_ids,
    CHECKERS_OT_use_example_rules,
    CHECKERS_OT_export_report,
    CHECKERS_OT_export_bcf,
    CHECKERS_OT_export_inventory,
    CHECKERS_OT_export_survey,
    CHECKERS_OT_export_tracking,
    CHECKERS_OT_diagnostic,
    CHECKERS_OT_add_lot,
    CHECKERS_OT_remove_lot,
    CHECKERS_OT_add_clash_rule,
    CHECKERS_OT_remove_clash_rule,
    CHECKERS_OT_default_clash_rules,
)
