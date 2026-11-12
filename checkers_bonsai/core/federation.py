"""Cohérence de référence entre lots : la maquette ouverte et les maquettes
des autres lots partagent-elles le même cadre ?

Contrôles sans géométrie, donc rapides : même schéma, même projet, même
unité, même point d'origine, mêmes étages (noms et altitudes), même
géoréférencement, même nord, aucun élément partagé. Un écart ici explique
la plupart des maquettes qui « ne se superposent pas » ; il fausse aussi la
détection de conflits géométriques (voir clashes.py), qui compare les
coordonnées des maquettes telles quelles.

Les gravités suivent celles des contrôles de structure (voir structure.py) :
choix par défaut, à adapter au cahier des charges du projet.
"""
import math
import os
from collections import Counter, defaultdict

import ifcopenshell
import ifcopenshell.util.element
import ifcopenshell.util.placement
import ifcopenshell.util.unit

from .spatial import MODEL_LEVEL
from .structure import BLOCKING, MAJOR, MINOR
from .trace import file_info

# Écart de position (mètres) et d'angle (degrés) en deçà duquel deux lots
# sont considérés alignés. 0,01° représente 1,7 cm à 100 m.
POSITION_TOLERANCE = 0.01
ANGLE_TOLERANCE = 0.01
MAX_LOTS = 20

# Libellé réservé à la maquette ouverte dans les règles de conflits.
CURRENT = "Ce modèle"


class FederationError(ValueError):
    """Liste des lots invalide ou maquette d'un lot illisible."""


def _by_type(ifc_file, ifc_class) -> list:
    try:
        return list(ifc_file.by_type(ifc_class))
    except Exception:
        return []


def _is_a(entity, ifc_class) -> bool:
    try:
        return bool(entity.is_a(ifc_class))
    except Exception:
        return False


def lot_name(name: str, path: str) -> str:
    return (name or "").strip() or os.path.splitext(os.path.basename(path))[0]


def open_lots(lots) -> list:
    """`lots` : liste de (nom, chemin). Retourne une liste de dicts name,
    path, ifc, info (empreinte du fichier). Les noms doivent être uniques :
    ils identifient les anomalies d'une recette à l'autre."""
    opened, names = [], set()
    if len(lots) > MAX_LOTS:
        raise FederationError(f"Trop de lots associés (maximum {MAX_LOTS})")
    for name, path in lots:
        name = lot_name(name, path)
        if name == CURRENT:
            raise FederationError(f"« {CURRENT} » est réservé à la maquette ouverte : renommez ce lot")
        if name in names:
            raise FederationError(f"Deux lots portent le nom « {name} » : chaque lot doit avoir un nom unique")
        names.add(name)
        if not path or not os.path.isfile(path):
            raise FederationError(f"Maquette du lot « {name} » introuvable : {path}")
        try:
            ifc_file = ifcopenshell.open(path)
        except Exception as exc:
            raise FederationError(f"Maquette du lot « {name} » illisible ({type(exc).__name__})") from exc
        opened.append({"name": name, "path": path, "ifc": ifc_file, "info": file_info(path)})
    return opened


# --------------------------------------------------------------------- lecture

def _unit_scale(ifc_file) -> float:
    try:
        return float(ifcopenshell.util.unit.calculate_unit_scale(ifc_file)) or 1.0
    except Exception:
        return 1.0


def _placement(entity, scale):
    """(position en m, matrice) du placement absolu de l'entité, ou None.
    IfcOpenShell 0.9 : util.placement.get_placement (en mètres avec
    should_return_si) ; versions antérieures : get_local_placement, en unités
    du projet."""
    placement = getattr(entity, "ObjectPlacement", None)
    if placement is None:
        return None
    try:
        if hasattr(ifcopenshell.util.placement, "get_placement"):
            matrix = ifcopenshell.util.placement.get_placement(placement, should_return_si=True)
            factor = 1.0
        else:
            matrix = ifcopenshell.util.placement.get_local_placement(placement)
            factor = scale
    except Exception:
        return None
    return tuple(float(matrix[i][3]) * factor for i in range(3)), matrix


def _origin(ifc_file, scale):
    """Position (m) et orientation (°) du premier bâtiment ou ouvrage, sinon
    du site : c'est le repère dans lequel tous les éléments sont placés."""
    for ifc_class in ("IfcBuilding", "IfcFacility", "IfcSite"):
        for entity in _by_type(ifc_file, ifc_class):
            placed = _placement(entity, scale)
            if placed is not None:
                position, matrix = placed
                angle = math.degrees(math.atan2(float(matrix[1][0]), float(matrix[0][0])))
                return entity, position, angle
    return None


def _storeys(ifc_file, scale) -> dict:
    """{nom: (étage, altitude en m)}. Altitude absolue tirée du placement,
    à défaut de l'attribut Elevation."""
    result = {}
    for storey in _by_type(ifc_file, "IfcBuildingStorey"):
        name = str(storey.Name or "").strip()
        if not name or name in result:
            continue
        placed = _placement(storey, scale)
        if placed is not None:
            altitude = placed[0][2]
        elif getattr(storey, "Elevation", None) is not None:
            altitude = float(storey.Elevation) * scale
        else:
            continue
        result[name] = (storey, altitude)
    return result


def _number(value, default: float) -> float:
    try:
        return float(value) if value is not None else default
    except (TypeError, ValueError):
        return default


def _pset(entity, name: str) -> dict:
    try:
        return ifcopenshell.util.element.get_pset(entity, name) or {}
    except Exception:
        return {}


def _map_conversion(ifc_file):
    """Géoréférencement : IfcMapConversion (IFC4 et après), ou, en IFC2X3,
    les jeux de propriétés ePSet_MapConversion et ePSet_ProjectedCRS de
    l'IfcProject (guide buildingSMART du géoréférencement)."""
    conversions = _by_type(ifc_file, "IfcMapConversion")
    if conversions:
        c = conversions[0]
        values = {key: getattr(c, key, None) for key in ("Eastings", "Northings", "OrthogonalHeight",
                                                          "XAxisAbscissa", "XAxisOrdinate", "Scale")}
        entity, crs = c, getattr(getattr(c, "TargetCRS", None), "Name", None) or ""
    else:
        project = _project(ifc_file)
        values = _pset(project, "ePSet_MapConversion") if project is not None else {}
        if not any(values.get(key) is not None for key in ("Eastings", "Northings", "OrthogonalHeight")):
            return None
        entity, crs = project, _pset(project, "ePSet_ProjectedCRS").get("Name") or ""
    abscissa, ordinate = _number(values.get("XAxisAbscissa"), 1.0), _number(values.get("XAxisOrdinate"), 0.0)
    if abscissa == 0.0 and ordinate == 0.0:
        abscissa = 1.0
    return {"entity": entity, "eastings": _number(values.get("Eastings"), 0.0),
            "northings": _number(values.get("Northings"), 0.0),
            "height": _number(values.get("OrthogonalHeight"), 0.0),
            "rotation": math.degrees(math.atan2(ordinate, abscissa)),
            "scale": _number(values.get("Scale"), 1.0), "crs": str(crs)}


def _true_north(ifc_file):
    """Angle (°) du nord géographique dans le contexte « Model » ; l'axe +Y
    par défaut."""
    for context in _by_type(ifc_file, "IfcGeometricRepresentationContext"):
        if _is_a(context, "IfcGeometricRepresentationSubContext"):
            continue
        if (context.ContextType or "") != "Model":
            continue
        north = getattr(context, "TrueNorth", None)
        ratios = getattr(north, "DirectionRatios", None) if north is not None else None
        if ratios and len(ratios) >= 2:
            return context, math.degrees(math.atan2(float(ratios[0]), float(ratios[1])))
        return context, 0.0
    return None, 0.0


def _angle_gap(a: float, b: float) -> float:
    return abs((a - b + 180.0) % 360.0 - 180.0)


def _project(ifc_file):
    projects = _by_type(ifc_file, "IfcProject")
    return projects[0] if projects else None


def _unit_label(scale: float) -> str:
    known = {1.0: "mètre", 0.001: "millimètre", 0.01: "centimètre", 0.1: "décimètre", 0.3048: "pied",
             0.0254: "pouce"}
    for value, label in known.items():
        if abs(scale - value) < 1e-9:
            return label
    return f"{scale:g} m"


# -------------------------------------------------------------------- contrôles
# Chaque contrôle reçoit (maquette ouverte, autre lot, lectures en cache) et
# retourne None (cohérent) ou (entités de la maquette ouverte, message[, gravité]).

def _check_schema(current, other, cache):
    a, b = current.schema, other.schema
    return None if a == b else ([], f"schéma {a} contre {b}")


def _check_project(current, other, cache):
    a, b = _project(current), _project(other)
    if a is None or b is None:
        return None  # déjà signalé par le contrôle de structure S01
    if a.GlobalId == b.GlobalId:
        return None
    if str(a.Name or "").strip() and str(a.Name).strip() == str(b.Name or "").strip():
        return [a], f"même nom de projet (« {a.Name} ») mais identifiants différents", MINOR
    return [a], f"projet « {a.Name or '—'} » contre « {b.Name or '—'} »"


def _check_unit(current, other, cache):
    a, b = cache["scale"][id(current)], cache["scale"][id(other)]
    if abs(a - b) < 1e-9:
        return None
    return [], f"{_unit_label(a)} contre {_unit_label(b)}"


def _check_origin(current, other, cache):
    a, b = cache["origin"][id(current)], cache["origin"][id(other)]
    if a is None or b is None:
        return None
    gap = math.dist(a[1], b[1])
    turn = _angle_gap(a[2], b[2])
    if gap <= POSITION_TOLERANCE and turn <= ANGLE_TOLERANCE:
        return None
    return [a[0]], f"origine décalée de {gap:.3f} m et tournée de {turn:.3f}°"


def _check_storey_altitudes(current, other, cache):
    mine, theirs = cache["storeys"][id(current)], cache["storeys"][id(other)]
    faulty, details = [], []
    for name, (storey, altitude) in mine.items():
        if name in theirs and abs(theirs[name][1] - altitude) > POSITION_TOLERANCE:
            faulty.append(storey)
            details.append(f"{name} : {altitude:+.3f} m contre {theirs[name][1]:+.3f} m")
    if not faulty:
        return None
    return faulty, f"{len(faulty)} étage(s) de même nom à des altitudes différentes ({'; '.join(details[:3])}" \
                   f"{' ; …' if len(details) > 3 else ''})"


def _check_storey_names(current, other, cache):
    mine, theirs = cache["storeys"][id(current)], cache["storeys"][id(other)]
    faulty, details = [], []
    for name, (storey, altitude) in mine.items():
        if name in theirs:
            continue
        twins = [n for n, (_, z) in theirs.items() if abs(z - altitude) <= POSITION_TOLERANCE and n not in mine]
        if twins:
            faulty.append(storey)
            details.append(f"« {name} » / « {twins[0]} »")
    if not faulty:
        return None
    return faulty, f"{len(faulty)} étage(s) à la même altitude mais nommé(s) autrement ({', '.join(details[:3])}" \
                   f"{', …' if len(details) > 3 else ''})"


def _check_georeferencing(current, other, cache):
    a, b = cache["map"][id(current)], cache["map"][id(other)]
    if a is None and b is None:
        return None
    if a is None or b is None:
        return [], "géoréférencement présent dans un seul des deux lots"
    gaps = []
    planar = math.dist((a["eastings"], a["northings"], a["height"]), (b["eastings"], b["northings"], b["height"]))
    if planar > POSITION_TOLERANCE:
        gaps.append(f"point de référence décalé de {planar:.3f}")
    if _angle_gap(a["rotation"], b["rotation"]) > ANGLE_TOLERANCE:
        gaps.append(f"rotation {a['rotation']:.4f}° contre {b['rotation']:.4f}°")
    if abs(a["scale"] - b["scale"]) > 1e-9:
        gaps.append(f"échelle {a['scale']:g} contre {b['scale']:g}")
    if a["crs"] != b["crs"]:
        gaps.append(f"système « {a['crs'] or '—'} » contre « {b['crs'] or '—'} »")
    return ([a["entity"]], ", ".join(gaps)) if gaps else None


def _check_true_north(current, other, cache):
    a, b = cache["north"][id(current)], cache["north"][id(other)]
    turn = _angle_gap(a[1], b[1])
    if turn <= ANGLE_TOLERANCE:
        return None
    return [a[0]] if a[0] is not None else [], f"nord à {a[1]:.3f}° contre {b[1]:.3f}°"


def _check_shared_elements(current, other, cache):
    """Un même GlobalId dans deux lots désigne deux objets différents dans
    tous les échanges (BCF, suivi, GMAO). Les entités spatiales (site,
    bâtiment, étages) ne sont pas concernées : les partager entre lots est
    une pratique courante pour fédérer les maquettes."""
    theirs = cache["element_guids"][id(other)]
    shared = [e for e in cache["elements"][id(current)] if e.GlobalId in theirs]
    return (shared, f"{len(shared)} élément(s) portent un GlobalId déjà utilisé dans ce lot") if shared else None


# (code, intitulé, gravité, contrôle, constat rattaché aux éléments ?)
CHECKS = (
    ("C01", "Même schéma IFC", MAJOR, _check_schema, False),
    ("C02", "Même projet", MAJOR, _check_project, False),
    ("C03", "Même unité de longueur", MINOR, _check_unit, False),
    ("C04", "Même point d'origine", MAJOR, _check_origin, False),
    ("C05", "Étages : mêmes altitudes", MAJOR, _check_storey_altitudes, True),
    ("C06", "Étages : mêmes noms", MINOR, _check_storey_names, True),
    ("C07", "Même géoréférencement", MAJOR, _check_georeferencing, False),
    ("C08", "Même nord géographique", MAJOR, _check_true_north, False),
    ("C09", "Aucun élément partagé", BLOCKING, _check_shared_elements, True),
)
# Contrôles dont l'échec rend les conflits géométriques entre lots non fiables.
FRAME_CHECKS = {"C03", "C04", "C05"}


def requirement(lot: str, title: str) -> str:
    return f"Cohérence avec « {lot} » : {title}"


def _read(ifc_file, cache) -> None:
    key = id(ifc_file)
    if key in cache["scale"]:
        return
    scale = _unit_scale(ifc_file)
    cache["scale"][key] = scale
    cache["origin"][key] = _origin(ifc_file, scale)
    cache["storeys"][key] = _storeys(ifc_file, scale)
    cache["map"][key] = _map_conversion(ifc_file)
    cache["north"][key] = _true_north(ifc_file)
    elements = [e for e in _by_type(ifc_file, "IfcElement") if getattr(e, "GlobalId", None)]
    cache["elements"][key] = elements
    cache["element_guids"][key] = {e.GlobalId for e in elements}


def check_federation(ifc_file, lots: list, level_of) -> dict:
    """`lots` : résultat de open_lots. Retourne :
    - lots : nom, fichier (empreinte), schéma et projet de chaque lot ;
    - checks : un dict par contrôle et par lot (code, title, lot, severity,
      status parmi « ok », « ko », « erreur », message, count) ;
    - anomalies : {(niveau, classe IFC, intitulé): {"severity", "guids", "count"}} ;
    - blocking / major / minor / errors : nombre de contrôles en échec ;
    - frame_ok : False si un lot n'est pas dans le même repère (unité,
      origine, altitudes) : ses conflits géométriques ne sont pas fiables."""
    cache = defaultdict(dict)
    _read(ifc_file, cache)
    checks, anomalies, failed = [], {}, Counter()
    frame_ok = True
    summary = []

    for lot in lots:
        other = lot["ifc"]
        _read(other, cache)
        project = _project(other)
        summary.append({"name": lot["name"], "file": lot["info"], "schema": other.schema,
                        "project": str(project.Name or "") if project is not None else ""})
        for code, title, severity, check, per_element in CHECKS:
            base = {"code": code, "title": title, "lot": lot["name"], "severity": severity}
            try:
                outcome = check(ifc_file, other, cache)
            except Exception as exc:
                checks.append({**base, "status": "erreur", "message": f"Contrôle impossible ({type(exc).__name__})",
                               "count": 0})
                failed["erreur"] += 1
                if code in FRAME_CHECKS:
                    frame_ok = False
                continue
            if outcome is None:
                checks.append({**base, "status": "ok", "message": "", "count": 0})
                continue
            entities, message = outcome[0], outcome[1]
            severity = outcome[2] if len(outcome) > 2 else severity
            count = max(len(entities), 1)
            checks.append({**base, "severity": severity, "status": "ko", "message": message, "count": count})
            failed[severity] += 1
            if code in FRAME_CHECKS:
                frame_ok = False
            label = requirement(lot["name"], title)
            if per_element:
                groups = defaultdict(set)
                for entity in entities:
                    level = str(entity.Name or "").strip() if _is_a(entity, "IfcBuildingStorey") else level_of(entity)
                    groups[(level or MODEL_LEVEL, entity.is_a(), label)].add(entity.GlobalId)
                for key, guids in groups.items():
                    anomalies[key] = {"severity": severity, "guids": guids, "count": len(guids)}
            else:
                guids = {g for g in (getattr(e, "GlobalId", None) for e in entities) if g}
                ifc_class = entities[0].is_a() if entities else "IfcProject"
                anomalies[(MODEL_LEVEL, ifc_class, label)] = {"severity": severity, "guids": guids, "count": count,
                                                              "message": message}

    return {"lots": summary, "checks": checks, "anomalies": anomalies, "blocking": failed[BLOCKING],
            "major": failed[MAJOR], "minor": failed[MINOR], "errors": failed["erreur"], "frame_ok": frame_ok}
