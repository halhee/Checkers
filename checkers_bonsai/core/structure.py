"""Contrôles de structure du modèle, indépendants de tout référentiel.

Ils vérifient que la maquette est exploitable avant même de regarder ses
propriétés : un projet unique, des unités, une arborescence spatiale, des
éléments rattachés à un niveau, des identifiants uniques. Chaque contrôle a
une gravité, qui fixe son effet sur le verdict :

- Bloquant : la maquette ne peut pas être exploitée telle quelle (rejet) ;
- Majeur : exploitable avec des réserves ;
- Mineur : signalé, sans effet sur le verdict.

Les gravités et le seuil ci-dessous sont des choix par défaut, à adapter au
cahier des charges du projet.
"""
from collections import Counter, defaultdict

import ifcopenshell.util.element
import ifcopenshell.util.unit

BLOCKING, MAJOR, MINOR = "Bloquant", "Majeur", "Mineur"
MODEL_LEVEL = "Modèle"
# Part d'objets génériques (IfcBuildingElementProxy) au-delà de laquelle le
# constat passe de mineur à majeur.
PROXY_SHARE_MAJOR = 0.10


def _by_type(ifc_file, ifc_class) -> list:
    """by_type tolérant : une classe absente du schéma (IfcFacility en IFC4,
    IfcMapConversion en IFC2X3) donne une liste vide."""
    try:
        return list(ifc_file.by_type(ifc_class))
    except Exception:
        return []


def _is_a(entity, ifc_class) -> bool:
    try:
        return bool(entity.is_a(ifc_class))
    except Exception:
        return False


def _blank(value) -> bool:
    return value is None or not str(value).strip()


def _physical_elements(ifc_file) -> list:
    """Éléments physiques, hors réservations et ouvertures (IfcFeatureElement),
    qui ne sont pas rangés dans la structure spatiale."""
    return [e for e in _by_type(ifc_file, "IfcElement") if not _is_a(e, "IfcFeatureElement")]


def _check_single_project(ifc_file, elements):
    projects = _by_type(ifc_file, "IfcProject")
    if len(projects) == 1:
        return None
    return projects, f"{len(projects)} IfcProject dans le modèle (un seul attendu)"


def _check_length_unit(ifc_file, elements):
    if ifcopenshell.util.unit.get_project_unit(ifc_file, "LENGTHUNIT") is not None:
        return None
    return [], "Aucune unité de longueur déclarée dans le projet"


def _check_site_or_facility(ifc_file, elements):
    if any(_by_type(ifc_file, c) for c in ("IfcSite", "IfcBuilding", "IfcFacility")):
        return None
    return [], "Ni site, ni bâtiment, ni ouvrage dans le modèle"


def _check_buildings_have_storeys(ifc_file, elements):
    with_storeys = set()
    for storey in _by_type(ifc_file, "IfcBuildingStorey"):
        for rel in getattr(storey, "Decomposes", None) or ():
            with_storeys.add(rel.RelatingObject)
    missing = [b for b in _by_type(ifc_file, "IfcBuilding") if b not in with_storeys]
    return (missing, f"{len(missing)} bâtiment(s) sans aucun étage") if missing else None


def _check_elements_contained(ifc_file, elements):
    orphans = [e for e in elements if ifcopenshell.util.element.get_container(e) is None]
    return (orphans, f"{len(orphans)} élément(s) hors de la structure spatiale") if orphans else None


def _check_unique_guids(ifc_file, elements):
    roots = _by_type(ifc_file, "IfcRoot")
    counts = Counter(getattr(e, "GlobalId", None) for e in roots)
    duplicated = [e for e in roots if counts[getattr(e, "GlobalId", None)] > 1]
    if not duplicated:
        return None
    return duplicated, f"{len(duplicated)} objets partagent un GlobalId avec un autre"


def _check_elements_named(ifc_file, elements):
    unnamed = [e for e in elements if _blank(getattr(e, "Name", None))]
    return (unnamed, f"{len(unnamed)} élément(s) sans nom") if unnamed else None


def _check_proxies(ifc_file, elements):
    proxies = _by_type(ifc_file, "IfcBuildingElementProxy")
    if not proxies:
        return None
    share = len(proxies) / len(elements) if elements else 1.0
    severity = MAJOR if share > PROXY_SHARE_MAJOR else MINOR
    return proxies, f"{len(proxies)} objet(s) générique(s), soit {share:.0%} des éléments", severity


def _check_storeys_named(ifc_file, elements):
    storeys = _by_type(ifc_file, "IfcBuildingStorey")
    names = Counter(str(s.Name).strip() for s in storeys if not _blank(s.Name))
    faulty = [s for s in storeys if _blank(s.Name) or names[str(s.Name).strip()] > 1]
    return (faulty, f"{len(faulty)} étage(s) sans nom ou de même nom") if faulty else None


def _check_georeferencing(ifc_file, elements):
    if _by_type(ifc_file, "IfcMapConversion"):
        return None
    if any(getattr(s, "RefLatitude", None) and getattr(s, "RefLongitude", None) for s in _by_type(ifc_file, "IfcSite")):
        return None
    return [], "Ni conversion de coordonnées (IfcMapConversion), ni latitude et longitude du site"


def _check_elements_have_geometry(ifc_file, elements):
    # Un assemblage (escalier, mur-rideau...) porte sa géométrie dans ses parties.
    bare = [e for e in elements
            if getattr(e, "Representation", None) is None and not getattr(e, "IsDecomposedBy", None)]
    return (bare, f"{len(bare)} élément(s) sans géométrie") if bare else None


# (code, intitulé, gravité, contrôle, constat rattaché aux éléments ?)
CHECKS = (
    ("S01", "Projet unique", BLOCKING, _check_single_project, False),
    ("S02", "Identifiants (GlobalId) uniques", BLOCKING, _check_unique_guids, False),
    ("S03", "Unité de longueur déclarée", MAJOR, _check_length_unit, False),
    ("S04", "Site, bâtiment ou ouvrage présent", MAJOR, _check_site_or_facility, False),
    ("S05", "Bâtiments découpés en étages", MAJOR, _check_buildings_have_storeys, False),
    ("S06", "Éléments rattachés à la structure spatiale", MAJOR, _check_elements_contained, True),
    ("S07", "Objets génériques (IfcBuildingElementProxy)", MINOR, _check_proxies, True),
    ("S08", "Étages nommés, sans doublon", MINOR, _check_storeys_named, False),
    ("S09", "Éléments nommés", MINOR, _check_elements_named, True),
    ("S10", "Éléments avec géométrie", MINOR, _check_elements_have_geometry, True),
    ("S11", "Géoréférencement renseigné", MINOR, _check_georeferencing, False),
)


def check_structure(ifc_file, level_of) -> dict:
    """Retourne :
    - checks : un dict par contrôle (code, title, severity, status parmi
      « ok », « ko », « erreur », message, count) ;
    - anomalies : {(niveau, classe IFC, intitulé): {"severity", "guids", "count"}} ;
    - blocking / major / minor : nombre de contrôles en échec par gravité."""
    elements = _physical_elements(ifc_file)
    checks = []
    anomalies = {}
    failed = Counter()

    for code, title, severity, check, per_element in CHECKS:
        try:
            outcome = check(ifc_file, elements)
        except Exception as exc:
            # Un contrôle qui échoue ne doit pas faire perdre les autres.
            checks.append({"code": code, "title": title, "severity": severity, "status": "erreur",
                           "message": f"Contrôle impossible ({type(exc).__name__})", "count": 0})
            continue
        if outcome is None:
            checks.append({"code": code, "title": title, "severity": severity, "status": "ok",
                           "message": "", "count": 0})
            continue
        entities, message = outcome[0], outcome[1]
        if len(outcome) > 2:
            severity = outcome[2]
        count = max(len(entities), 1)
        checks.append({"code": code, "title": title, "severity": severity, "status": "ko",
                       "message": message, "count": count})
        failed[severity] += 1

        if per_element:
            groups = defaultdict(set)
            for entity in entities:
                guid = getattr(entity, "GlobalId", None)
                if guid:
                    groups[(level_of(entity), entity.is_a(), title)].add(guid)
            for key, guids in groups.items():
                anomalies[key] = {"severity": severity, "guids": guids, "count": len(guids)}
        else:
            guids = {g for g in (getattr(e, "GlobalId", None) for e in entities) if g}
            ifc_class = entities[0].is_a() if entities else "IfcProject"
            anomalies[(MODEL_LEVEL, ifc_class, title)] = {"severity": severity, "guids": guids, "count": count}

    return {"checks": checks, "anomalies": anomalies,
            "blocking": failed[BLOCKING], "major": failed[MAJOR], "minor": failed[MINOR]}
