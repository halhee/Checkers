"""Outils communs aux tests : chemins, modèles IFC de test et doublures de
spécifications IDS.

Les tests s'exécutent sans Blender ni IfcOpenShell, grâce aux doublures de
tests/fakes/ (mises en tête de sys.path ici).
"""
import os
import sys
import tempfile
import types
from unittest import mock

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FAKES = os.path.join(ROOT, "tests", "fakes")
PACKAGE = os.path.join(ROOT, "checkers_bonsai")
for path in (ROOT, FAKES):
    if path not in sys.path:
        sys.path.insert(0, path)

import ifcopenshell  # noqa: E402  (doublure)

from checkers_bonsai.core import ids as ids_check, recette  # noqa: E402

EXAMPLES = os.path.join(PACKAGE, "examples")
RULES_XLSX = os.path.join(EXAMPLES, "regles_exemple.xlsx")
# Dossier Schemas/ de https://github.com/buildingSMART/BCF-XML (branche
# release_2_1) : s'il est fourni et que lxml est installé, les BCF produits
# sont validés contre les schémas officiels.
BCF_XSD_DIR = os.environ.get("BCF_XSD_DIR")


def guid(n: int) -> str:
    return f"0{n:021d}"  # 22 caractères valides pour un IfcGuid


def build_model(schema="IFC4"):
    """Bâtiment sain côté structure : site géoréférencé, 2 étages, un local.
    Quelques murs, une porte et une dalle, dont certains sans leurs PSet."""
    f = ifcopenshell.file(schema)
    f.add("IfcProject", Name="Projet de démonstration – bâtiment A", GlobalId=guid(900))
    site = f.add("IfcSite", Name="Site", GlobalId=guid(905), RefLatitude=(48, 51, 0), RefLongitude=(2, 21, 0))
    building = f.add("IfcBuilding", Name="Bâtiment A", GlobalId=guid(901))
    ifcopenshell.aggregate(building, site)
    rdc = f.add("IfcBuildingStorey", Name="RDC", GlobalId=guid(902))
    r1 = f.add("IfcBuildingStorey", Name="R+1", GlobalId=guid(903))
    ifcopenshell.aggregate(rdc, building)
    ifcopenshell.aggregate(r1, building)
    local = f.add("IfcSpace", Name="012", LongName="Local technique", GlobalId=guid(904),
                  psets={"Qto_SpaceBaseQuantities": {"NetFloorArea": 18.456}})
    ifcopenshell.aggregate(local, rdc)

    ok = {"Pset_WallCommon": {"AcousticRating": "Rw 45", "LoadBearing": True}}
    walls = [
        f.add("IfcWall", GlobalId=guid(1), Name="Mur conforme", psets=ok),
        f.add("IfcWallStandardCase", GlobalId=guid(2), Name="=HYPERLINK(\"x\")",
              psets={"Pset_WallCommon": {"AcousticRating": "  "}}),
        f.add("IfcWall", GlobalId=guid(3), Name="Mur sans PSet"),
    ]
    for wall in walls[:2]:
        ifcopenshell.contain(wall, r1)
    ifcopenshell.contain(walls[2], local)            # contenu dans un local -> niveau RDC
    door = f.add("IfcDoor", GlobalId=guid(4), Name="Porte", psets={"BaseQuantities": {"Height": "deux mètres"}})
    ifcopenshell.contain(door, rdc)
    slab = f.add("IfcSlab", GlobalId=guid(5), Name="Dalle", psets={"BaseQuantities": {"GrossArea": 12.0}})
    ifcopenshell.contain(slab, building)            # directement dans le bâtiment -> sans niveau
    return f


class _Entity(types.SimpleNamespace):
    """Objet lié à un élément (type, système, référence de classification) :
    utilisable dans un ensemble, comme une entité IfcOpenShell."""
    __hash__ = object.__hash__


def add_equipment(f):
    """Ajoute au modèle de build_model des équipements techniques."""
    rdc = f.by_guid(guid(902))
    local = f.by_guid(guid(904))
    pump_type = _Entity(Name="Pompe double DN50")
    system = _Entity(Name="Chauffage")
    reference = _Entity(Identification="52.10", Name="Pompes")
    pump = f.add("IfcPump", GlobalId=guid(20), Name="Pompe de circulation", Tag="PMP-001", PredefinedType="CIRCULATOR",
                 _type=pump_type, _systems=[system], _classifications=[reference], psets={
                     "Pset_ManufacturerTypeInformation": {"Manufacturer": "=Fabricant", "ModelReference": "X200"},
                     "Pset_ManufacturerOccurrence": {"SerialNumber": "SN-42"},
                     "Pset_Warranty": {"WarrantyEndDate": "2028-06-30"}})
    ifcopenshell.contain(pump, local)
    bare = f.add("IfcPump", GlobalId=guid(21), Name="Pompe de relevage")     # rien de renseigné
    ifcopenshell.contain(bare, rdc)
    lift = f.add("IfcTransportElement", GlobalId=guid(22), Name="Ascenseur", Tag="ASC-01")
    ifcopenshell.contain(lift, rdc)
    pipe = f.add("IfcPipeSegment", GlobalId=guid(23), Name="Tuyau")           # réseau linéaire : hors inventaire
    ifcopenshell.contain(pipe, rdc)
    return f


def build_broken_model():
    """Modèle qui échoue à chaque contrôle de structure."""
    f = ifcopenshell.file("IFC4")
    f.length_unit = None
    f.add("IfcProject", Name="Projet 1", GlobalId=guid(900))
    f.add("IfcProject", Name="Projet 2", GlobalId=guid(899))
    empty = f.add("IfcBuilding", Name="Bâtiment sans étage", GlobalId=guid(901))
    building = f.add("IfcBuilding", Name="Bâtiment", GlobalId=guid(906))
    storeys = [f.add("IfcBuildingStorey", Name="Niveau", GlobalId=guid(902)),
               f.add("IfcBuildingStorey", Name="Niveau", GlobalId=guid(903)),
               f.add("IfcBuildingStorey", GlobalId=guid(904))]
    for storey in storeys:
        ifcopenshell.aggregate(storey, building)

    orphan = f.add("IfcWall", GlobalId=guid(1), Name="Mur hors structure")
    unnamed = f.add("IfcWall", GlobalId=guid(2))
    ifcopenshell.contain(unnamed, storeys[0])
    proxy = f.add("IfcBuildingElementProxy", GlobalId=guid(3), Name="Objet générique")
    ifcopenshell.contain(proxy, storeys[0])
    bare = f.add("IfcBeam", GlobalId=guid(4), Name="Poutre sans géométrie", Representation=None)
    ifcopenshell.contain(bare, storeys[0])
    twins = [f.add("IfcColumn", GlobalId=guid(5), Name="Poteau A"), f.add("IfcColumn", GlobalId=guid(5), Name="Poteau B")]
    for twin in twins:
        ifcopenshell.contain(twin, storeys[0])

    # Cas qui ne doivent PAS être signalés : une réservation, la partie d'un
    # assemblage rangé dans un étage, et l'assemblage lui-même sans géométrie.
    f.add("IfcOpeningElement", GlobalId=guid(6))
    stair = f.add("IfcStair", GlobalId=guid(7), Name="Escalier", Representation=None)
    ifcopenshell.contain(stair, storeys[0])
    flight = f.add("IfcBuildingElement", GlobalId=guid(8), Name="Volée")
    ifcopenshell.aggregate(flight, stair)
    return f, {"empty": empty, "orphan": orphan, "unnamed": unnamed, "proxy": proxy, "bare": bare, "twins": twins}


def fake_requirement(description, failures):
    return types.SimpleNamespace(failures=[{"element": e, "reason": "Valeur absente"} for e in failures],
                                 status=not failures, to_string=lambda *a: description)


def fake_spec(name, applicable, failed, cardinality="required", status=None):
    """Spécification IDS telle qu'ifctester la présente après validation."""
    spec = types.SimpleNamespace(
        name=name, identifier=name[:3].upper(), description="", instructions="",
        minOccurs=0 if cardinality == "optional" else 1, maxOccurs=0 if cardinality == "prohibited" else "unbounded",
        applicable_entities=applicable, failed_entities=set(failed),
        requirements=[fake_requirement(f"{name} : exigence", failed)])
    spec.status = bool(not failed and (applicable or cardinality != "required")) if status is None else status
    return spec


def fake_document(*specs, title="Exigences DOE"):
    return types.SimpleNamespace(info={"title": title}, validate=lambda *a, **k: None, specifications=list(specs))


def write_model_file(content="ISO-10303-21;") -> str:
    path = os.path.join(tempfile.mkdtemp(), "maquette.ifc")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(content)
    return path


def run_recette(model=None, failed_walls=slice(2, None), model_content="ISO-10303-21;", **settings):
    """Recette complète (structure, IDS simulé, règles Excel ARC) sur le
    modèle de build_model. `failed_walls` : murs en échec IDS."""
    model = model or build_model()
    walls = model.by_type("IfcWall")
    document = fake_document(fake_spec("Murs – repère", walls, walls[failed_walls]))
    values = dict(ids_path=os.path.join(EXAMPLES, "conception_minimale.ids"), rules_path=RULES_XLSX,
                  corps_etat="ARC", lot="Lot 02 – Gros œuvre", phase="DOE – Dossier des ouvrages exécutés",
                  tester="J. Martin", marking="CONFIDENTIEL")
    values.update(settings)
    with mock.patch.object(ids_check, "load_ids", return_value=document):
        return recette.run_recette(model, recette.RecetteSettings(**values),
                                   model_path=write_model_file(model_content))
