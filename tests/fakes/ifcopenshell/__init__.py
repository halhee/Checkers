"""Doublure minimale d'IfcOpenShell pour les tests (voir tests/fakes/README.md)."""
import itertools
import math
import os

# Hiérarchie de classes utilisée par les tests.
PARENTS = {
    "IfcWallStandardCase": "IfcWall", "IfcWall": "IfcBuildingElement", "IfcDoor": "IfcBuildingElement",
    "IfcSlab": "IfcBuildingElement", "IfcBeam": "IfcBuildingElement", "IfcColumn": "IfcBuildingElement",
    "IfcStair": "IfcBuildingElement", "IfcBuildingElementProxy": "IfcBuildingElement",
    "IfcWindow": "IfcBuildingElement", "IfcCovering": "IfcBuildingElement",
    "IfcBuildingElement": "IfcElement",
    "IfcPump": "IfcFlowMovingDevice", "IfcFlowMovingDevice": "IfcDistributionFlowElement",
    "IfcPipeSegment": "IfcFlowSegment", "IfcFlowSegment": "IfcDistributionFlowElement",
    "IfcFlowFitting": "IfcDistributionFlowElement", "IfcDistributionFlowElement": "IfcDistributionElement",
    "IfcDistributionElement": "IfcElement", "IfcTransportElement": "IfcElement",
    "IfcOpeningElement": "IfcFeatureElement", "IfcFeatureElement": "IfcElement",
    "IfcElement": "IfcProduct",
    "IfcBuildingStorey": "IfcSpatialStructureElement", "IfcBuilding": "IfcFacility",
    "IfcFacility": "IfcSpatialStructureElement", "IfcMarineFacility": "IfcFacility",
    "IfcMarinePart": "IfcFacilityPart", "IfcFacilityPart": "IfcSpatialStructureElement",
    "IfcSite": "IfcSpatialStructureElement", "IfcSpace": "IfcSpatialStructureElement",
    "IfcSpatialStructureElement": "IfcProduct", "IfcProduct": "IfcObject", "IfcProject": "IfcObject",
    "IfcSystem": "IfcObject", "IfcObject": "IfcObjectDefinition", "IfcObjectDefinition": "IfcRoot",
    "IfcPropertySet": "IfcRoot", "IfcMapConversion": None, "IfcRoot": None,
    "IfcDuctSegment": "IfcFlowSegment",
    "IfcGeometricRepresentationSubContext": "IfcGeometricRepresentationContext",
    "IfcGeometricRepresentationContext": None,
}
# Classes absentes d'un schéma : by_type() lève une erreur, comme IfcOpenShell.
MISSING = {
    "IFC2X3": {"IfcFacility", "IfcFacilityPart", "IfcMarineFacility", "IfcMarinePart", "IfcMapConversion"},
    "IFC4": {"IfcFacility", "IfcFacilityPart", "IfcMarineFacility", "IfcMarinePart"},
}
_ids = itertools.count(1)
version = "test"
# Maquettes « enregistrées » pour ifcopenshell.open : {chemin absolu: modèle}.
REGISTRY = {}


def register(path, model):
    """Associe un fichier (créé vide sur disque) à un modèle de test, pour
    ifcopenshell.open."""
    path = os.path.abspath(path)
    with open_file(path, "w") as fh:
        fh.write("ISO-10303-21;")
    REGISTRY[path] = model
    return path


open_file = open


def open(path):  # noqa: A001  (même nom que ifcopenshell.open)
    try:
        return REGISTRY[os.path.abspath(path)]
    except KeyError:
        raise OSError(f"Unable to open file for reading: {path}")


class Placement:
    """Placement absolu : matrice 4×4, translation en mètres."""

    def __init__(self, x=0.0, y=0.0, z=0.0, angle=0.0):
        c, s = math.cos(math.radians(angle)), math.sin(math.radians(angle))
        self.matrix = [[c, -s, 0.0, x], [s, c, 0.0, y], [0.0, 0.0, 1.0, z], [0.0, 0.0, 0.0, 1.0]]


def place(entity, x=0.0, y=0.0, z=0.0, angle=0.0):
    entity.ObjectPlacement = Placement(x, y, z, angle)
    return entity


class entity_instance:
    def __init__(self, cls, **attrs):
        self._cls, self._id = cls, next(_ids)
        self.ContainedInStructure, self.Decomposes, self.IsDefinedBy = [], [], []
        self.Representation = "géométrie"
        self._psets = attrs.pop("psets", {})
        self.__dict__.update(attrs)

    def id(self):
        return self._id

    def is_a(self, name=None):
        if name is None:
            return self._cls
        cls = self._cls
        while cls:
            if cls.lower() == name.lower():
                return True
            cls = PARENTS.get(cls)
        return False

    def __getattr__(self, name):
        if name.startswith("_"):
            raise AttributeError(name)
        return None

    def __hash__(self):
        return self._id

    def __eq__(self, other):
        return isinstance(other, entity_instance) and other._id == self._id


class _Rel:
    def __init__(self, **attrs):
        self.__dict__.update(attrs)


def contain(element, structure):
    element.ContainedInStructure = [_Rel(RelatingStructure=structure)]


def aggregate(child, parent):
    child.Decomposes = [_Rel(RelatingObject=parent)]
    parent.IsDecomposedBy = (parent.IsDecomposedBy or []) + [_Rel(RelatedObjects=[child])]


class _FileName:
    originating_system = "Outil auteur de test"
    author = ("Jean Dupont",)
    organization = ("Bureau d'études",)
    time_stamp = "2026-11-01T10:00:00"


class _Header:
    file_name = _FileName()


class file:
    def __init__(self, schema="IFC4"):
        self.schema = self.schema_identifier = schema
        self.header = _Header()
        self.entities = []
        self.length_unit = "METRE"      # lu par util.unit.get_project_unit
        self.unit_scale = 1.0           # lu par util.unit.calculate_unit_scale

    def add(self, cls, **attrs):
        entity = entity_instance(cls, **attrs)
        self.entities.append(entity)
        return entity

    def by_type(self, name):
        if name not in PARENTS or name in MISSING.get(self.schema, ()):
            raise RuntimeError(f"Entity with name '{name}' not found in schema '{self.schema}'")
        return [e for e in self.entities if e.is_a(name)]

    def by_guid(self, guid):
        for e in self.entities:
            if getattr(e, "GlobalId", None) == guid:
                return e
        raise RuntimeError(guid)
