"""Doublure minimale d'IfcOpenShell pour les tests (voir tests/fakes/README.md)."""
import itertools

# Hiérarchie de classes utilisée par les tests (IFC4 / IFC4X3).
PARENTS = {
    "IfcWallStandardCase": "IfcWall", "IfcWall": "IfcBuildingElement", "IfcDoor": "IfcBuildingElement",
    "IfcSlab": "IfcBuildingElement", "IfcBeam": "IfcBuildingElement", "IfcColumn": "IfcBuildingElement",
    "IfcBuildingElement": "IfcElement", "IfcPump": "IfcFlowMovingDevice", "IfcFlowMovingDevice": "IfcElement",
    "IfcElement": "IfcProduct",
    "IfcBuildingStorey": "IfcSpatialStructureElement", "IfcBuilding": "IfcFacility", "IfcFacility": "IfcSpatialStructureElement",
    "IfcMarineFacility": "IfcFacility", "IfcMarinePart": "IfcFacilityPart", "IfcFacilityPart": "IfcSpatialStructureElement",
    "IfcSite": "IfcSpatialStructureElement", "IfcSpace": "IfcSpatialStructureElement",
    "IfcSpatialStructureElement": "IfcProduct", "IfcProduct": "IfcRoot", "IfcProject": "IfcRoot",
}
_ids = itertools.count(1)
version = "test"


class entity_instance:
    def __init__(self, cls, **attrs):
        self._cls, self._id = cls, next(_ids)
        self.ContainedInStructure, self.Decomposes, self.IsDefinedBy = [], [], []
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

    def add(self, cls, **attrs):
        entity = entity_instance(cls, **attrs)
        self.entities.append(entity)
        return entity

    def by_type(self, name):
        if name not in PARENTS and name not in PARENTS.values():
            raise RuntimeError(f"Entity with name '{name}' not found in schema '{self.schema}'")
        return [e for e in self.entities if e.is_a(name)]

    def by_guid(self, guid):
        for e in self.entities:
            if getattr(e, "GlobalId", None) == guid:
                return e
        raise RuntimeError(guid)
