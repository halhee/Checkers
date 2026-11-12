"""Doublure de bonsai.tool : modèle ouvert et objets Blender associés."""


class Ifc:
    file = None
    path = ""
    objects = {}

    @classmethod
    def get(cls):
        return cls.file

    @classmethod
    def get_path(cls):
        return cls.path

    @classmethod
    def get_object(cls, entity):
        return cls.objects.get(entity.GlobalId)

    @classmethod
    def get_entity(cls, obj):
        try:
            return cls.file.by_guid(obj.name)
        except Exception:
            return None


class Georeference:
    props = None

    @classmethod
    def get_georeference_props(cls):
        if cls.props is None:
            raise AttributeError("pas de propriétés de géoréférencement")
        return cls.props
