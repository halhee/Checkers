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
