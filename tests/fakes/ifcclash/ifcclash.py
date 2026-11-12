"""Doublure d'ifcclash.ifcclash : même interface que le vrai Clasher
(IfcOpenShell 0.9), mais la géométrie est une boîte alignée sur les axes,
portée par l'attribut `_box` des entités de test : ((x0, y0, z0), (x1, y1, z1)).
"""
import ifcopenshell
import ifcopenshell.util.selector


class ClashSettings:
    def __init__(self):
        self.logger = None
        self.output = "clashes.json"


class Clasher:
    def __init__(self, settings):
        self.settings = settings
        self.clash_sets = []
        self.ifcs = {}

    def load_ifc(self, path):
        if path not in self.ifcs:
            self.ifcs[path] = ifcopenshell.open(path)
        return self.ifcs[path]

    def _elements(self, source):
        ifc_file = self.load_ifc(source["file"])
        if source.get("mode") == "i" and source.get("selector"):
            elements = ifcopenshell.util.selector.filter_elements(ifc_file, source["selector"])
        else:
            elements = set(ifc_file.by_type("IfcElement")) - set(ifc_file.by_type("IfcFeatureElement"))
        boxed = [e for e in elements if getattr(e, "_box", None)]
        # Le vrai Clasher échoue (assert) quand rien n'a de géométrie.
        assert boxed, "iterator.initialize() a échoué"
        return boxed

    def clash(self):
        for clash_set in self.clash_sets:
            a = [e for source in clash_set["a"] for e in self._elements(source)]
            b = [e for source in clash_set.get("b") or clash_set["a"] for e in self._elements(source)]
            clashes = {}
            for x in a:
                for y in b:
                    found = _compare(x._box, y._box, clash_set)
                    if found:
                        kind, distance, point = found
                        clashes[f"{x.GlobalId}-{y.GlobalId}"] = {
                            "a_global_id": x.GlobalId, "b_global_id": y.GlobalId, "a_ifc_class": x.is_a(),
                            "b_ifc_class": y.is_a(), "a_name": x.Name, "b_name": y.Name, "type": kind,
                            "p1": list(point), "p2": list(point), "distance": distance}
            clash_set["clashes"] = clashes


def _compare(box_a, box_b, clash_set):
    (a0, a1), (b0, b1) = box_a, box_b
    overlaps = [min(a1[i], b1[i]) - max(a0[i], b0[i]) for i in range(3)]
    point = tuple((max(a0[i], b0[i]) + min(a1[i], b1[i])) / 2 for i in range(3))
    depth = min(overlaps)
    mode = clash_set["mode"]
    if mode == "intersection":
        return ("protrusion", depth, point) if depth > clash_set["tolerance"] else None
    if mode == "collision":
        return ("collision", depth, point) if depth > 0 or (depth == 0 and not clash_set["allow_touching"]) else None
    gap = max(0.0, -depth)
    return ("clearance", gap, point) if 0 < gap < clash_set["clearance"] else None
