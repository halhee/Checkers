"""Calcul du bilan carbone d'un modèle IFC.

Principes, en réaction aux défauts du calcul précédent :
- Le matériau de chaque élément est lu depuis les données matériau réelles du
  fichier IFC (IfcRelAssociatesMaterial et ses variantes) en priorité. Le
  mapping par classe IFC (ex. "un IfcWall est toujours en béton") n'est utilisé
  qu'en tout dernier recours, quand le fichier ne porte aucune information
  matériau pour l'élément -- et c'est tracé explicitement (`material_source`).
- Les facteurs d'émission ci-dessous sont des ordres de grandeur indicatifs
  (littérature générale, pas une extraction d'une base FDES/INIES réelle).
  Cette table est volontairement simple à éditer si de vraies données
  deviennent disponibles.
- Chaque matériau a un seul type de quantité pertinent (volume OU surface),
  jamais les deux -- ça évite d'avoir à "convertir" un facteur volumique en
  facteur surfacique par une division arbitraire (l'ancien code divisait par
  10, sans justification physique).
- Quand ni les quantités du PSet ni la géométrie ne permettent de déterminer
  une quantité, l'élément est compté comme "non calculable" (carbone = 0),
  jamais remplacé par une valeur par défaut inventée.
"""

import difflib
from collections import defaultdict
from dataclasses import dataclass
from typing import Optional

import ifcopenshell
import ifcopenshell.geom as geom
import ifcopenshell.util.element as element_util
import ifcopenshell.util.shape as shape_util
import ifcopenshell.util.unit as unit_util
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from excel_safety import excel_safe

# Un seul jeu de settings de géométrie, réutilisé pour tous les éléments.
# Volontairement SANS `USE_PYTHON_OPENCASCADE` : cette option nécessite une
# compilation d'ifcopenshell avec OpenCASCADE, absente de l'installation
# (voir requirements.txt/Dockerfile). Le moteur de triangulation natif
# d'ifcopenshell (utilisé ici) fonctionne avec le wheel PyPI standard et
# suffit pour calculer volume/surface depuis le maillage.
_GEOM_SETTINGS = geom.settings()

# Facteurs d'émission indicatifs (ordres de grandeur de la littérature
# générale sur la construction), PAS une extraction d'une base FDES/INIES
# réelle -- à affiner si une vraie base carbone est branchée un jour.
# `quantity_kind` fixe une fois pour toutes si ce matériau se mesure en
# volume ou en surface, pour ne jamais avoir à "convertir" un facteur.
MATERIAL_CARBON_FACTORS = {
    "Béton": {
        "quantity_kind": "volume",
        "factor": 320,  # kg CO2e/m³
        "description": "Béton armé standard",
        "source": "Valeur indicative (à affiner avec une base FDES/INIES réelle)",
        "details": "Inclut la production du ciment, des agrégats, le transport et la mise en œuvre",
    },
    "Acier": {
        "quantity_kind": "volume",
        "factor": 12000,  # kg CO2e/m³
        "description": "Acier de construction",
        "source": "Valeur indicative (à affiner avec une base FDES/INIES réelle)",
        "details": "Inclut la production primaire, le laminage et le transport",
    },
    "Aluminium": {
        "quantity_kind": "volume",
        "factor": 20000,  # kg CO2e/m³ -- plus léger que l'acier mais bien plus carboné au kg
        "description": "Aluminium (menuiseries, façades)",
        "source": "Valeur indicative (à affiner avec une base FDES/INIES réelle)",
        "details": "Production primaire (l'aluminium recyclé a un impact très inférieur, non distingué ici), extrusion, anodisation, transport",
    },
    "Bois": {
        "quantity_kind": "volume",
        "factor": -750,  # kg CO2e/m³ (négatif car stockage de carbone biogénique)
        "description": "Bois de construction",
        "source": "Valeur indicative (à affiner avec une base FDES/INIES réelle)",
        "details": "Bois local, inclut l'exploitation forestière, le transport et la transformation ; le solde négatif reflète le carbone stocké, hors fin de vie",
    },
    "Maçonnerie": {
        "quantity_kind": "volume",
        "factor": 220,  # kg CO2e/m³
        "description": "Parpaing, brique, blocs béton",
        "source": "Valeur indicative (à affiner avec une base FDES/INIES réelle)",
        "details": "Maçonnerie courante (agglos creux, briques) ; moins de ciment/armature au m³ qu'un béton armé plein",
    },
    "Plâtre": {
        "quantity_kind": "area",
        "factor": 13,  # kg CO2e/m²
        "description": "Plaque de plâtre (type BA13)",
        "source": "Valeur indicative (à affiner avec une base FDES/INIES réelle)",
        "details": "Cloisons et doublages, inclut production et transport",
    },
    "Verre": {
        "quantity_kind": "area",
        "factor": 250,  # kg CO2e/m²
        "description": "Double vitrage standard",
        "source": "Valeur indicative (à affiner avec une base FDES/INIES réelle)",
        "details": "Inclut la production du verre, l'assemblage et le transport",
    },
    "Isolation": {
        "quantity_kind": "area",
        "factor": 10,  # kg CO2e/m²
        "description": "Isolation thermique standard",
        "source": "Valeur indicative (à affiner avec une base FDES/INIES réelle)",
        "details": "Laine minérale, épaisseur courante, inclut production, transport et mise en œuvre",
    },
}

# Mots-clés (sous-chaînes, en minuscules) reconnus dans un nom de matériau IFC
# réel pour le rattacher à une catégorie ci-dessus. Éditable : si un modèle
# utilise une nomenclature différente, ajouter les mots-clés correspondants.
MATERIAL_KEYWORD_ALIASES = {
    "Béton": ["béton", "beton", "ciment"],
    # "métal"/"metal" volontairement absents : trop générique (matcherait
    # aussi "Métal - Aluminium") -- l'aluminium a sa propre catégorie
    # ci-dessous, un métal non qualifié reste non reconnu plutôt que deviné.
    "Acier": ["acier", "inox", "zinc"],
    "Aluminium": ["aluminium", "alu"],
    "Bois": ["bois", "douglas", "sapin", "chêne", "chene", "pin", "mélèze", "meleze", "osb", "contreplaqué", "contreplaque"],
    "Maçonnerie": ["parpaing", "brique", "agglo", "maçonnerie", "maconnerie", "pierre"],
    "Plâtre": ["plâtre", "platre", "placo", "ba13", "ba 13", "gypse"],
    "Verre": ["verre", "vitrage", "vitre"],
    "Isolation": ["isolant", "isolation", "laine de verre", "laine de roche", "laine minérale", "laine minerale", "polystyrène", "polystyrene", "xps", "mousse", "ouate"],
}

# Dernier recours SEULEMENT : utilisé quand l'élément ne porte AUCUNE donnée
# matériau dans le fichier IFC. Contrairement à l'ancien code, ce n'est plus
# le chemin par défaut -- le vrai matériau du fichier est toujours essayé
# d'abord (voir `resolve_material`).
IFC_CLASS_MATERIAL_FALLBACK = {
    "IfcWall": "Béton",
    "IfcWallStandardCase": "Béton",
    "IfcSlab": "Béton",
    "IfcFooting": "Béton",
    "IfcBeam": "Béton",
    "IfcColumn": "Béton",
    "IfcStair": "Béton",
    "IfcRailing": "Acier",
    "IfcWindow": "Verre",
    "IfcDoor": "Bois",
    "IfcRoof": "Béton",
    "IfcCovering": "Isolation",
}


@dataclass
class ElementCarbonResult:
    global_id: str
    ifc_class: str
    floor: str
    lot: str
    material: str
    material_source: str  # "ifc_material" | "ifc_material_unrecognized" | "class_fallback" | "no_material_no_fallback"
    quantity: float
    quantity_unit: str  # "m³" | "m²" | "-"
    quantity_source: str  # "pset_quantity" | "geometry_mesh" | "pset_quantity_x_layer_thickness" | "geometry_mesh_x_layer_thickness" | "unavailable"
    emission_factor: float
    carbon_kg: float
    reliability: str  # "mesuré" | "estimé" | "non calculable"


def get_length_unit_scale(ifc_file) -> float:
    """Facteur pour convertir les longueurs du fichier IFC en mètres.
    Un fichier modélisé en millimètres a un facteur de 0.001 par ex. -- sans
    cette conversion, un volume calculé serait faux d'un facteur 10^9."""
    try:
        return float(unit_util.calculate_unit_scale(ifc_file))
    except Exception:
        return 1.0


def _keywords_by_specificity() -> list:
    # Toutes les paires (mot-clé, catégorie), triées du mot-clé le plus long
    # (donc le plus spécifique) au plus court -- ainsi une phrase spécifique
    # (ex. "laine de verre" -> Isolation) l'emporte sur un mot-clé générique
    # plus court d'une autre catégorie qui serait aussi une sous-chaîne
    # (ex. "verre" -> Verre). Calculé une fois, pas à chaque appel.
    pairs = [
        (keyword, canonical)
        for canonical, keywords in MATERIAL_KEYWORD_ALIASES.items()
        for keyword in keywords
    ]
    pairs.sort(key=lambda pair: len(pair[0]), reverse=True)
    return pairs


_KEYWORDS_BY_SPECIFICITY = _keywords_by_specificity()


def normalize_material_name(raw_name: Optional[str]) -> Optional[str]:
    if not raw_name:
        return None
    name_l = raw_name.lower()
    for keyword, canonical in _KEYWORDS_BY_SPECIFICITY:
        if keyword in name_l:
            return canonical
    close = difflib.get_close_matches(name_l, [c.lower() for c in MATERIAL_CARBON_FACTORS], n=1, cutoff=0.75)
    if close:
        for canonical in MATERIAL_CARBON_FACTORS:
            if canonical.lower() == close[0]:
                return canonical
    return None


def get_material_layers(element) -> Optional[list]:
    """Retourne les composants matériau de l'élément tels que déclarés dans
    le fichier IFC : une liste de {"material_raw": str|None, "thickness": float|None}.
    `thickness` (en unités du fichier, non convertie) n'est renseigné que pour
    les parois multicouches (IfcMaterialLayerSet*), où chaque couche a sa
    propre épaisseur. Retourne None si l'élément ne porte aucune donnée
    matériau dans le fichier."""
    try:
        material = element_util.get_material(element)
    except Exception:
        material = None
    if material is None:
        return None

    try:
        if material.is_a("IfcMaterial"):
            return [{"material_raw": material.Name, "thickness": None}]

        if material.is_a("IfcMaterialList"):
            # IFC2X3 : pas de répartition par épaisseur/fraction disponible ;
            # on ne retient que le premier matériau déclaré, en le signalant.
            materials = material.Materials
            if not materials:
                return None
            return [{"material_raw": materials[0].Name, "thickness": None}]

        if material.is_a("IfcMaterialLayerSetUsage"):
            layers = material.ForLayerSet.MaterialLayers
            return [
                {"material_raw": layer.Material.Name if layer.Material else None, "thickness": layer.LayerThickness}
                for layer in layers
            ]

        if material.is_a("IfcMaterialLayerSet"):
            layers = material.MaterialLayers
            return [
                {"material_raw": layer.Material.Name if layer.Material else None, "thickness": layer.LayerThickness}
                for layer in layers
            ]

        if material.is_a("IfcMaterialConstituentSet"):
            constituents = material.MaterialConstituents or []
            return [
                {"material_raw": c.Material.Name if c.Material else None, "thickness": None}
                for c in constituents
            ]
    except Exception:
        return None

    # Structure matériau non gérée (ex. IfcMaterialProfileSet) : pas de
    # support dédié dans ce prototype, on retombe sur le fallback par classe.
    return None


def resolve_material(material_raw: Optional[str], ifc_class: str) -> tuple:
    """Retourne (nom_matériau_affiché, material_source)."""
    if material_raw:
        canonical = normalize_material_name(material_raw)
        if canonical:
            return canonical, "ifc_material"
        return material_raw, "ifc_material_unrecognized"

    fallback = IFC_CLASS_MATERIAL_FALLBACK.get(ifc_class)
    if fallback:
        return fallback, "class_fallback"
    return "Non spécifié", "no_material_no_fallback"


def _get_pset_quantity(element, quantity_kind: str) -> Optional[float]:
    wanted = "IfcQuantityVolume" if quantity_kind == "volume" else "IfcQuantityArea"
    if not hasattr(element, "IsDefinedBy"):
        return None
    for rel in element.IsDefinedBy:
        if not rel.is_a("IfcRelDefinesByProperties"):
            continue
        pdef = rel.RelatingPropertyDefinition
        if not pdef.is_a("IfcElementQuantity"):
            continue
        for q in pdef.Quantities:
            if q.is_a(wanted):
                return float(q.VolumeValue if quantity_kind == "volume" else q.AreaValue)
    return None


def _get_geometry_quantity(element, quantity_kind: str) -> Optional[float]:
    try:
        shape = geom.create_shape(_GEOM_SETTINGS, element)
    except Exception:
        return None
    try:
        if quantity_kind == "volume":
            return float(shape_util.get_volume(shape.geometry))
        # Surface totale du maillage : pour un élément fin (vitrage, cloison),
        # ceci compte les deux faces et surestime probablement la vraie
        # surface "utile" -- reliability est dégradée en conséquence par
        # l'appelant plutôt que d'appliquer un facteur correctif arbitraire.
        return float(shape_util.get_area(shape.geometry))
    except Exception:
        return None


def get_element_quantity(element, quantity_kind: str, unit_scale: float) -> tuple:
    """Retourne (quantité en m ou m², source) où source est
    "pset_quantity" | "geometry_mesh" | None si rien n'a fonctionné."""
    pset_value = _get_pset_quantity(element, quantity_kind)
    if pset_value is not None:
        power = 3 if quantity_kind == "volume" else 2
        return pset_value * (unit_scale ** power), "pset_quantity"

    geom_value = _get_geometry_quantity(element, quantity_kind)
    if geom_value is not None:
        power = 3 if quantity_kind == "volume" else 2
        return geom_value * (unit_scale ** power), "geometry_mesh"

    return None, "unavailable"


def _unavailable_result(global_id, ifc_class, floor, lot, material, material_source) -> ElementCarbonResult:
    return ElementCarbonResult(
        global_id=global_id, ifc_class=ifc_class, floor=floor, lot=lot,
        material=material, material_source=material_source,
        quantity=0.0, quantity_unit="-", quantity_source="unavailable",
        emission_factor=0.0, carbon_kg=0.0, reliability="non calculable",
    )


def calculate_element_carbon(element, floor: str, unit_scale: float, lot: str = "default") -> list:
    """Calcule l'empreinte carbone d'un élément IFC. Retourne une liste
    d'ElementCarbonResult : un seul élément pour un matériau simple, plusieurs
    (un par couche) pour une paroi multicouche. `lot` identifie le fichier IFC
    d'origine dans une analyse multilot (voir Checkers.py::process_files)."""
    ifc_class = element.is_a()
    global_id = getattr(element, "GlobalId", "")
    layers = get_material_layers(element)
    has_thickness_layers = bool(layers) and any(l["thickness"] is not None for l in layers)

    if has_thickness_layers:
        area_m2, area_source = get_element_quantity(element, "area", unit_scale)
        if area_m2 is None:
            return [_unavailable_result(global_id, ifc_class, floor, lot, "Multicouche", "unavailable")]

        results = []
        for layer in layers:
            material, material_source = resolve_material(layer["material_raw"], ifc_class)
            factor_info = MATERIAL_CARBON_FACTORS.get(material)
            thickness_m = (layer["thickness"] or 0.0) * unit_scale
            volume_m3 = area_m2 * thickness_m
            if factor_info is None or factor_info["quantity_kind"] != "volume":
                # Couche non reconnue, ou matériau normalement mesuré en
                # surface (rare pour une couche épaisse) : non calculable
                # pour cette couche plutôt que d'inventer un facteur.
                results.append(ElementCarbonResult(
                    global_id=global_id, ifc_class=ifc_class, floor=floor, lot=lot,
                    material=material, material_source=material_source,
                    quantity=volume_m3, quantity_unit="m³",
                    quantity_source=f"{area_source}_x_layer_thickness",
                    emission_factor=0.0, carbon_kg=0.0, reliability="non calculable",
                ))
                continue
            factor = factor_info["factor"]
            reliability = "mesuré" if area_source == "pset_quantity" else "estimé"
            results.append(ElementCarbonResult(
                global_id=global_id, ifc_class=ifc_class, floor=floor, lot=lot,
                material=material, material_source=material_source,
                quantity=volume_m3, quantity_unit="m³",
                quantity_source=f"{area_source}_x_layer_thickness",
                emission_factor=factor, carbon_kg=volume_m3 * factor, reliability=reliability,
            ))
        return results

    material_raw = layers[0]["material_raw"] if layers else None
    material, material_source = resolve_material(material_raw, ifc_class)
    factor_info = MATERIAL_CARBON_FACTORS.get(material)
    if factor_info is None:
        return [_unavailable_result(global_id, ifc_class, floor, lot, material, material_source)]

    quantity_kind = factor_info["quantity_kind"]
    quantity, quantity_source = get_element_quantity(element, quantity_kind, unit_scale)
    if quantity is None:
        return [_unavailable_result(global_id, ifc_class, floor, lot, material, material_source)]

    factor = factor_info["factor"]
    reliability = "mesuré" if quantity_source == "pset_quantity" else "estimé"
    return [ElementCarbonResult(
        global_id=global_id, ifc_class=ifc_class, floor=floor, lot=lot,
        material=material, material_source=material_source,
        quantity=quantity, quantity_unit="m³" if quantity_kind == "volume" else "m²",
        quantity_source=quantity_source,
        emission_factor=factor, carbon_kg=quantity * factor, reliability=reliability,
    )]


def aggregate_carbon_results(results: list) -> dict:
    total = sum(r.carbon_kg for r in results)
    by_type = defaultdict(float)
    by_floor = defaultdict(float)
    by_material = defaultdict(float)
    by_lot = defaultdict(float)
    # Quantité réelle (m³ ou m², selon le quantity_kind du matériau) par
    # matériau reconnu -- sert à calculer des scénarios de substitution
    # chiffrés (voir compute_substitution_scenarios). Les résultats "non
    # calculable" ou dont le matériau n'est pas dans MATERIAL_CARBON_FACTORS
    # sont ignorés ici, comme pour les autres totaux.
    quantity_by_material = defaultdict(float)
    reliability_by_type = defaultdict(lambda: {"mesuré": 0, "estimé": 0, "non calculable": 0})
    reliability_by_lot = defaultdict(lambda: {"mesuré": 0, "estimé": 0, "non calculable": 0})
    reliability_counts = {"mesuré": 0, "estimé": 0, "non calculable": 0}
    # Un élément multicouche produit plusieurs ElementCarbonResult (un par
    # couche) partageant le même global_id -- on déduplique pour ne compter
    # chaque élément IFC qu'une fois par type/lot. (lot, global_id) plutôt que
    # global_id seul : deux lots distincts pourraient en théorie réutiliser un
    # GlobalId (peu probable en pratique, mais pas garanti par la norme).
    element_ids_by_type = defaultdict(set)
    element_ids_by_lot = defaultdict(set)

    for r in results:
        by_type[r.ifc_class] += r.carbon_kg
        by_floor[r.floor] += r.carbon_kg
        by_material[r.material] += r.carbon_kg
        by_lot[r.lot] += r.carbon_kg
        reliability_by_type[r.ifc_class][r.reliability] += 1
        reliability_by_lot[r.lot][r.reliability] += 1
        reliability_counts[r.reliability] += 1
        element_ids_by_type[r.ifc_class].add(r.global_id)
        element_ids_by_lot[r.lot].add((r.lot, r.global_id))
        if r.material in MATERIAL_CARBON_FACTORS:
            quantity_by_material[r.material] += r.quantity

    n = len(results) or 1
    reliability_summary = {
        "pct_mesure": reliability_counts["mesuré"] / n * 100,
        "pct_estime": reliability_counts["estimé"] / n * 100,
        "pct_non_calculable": reliability_counts["non calculable"] / n * 100,
        "nb_elements_total": n,
        "nb_elements_non_calculables": reliability_counts["non calculable"],
    }

    return {
        "total": total,
        "by_type": dict(by_type),
        "by_floor": dict(by_floor),
        "by_material": dict(by_material),
        "by_lot": dict(by_lot),
        "quantity_by_material": dict(quantity_by_material),
        "reliability_by_type": {k: dict(v) for k, v in reliability_by_type.items()},
        "reliability_by_lot": {k: dict(v) for k, v in reliability_by_lot.items()},
        "element_count_by_type": {k: len(v) for k, v in element_ids_by_type.items()},
        "element_count_by_lot": {k: len(v) for k, v in element_ids_by_lot.items()},
        "reliability_summary": reliability_summary,
        "results": results,
    }


def compute_substitution_scenarios(carbon_data: dict, top_n: int = 5) -> list:
    """Scénarios "et si" chiffrés : pour les matériaux à plus fort impact
    (facteur positif -- pas de scénario pour un matériau déjà négatif comme
    le Bois), compare à l'alternative la moins carbonée de la table parmi
    les matériaux mesurés de la même façon (volume vs volume, surface vs
    surface -- jamais de conversion croisée). Chaque chiffre retourné est un
    produit quantité réelle × facteur de la table, aucun n'est inventé.
    Comparaison d'ordre de grandeur, pas une préconisation d'ingénierie
    prête à l'emploi (ex. la faisabilité structurelle n'est pas évaluée)."""
    by_material = carbon_data.get("by_material", {})
    quantity_by_material = carbon_data.get("quantity_by_material", {})

    candidates = [
        (mat, kg) for mat, kg in by_material.items()
        if mat in MATERIAL_CARBON_FACTORS and MATERIAL_CARBON_FACTORS[mat]["factor"] > 0
        and quantity_by_material.get(mat, 0) > 0
    ]
    candidates.sort(key=lambda x: abs(x[1]), reverse=True)

    scenarios = []
    for material, current_kg in candidates[:top_n]:
        info = MATERIAL_CARBON_FACTORS[material]
        quantity_kind = info["quantity_kind"]
        same_kind_alternatives = [
            (alt, alt_info["factor"]) for alt, alt_info in MATERIAL_CARBON_FACTORS.items()
            if alt != material and alt_info["quantity_kind"] == quantity_kind
        ]
        if not same_kind_alternatives:
            continue
        to_material, to_factor = min(same_kind_alternatives, key=lambda x: x[1])
        if to_factor >= info["factor"]:
            continue  # aucune alternative moins carbonée disponible dans la table

        quantity = quantity_by_material[material]
        alternative_kg = quantity * to_factor
        delta_kg = alternative_kg - current_kg
        scenarios.append({
            "from_material": material,
            "to_material": to_material,
            "quantity": quantity,
            "quantity_unit": "m³" if quantity_kind == "volume" else "m²",
            "current_kg": current_kg,
            "alternative_kg": alternative_kg,
            "delta_kg": delta_kg,
            "delta_pct": (delta_kg / current_kg * 100) if current_kg else 0,
        })

    return scenarios


def _reliability_label(counts: dict) -> str:
    total = sum(counts.values()) or 1
    mesure_pct = counts.get("mesuré", 0) / total * 100
    non_calc = counts.get("non calculable", 0)
    if non_calc == total:
        return "Non calculable"
    if mesure_pct == 100:
        return "100 % mesuré"
    if mesure_pct == 0:
        return "Estimé"
    return f"Mixte ({mesure_pct:.0f} % mesuré)"


def create_carbon_footprint_sheet(workbook, carbon_data: dict):
    """Crée l'onglet Empreinte_Carbone : note méthodologique, bilan par type
    d'élément (avec fiabilité), détail par matériau, graphiques."""
    carbon_sheet = workbook.create_sheet("Empreinte_Carbone")

    header_font = Font(bold=True, size=14, color="FFFFFF")
    subheader_font = Font(bold=True, size=12)
    normal_font = Font(size=11)
    note_font = Font(size=10, italic=True, color="555555")

    header_fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
    subheader_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    total_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    warn_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    # --- Titre + note méthodologique ---
    carbon_sheet.merge_cells("A1:F1")
    cell = carbon_sheet.cell(row=1, column=1)
    cell.value = "BILAN CARBONE DU PROJET"
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = header_fill
    carbon_sheet.row_dimensions[1].height = 30

    reliability = carbon_data["reliability_summary"]
    carbon_sheet.merge_cells("A2:F2")
    note_cell = carbon_sheet.cell(row=2, column=1)
    note_cell.value = (
        "Note méthodologique : les facteurs d'émission utilisés sont des valeurs indicatives "
        "(ordres de grandeur), pas une extraction d'une base FDES/INIES réelle -- voir le détail "
        "par matériau plus bas. "
        f"Sur {reliability['nb_elements_total']} éléments : {reliability['pct_mesure']:.0f} % mesurés "
        f"(quantités du modèle), {reliability['pct_estime']:.0f} % estimés (calcul géométrique), "
        f"{reliability['pct_non_calculable']:.0f} % non calculables "
        f"({reliability['nb_elements_non_calculables']} éléments, comptés à 0 kg CO2e, jamais remplacés par une valeur inventée)."
    )
    note_cell.font = note_font
    note_cell.fill = warn_fill
    note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    carbon_sheet.row_dimensions[2].height = 45

    # --- Section 1 : répartition par lot (analyse multilot) ---
    lot_headers = [
        "Lot", "Nombre d'éléments",
        "Empreinte carbone (kg CO2e)", "% du total", "Fiabilité",
    ]
    lot_header_row = 4
    for col, header in enumerate(lot_headers, 1):
        cell = carbon_sheet.cell(row=lot_header_row, column=col)
        cell.value = header
        cell.font = subheader_font
        cell.fill = subheader_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        carbon_sheet.column_dimensions[get_column_letter(col)].width = 25
    carbon_sheet.row_dimensions[lot_header_row].height = 30

    total_carbon_for_lots = carbon_data["total"]
    reliability_by_lot = carbon_data.get("reliability_by_lot", {})
    element_count_by_lot = carbon_data.get("element_count_by_lot", {})
    sorted_lots = sorted(carbon_data.get("by_lot", {}).items(), key=lambda x: x[1], reverse=True)

    row = lot_header_row + 1
    for lot_name, carbon_value in sorted_lots:
        counts = reliability_by_lot.get(lot_name, {})
        percentage = (carbon_value / total_carbon_for_lots * 100) if total_carbon_for_lots else 0
        cells = [
            (lot_name, "left"),
            (element_count_by_lot.get(lot_name, 0), "right"),
            (f"{carbon_value:.2f}", "right"),
            (f"{percentage:.1f}%", "right"),
            (_reliability_label(counts), "left"),
        ]
        for col, (value, align) in enumerate(cells, 1):
            cell = carbon_sheet.cell(row=row, column=col)
            cell.value = excel_safe(value)
            cell.border = thin_border
            cell.font = normal_font
            cell.alignment = Alignment(horizontal=align, vertical="center")
        row += 1
    lot_section_end = row

    # --- Section 2 : bilan par type d'élément ---
    headers = [
        "Type d'élément", "Nombre d'éléments",
        "Empreinte carbone (kg CO2e)", "% du total", "Fiabilité",
    ]
    header_row = lot_section_end + 2
    for col, header in enumerate(headers, 1):
        cell = carbon_sheet.cell(row=header_row, column=col)
        cell.value = header
        cell.font = subheader_font
        cell.fill = subheader_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        carbon_sheet.column_dimensions[get_column_letter(col)].width = 25
    carbon_sheet.row_dimensions[header_row].height = 30

    total_carbon = carbon_data["total"]
    reliability_by_type = carbon_data["reliability_by_type"]
    element_count_by_type = carbon_data["element_count_by_type"]
    sorted_types = sorted(carbon_data["by_type"].items(), key=lambda x: x[1], reverse=True)

    row = header_row + 1
    row_start = row
    for element_type, carbon_value in sorted_types:
        counts = reliability_by_type.get(element_type, {})
        percentage = (carbon_value / total_carbon * 100) if total_carbon else 0
        cells = [
            (element_type, "left"),
            (element_count_by_type.get(element_type, 0), "right"),
            (f"{carbon_value:.2f}", "right"),
            (f"{percentage:.1f}%", "right"),
            (_reliability_label(counts), "left"),
        ]
        for col, (value, align) in enumerate(cells, 1):
            cell = carbon_sheet.cell(row=row, column=col)
            cell.value = excel_safe(value)
            cell.border = thin_border
            cell.font = normal_font
            cell.alignment = Alignment(horizontal=align, vertical="center")
        row += 1
    row_end = row

    carbon_sheet.merge_cells(f"A{row}:B{row}")
    cell = carbon_sheet.cell(row=row, column=1)
    cell.value = "TOTAL"
    cell.font = Font(bold=True, size=12)
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.fill = total_fill
    for col in range(1, 6):
        cell = carbon_sheet.cell(row=row, column=col)
        cell.border = thin_border
        cell.fill = total_fill
        if col == 3:
            cell.value = f"{total_carbon:.2f}"
        elif col == 4:
            cell.value = "100%"
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="right", vertical="center")
    row += 1

    # --- Section 3 : détail des matériaux ---
    row += 2
    carbon_sheet.merge_cells(f"A{row}:F{row}")
    cell = carbon_sheet.cell(row=row, column=1)
    cell.value = "DÉTAILS DES MATÉRIAUX ET FACTEURS D'ÉMISSION"
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.fill = header_fill
    row += 1

    material_headers = ["Matériau", "Description", "Facteur d'émission", "Source", "Détails", "Impact total (kg CO2e)"]
    for col, header in enumerate(material_headers, 1):
        cell = carbon_sheet.cell(row=row, column=col)
        cell.value = header
        cell.font = subheader_font
        cell.fill = subheader_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
    row += 1
    material_row_start = row

    material_totals = carbon_data["by_material"]
    sorted_materials = sorted(
        [(mat, data, material_totals.get(mat, 0)) for mat, data in MATERIAL_CARBON_FACTORS.items()],
        key=lambda x: x[2], reverse=True,
    )
    for material, material_data, total in sorted_materials:
        unit = "m³" if material_data["quantity_kind"] == "volume" else "m²"
        cells = [
            (material, "left"),
            (material_data["description"], "left"),
            (f"{material_data['factor']} kg CO2e/{unit}", "right"),
            (material_data["source"], "left"),
            (material_data["details"], "left"),
            (f"{total:.2f}", "right"),
        ]
        for col, (value, align) in enumerate(cells, 1):
            cell = carbon_sheet.cell(row=row, column=col)
            cell.value = excel_safe(value)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        row += 1
    material_row_end = row

    for row_idx in range(material_row_start, material_row_end):
        max_lines = max(
            len(str(carbon_sheet.cell(row=row_idx, column=col).value).split("\n"))
            for col in range(1, len(material_headers) + 1)
        )
        carbon_sheet.row_dimensions[row_idx].height = max(15 * max_lines, 30)

    # --- Graphiques ---
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "Empreinte carbone par élément"
    chart1.y_axis.title = "Empreinte carbone (kgCO2e)"
    chart1.x_axis.title = "Éléments"
    data = Reference(carbon_sheet, min_col=3, min_row=header_row, max_row=row_end - 1)
    cats = Reference(carbon_sheet, min_col=1, min_row=row_start, max_row=row_end - 1)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.height = 15
    chart1.width = 25
    carbon_sheet.add_chart(chart1, "H4")

    chart2 = PieChart()
    chart2.title = "Répartition de l'empreinte carbone par matériau"
    chart2.style = 10
    pie_data = Reference(carbon_sheet, min_col=6, min_row=material_row_start - 1, max_row=material_row_end - 1)
    pie_labels = Reference(carbon_sheet, min_col=1, min_row=material_row_start, max_row=material_row_end - 1)
    chart2.add_data(pie_data, titles_from_data=True)
    chart2.set_categories(pie_labels)
    chart2.dataLabels = DataLabelList()
    chart2.dataLabels.showPercent = True
    chart2.dataLabels.showVal = False
    chart2.dataLabels.showCatName = True
    chart2.height = 15
    chart2.width = 25
    carbon_sheet.add_chart(chart2, "H22")
