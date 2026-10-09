"""Géoréférencement du site (IfcSite.RefLatitude/RefLongitude), pour la
carte du nouvel onglet "Géoréférencement". Nommé site_location.py (pas
site.py) pour ne pas entrer en conflit avec le module `site` de la
bibliothèque standard Python.
"""

from typing import Optional


def _dms_to_decimal(dms) -> Optional[float]:
    """Convertit un IfcCompoundPlaneAngleMeasure (degrés, minutes, secondes,
    [millionièmes de seconde]) en degrés décimaux. Le signe des degrés porte
    le signe de la mesure entière (IFC : Sud/Ouest sont négatifs)."""
    if not dms:
        return None
    degrees = dms[0]
    minutes = dms[1] if len(dms) > 1 else 0
    seconds = dms[2] if len(dms) > 2 else 0
    micro = dms[3] if len(dms) > 3 else 0
    sign = -1 if degrees < 0 else 1
    return sign * (abs(degrees) + minutes / 60 + seconds / 3600 + micro / (3600 * 1_000_000))


def get_site_location(ifc_file) -> Optional[dict]:
    """Retourne {"name", "latitude", "longitude", "elevation"} pour le
    premier IfcSite géoréférencé du modèle, ou None si le fichier ne porte
    aucune coordonnée (RefLatitude/RefLongitude non renseignées -- fréquent,
    ce n'est pas une obligation IFC)."""
    for ifc_site in ifc_file.by_type("IfcSite"):
        latitude = _dms_to_decimal(getattr(ifc_site, "RefLatitude", None))
        longitude = _dms_to_decimal(getattr(ifc_site, "RefLongitude", None))
        if latitude is None or longitude is None:
            continue
        return {
            "name": ifc_site.Name or "Site",
            "latitude": latitude,
            "longitude": longitude,
            "elevation": getattr(ifc_site, "RefElevation", None),
        }
    return None
