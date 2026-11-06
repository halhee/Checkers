"""Mode hors ligne (CHECKERS_OFFLINE=true) : aucun appel vers Internet.

Pensé pour un déploiement sur réseau isolé ou sensible, où ni le navigateur
ni le serveur ne doivent contacter de service externe. Le mode hors ligne :

- sert les bibliothèques front (Bootstrap, Chart.js, Leaflet, visionneuse
  3D) depuis static/vendor/ au lieu des CDN, après installation par
  scripts/vendor_assets.py sur une machine connectée. Une bibliothèque
  absente est désactivée : jamais de repli silencieux vers le CDN ;
- n'affiche aucun fond de carte externe. Le fond OpenStreetMap enverrait
  les coordonnées du site au serveur de tuiles. Un serveur de tuiles
  interne peut être déclaré avec MAP_TILE_URL ;
- interdit l'escalade des agents IA vers le modèle cloud. Seul le modèle
  local (Ollama, OLLAMA_BASE_URL) reste utilisable ;
- coupe la télémétrie des bibliothèques d'IA (CrewAI, OpenTelemetry,
  LiteLLM) ;
- envoie au navigateur une politique de sécurité du contenu (CSP) qui
  bloque toute requête vers une autre origine que le serveur Checkers (et
  le serveur de tuiles déclaré). Même un oubli dans le code serait donc
  bloqué par le navigateur.

Sans CHECKERS_OFFLINE, le comportement historique est conservé (CDN et
OpenStreetMap), avec une seule différence : une bibliothèque déjà installée
dans static/vendor/ est servie localement.
"""
import os
from urllib.parse import urlsplit

_TRUE = {"1", "true", "yes", "oui", "on"}


def _flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() in _TRUE


OFFLINE = _flag("CHECKERS_OFFLINE")

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
VENDOR_DIR = os.path.join(STATIC_DIR, "vendor")

OSM_TILE_URL = "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png"
OSM_ATTRIBUTION = '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'

# Bibliothèques front. `local` : chemin sous static/vendor/ (écrit par
# scripts/vendor_assets.py) ; `cdn` : URL historique ; `integrity` : empreinte
# SRI, vérifiée par le navigateur dans les deux cas, et par le script
# d'installation au téléchargement.
ASSETS = {
    "bootstrap_css": {
        "local": "bootstrap/3.3.7/bootstrap.min.css",
        "cdn": "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css",
        "integrity": None,
    },
    "chartjs": {
        "local": "chart.js/4.4.1/chart.umd.js",
        "cdn": "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js",
        "integrity": "sha384-dug+JxfBvklEQdJ4AYuBBAIScUz0bVN73xpy273gcAwHjb3qI0fXmuYNaNfdyYJG",
    },
    "leaflet_css": {
        "local": "leaflet/1.9.4/leaflet.css",
        "cdn": "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css",
        "integrity": "sha384-sHL9NAb7lN7rfvG5lfHpm643Xkcjzp4jFvuavGOndn6pjVqS6ny56CAt3nsEVT4H",
    },
    "leaflet_js": {
        "local": "leaflet/1.9.4/leaflet.js",
        "cdn": "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js",
        "integrity": "sha384-cxOPjt7s7Iz04uaHJceBmS+qpjv2JkIHNVcuOrM+YHwZOmJGBXI00mdUXEq65HTH",
    },
    # Visionneuse 3D : module ESM regroupé (le +esm de jsdelivr importe
    # lui-même d'autres URL), worker des fragments (getWorker() le télécharge
    # sinon depuis unpkg) et fichiers WebAssembly de web-ifc.
    "viewer_module": {
        "local": "thatopen/components-3.4.8.mjs",
        "cdn": "https://cdn.jsdelivr.net/npm/@thatopen/components@3.4.8/+esm",
        "integrity": None,
    },
    "viewer_worker": {
        "local": "thatopen/fragments-worker.mjs",
        "cdn": None,  # None = FragmentsManager.getWorker() (unpkg)
        "integrity": None,
    },
    "viewer_wasm": {
        "local": "web-ifc/0.0.77/web-ifc.wasm",
        "cdn": "https://unpkg.com/web-ifc@0.0.77/web-ifc.wasm",
        "integrity": None,
    },
}


def is_vendored(name: str) -> bool:
    return os.path.isfile(os.path.join(VENDOR_DIR, ASSETS[name]["local"]))


def resolve_assets(static_url, offline: bool = None) -> dict:
    """{nom: {"url", "integrity", "local"} ou None}. `static_url(chemin)`
    construit l'URL d'un fichier statique (url_for('static', ...)). None =
    bibliothèque indisponible (mode hors ligne sans copie locale) : la page
    désactive alors la fonctionnalité correspondante."""
    offline = OFFLINE if offline is None else offline
    resolved = {}
    for name, spec in ASSETS.items():
        if is_vendored(name):
            resolved[name] = {"url": static_url("vendor/" + spec["local"]),
                              "integrity": spec["integrity"], "local": True}
        elif offline:
            resolved[name] = None
        else:
            resolved[name] = {"url": spec["cdn"], "integrity": spec["integrity"], "local": False}
    return resolved


def viewer_config(assets: dict) -> dict:
    """Configuration de la visionneuse 3D pour la page. En mode hors ligne,
    la visionneuse n'est activée que si ses trois composants sont installés
    localement."""
    module, worker, wasm = assets["viewer_module"], assets["viewer_worker"], assets["viewer_wasm"]
    if module is None or wasm is None or (worker is None and OFFLINE):
        return {"enabled": False, "worker_url": None, "wasm_path": None}
    wasm_path = wasm["url"].rsplit("/", 1)[0] + "/"
    return {"enabled": True, "worker_url": worker["url"] if worker else None, "wasm_path": wasm_path}


def map_config(offline: bool = None) -> dict:
    """Fond de carte de l'onglet Géoréférencement. MAP_TILE_URL (modèle
    Leaflet, ex. https://tuiles.intranet/{z}/{x}/{y}.png) désigne un serveur
    interne ; sinon OpenStreetMap hors mode hors ligne, et aucun fond en
    mode hors ligne (marqueur et coordonnées seulement)."""
    offline = OFFLINE if offline is None else offline
    url = os.environ.get("MAP_TILE_URL", "").strip() or None
    if url:
        return {"tile_url": url, "attribution": os.environ.get("MAP_TILE_ATTRIBUTION", "")}
    if offline:
        return {"tile_url": None, "attribution": ""}
    return {"tile_url": OSM_TILE_URL, "attribution": OSM_ATTRIBUTION}


def _origin(url_template: str) -> str:
    """Origine (schéma://hôte[:port]) d'un modèle d'URL de tuiles, le
    sous-domaine Leaflet {s} devenant un joker CSP."""
    parts = urlsplit(url_template.replace("{s}.", "subdomain-placeholder."))
    host = parts.netloc.replace("subdomain-placeholder.", "*.")
    return f"{parts.scheme}://{host}"


def content_security_policy(tile_url: str = None) -> str:
    """CSP du mode hors ligne : uniquement la même origine (plus le serveur
    de tuiles interne s'il est déclaré). 'unsafe-inline' est nécessaire au
    code JavaScript et aux styles intégrés à la page ; 'wasm-unsafe-eval'
    et blob: servent à la visionneuse 3D (WebAssembly, workers)."""
    img = "'self' data: blob:" + (f" {_origin(tile_url)}" if tile_url else "")
    return "; ".join([
        "default-src 'self'",
        "script-src 'self' 'unsafe-inline' 'wasm-unsafe-eval'",
        "style-src 'self' 'unsafe-inline'",
        f"img-src {img}",
        "font-src 'self' data:",
        "connect-src 'self' blob: data:",
        "worker-src 'self' blob:",
        "object-src 'none'",
        "base-uri 'self'",
        "form-action 'self'",
        "frame-ancestors 'none'",
    ])


def cloud_llm_allowed() -> bool:
    """L'escalade vers le modèle cloud (agents/llm_tiers.py) est interdite
    en mode hors ligne, même si une clé est configurée."""
    return not OFFLINE


def apply_process_environment() -> None:
    """Coupe la télémétrie des bibliothèques d'IA. À appeler au démarrage,
    avant tout import de CrewAI. setdefault : un administrateur peut
    toujours forcer explicitement une valeur."""
    if not OFFLINE:
        return
    # Variables documentées par CrewAI (docs.crewai.com/telemetry) : la
    # télémétrie anonyme et le traçage vers la plateforme CrewAI sont deux
    # mécanismes distincts, à couper séparément.
    os.environ.setdefault("CREWAI_DISABLE_TELEMETRY", "true")
    os.environ.setdefault("OTEL_SDK_DISABLED", "true")
    os.environ.setdefault("CREWAI_TRACING_ENABLED", "false")
    # LiteLLM (utilisé par certains fournisseurs de CrewAI) télécharge sinon
    # sa table de prix des modèles depuis GitHub à l'import.
    os.environ.setdefault("LITELLM_LOCAL_MODEL_COST_MAP", "True")


def startup_report() -> str:
    lines = [f"Mode hors ligne : {'ACTIVÉ' if OFFLINE else 'désactivé'}"]
    for name in ASSETS:
        state = "local" if is_vendored(name) else ("DÉSACTIVÉ (non installé)" if OFFLINE else "CDN")
        lines.append(f"  - {name} : {state}")
    tiles = map_config()["tile_url"]
    lines.append(f"  - fond de carte : {tiles or 'aucun'}")
    if OFFLINE:
        lines.append(f"  - escalade IA cloud : interdite ; télémétrie IA : coupée")
    return "\n".join(lines)
