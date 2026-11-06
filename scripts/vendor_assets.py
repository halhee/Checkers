"""Installe localement (static/vendor/) les bibliothèques front de Checkers,
pour un déploiement en mode hors ligne (CHECKERS_OFFLINE=true, voir
offline_config.py).

À lancer sur une machine connectée à Internet, puis transférer le dossier
static/vendor/ (ou l'image Docker construite) sur le réseau isolé :

    python scripts/vendor_assets.py            # Bootstrap, Chart.js, Leaflet
    python scripts/vendor_assets.py --viewer   # + visionneuse 3D (Node.js/npm requis)

Chaque fichier dont l'empreinte SRI est connue (offline_config.ASSETS) est
vérifié après téléchargement : en cas d'écart, rien n'est écrit. Les
versions sont figées ; elles correspondent à celles chargées depuis les CDN
hors mode hors ligne.
"""
import argparse
import base64
import glob
import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import offline_config  # noqa: E402

# Fichiers annexes non listés dans offline_config.ASSETS : images
# référencées par leaflet.css (chemins relatifs images/...).
LEAFLET_IMAGES = ["layers.png", "layers-2x.png", "marker-icon.png", "marker-icon-2x.png", "marker-shadow.png"]
LEAFLET_BASE = "https://unpkg.com/leaflet@1.9.4/dist/"

VIEWER_PACKAGES = {"@thatopen/components": "3.4.8", "web-ifc": "0.0.77"}
WEB_IFC_WASM = ["web-ifc.wasm", "web-ifc-mt.wasm"]


def sri_sha384(data: bytes) -> str:
    return "sha384-" + base64.b64encode(hashlib.sha384(data).digest()).decode()


def download(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "checkers-vendor-assets"})
    with urllib.request.urlopen(request, timeout=60) as response:
        return response.read()


def write(rel_path: str, data: bytes) -> None:
    path = os.path.join(offline_config.VENDOR_DIR, rel_path)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(data)
    print(f"  ✓ static/vendor/{rel_path} ({len(data) // 1024} Ko)")


def fetch_verified(name: str) -> None:
    spec = offline_config.ASSETS[name]
    data = download(spec["cdn"])
    if spec["integrity"]:
        actual = sri_sha384(data)
        if actual != spec["integrity"]:
            raise SystemExit(f"Empreinte inattendue pour {spec['cdn']} :\n  attendue {spec['integrity']}\n"
                             f"  obtenue  {actual}\nRien n'a été écrit pour ce fichier.")
    else:
        print(f"  (pas d'empreinte de référence pour {name} : {sri_sha384(data)})")
    write(spec["local"], data)


def vendor_base() -> None:
    print("Bibliothèques de base :")
    for name in ("bootstrap_css", "chartjs", "leaflet_css", "leaflet_js"):
        fetch_verified(name)
    leaflet_dir = os.path.dirname(offline_config.ASSETS["leaflet_css"]["local"])
    for image in LEAFLET_IMAGES:
        write(f"{leaflet_dir}/images/{image}", download(LEAFLET_BASE + "images/" + image))


def vendor_viewer() -> None:
    """Regroupe @thatopen/components et ses dépendances (three, fragments,
    web-ifc) en un seul module ESM avec esbuild, et copie le worker des
    fragments et les fichiers WebAssembly de web-ifc."""
    print("Visionneuse 3D (npm + esbuild) :")
    npm = shutil.which("npm")
    if npm is None:
        raise SystemExit("npm introuvable : installez Node.js (18+) pour --viewer.")
    with tempfile.TemporaryDirectory() as work:
        packages = [f"{name}@{version}" for name, version in VIEWER_PACKAGES.items()] + ["esbuild@0.25"]
        subprocess.run([npm, "init", "-y"], cwd=work, check=True, stdout=subprocess.DEVNULL)
        subprocess.run([npm, "install", "--no-audit", "--no-fund", *packages], cwd=work, check=True)
        entry = os.path.join(work, "entry.mjs")
        with open(entry, "w", encoding="utf-8") as f:
            f.write('export * from "@thatopen/components";\n')
        bundle = os.path.join(work, "bundle.mjs")
        esbuild = os.path.join(work, "node_modules", ".bin", "esbuild")
        subprocess.run([esbuild, entry, "--bundle", "--format=esm", "--minify", f"--outfile={bundle}"],
                       cwd=work, check=True)
        with open(bundle, "rb") as f:
            write(offline_config.ASSETS["viewer_module"]["local"], f.read())

        workers = sorted(glob.glob(os.path.join(work, "node_modules", "@thatopen", "fragments", "dist", "**",
                                                "*orker*.mjs"), recursive=True))
        if not workers:
            raise SystemExit("Worker des fragments introuvable dans node_modules/@thatopen/fragments/dist.")
        with open(workers[0], "rb") as f:
            write(offline_config.ASSETS["viewer_worker"]["local"], f.read())

        wasm_dir = os.path.dirname(offline_config.ASSETS["viewer_wasm"]["local"])
        for wasm in WEB_IFC_WASM:
            source = os.path.join(work, "node_modules", "web-ifc", wasm)
            if os.path.isfile(source):
                with open(source, "rb") as f:
                    write(f"{wasm_dir}/{wasm}", f.read())
    print("  À valider : ouvrir l'onglet « Maquette 3D » une fois l'appli lancée en mode hors ligne.")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--viewer", action="store_true", help="installer aussi la visionneuse 3D (Node.js requis)")
    args = parser.parse_args()
    vendor_base()
    if args.viewer:
        vendor_viewer()
    print("\nTerminé. Au démarrage, Checkers liste les bibliothèques servies localement.")


if __name__ == "__main__":
    main()
