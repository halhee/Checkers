"""Construit le .zip de l'extension Bonsai, installable dans Blender via
Édition > Préférences > Extensions > (menu ⌄) > Installer depuis le disque.

    python bonsai_extension/build.py

Copie les modules de Checkers dont le moteur a besoin dans
checkers_bonsai/lib/ (pas de duplication dans Git : lib/ est généré), puis
produit dist/checkers_bonsai-<version>.zip. Aucune dépendance tierce n'est
embarquée : Bonsai fournit déjà ifcopenshell, ifctester, pandas et openpyxl.
"""
import os
import re
import shutil
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
EXTENSION = os.path.join(HERE, "checkers_bonsai")
LIB = os.path.join(EXTENSION, "lib")
DIST = os.path.join(HERE, "dist")

# Modules de Checkers utilisés par checkers_core.analysis.process_files.
# Volontairement absents : Checkers.py (Flask), history.py (historique de
# l'appli web) et agents/ (CrewAI, désactivé dans Blender).
LIB_MODULES = ["checkers_core", "carbon.py", "corps_etat.py", "ids_check.py",
               "excel_safety.py", "site_location.py"]

IGNORE = shutil.ignore_patterns("__pycache__", "*.pyc")


def copy_lib():
    shutil.rmtree(LIB, ignore_errors=True)
    os.makedirs(LIB)
    for name in LIB_MODULES:
        src = os.path.join(REPO, name)
        dst = os.path.join(LIB, name)
        if os.path.isdir(src):
            shutil.copytree(src, dst, ignore=IGNORE)
        else:
            shutil.copy2(src, dst)


def version():
    with open(os.path.join(EXTENSION, "blender_manifest.toml"), encoding="utf-8") as f:
        return re.search(r'^version\s*=\s*"([^"]+)"', f.read(), re.M).group(1)


def build_zip() -> str:
    os.makedirs(DIST, exist_ok=True)
    zip_path = os.path.join(DIST, f"checkers_bonsai-{version()}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(EXTENSION):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            for file in files:
                if file.endswith(".pyc"):
                    continue
                path = os.path.join(root, file)
                # Le manifeste doit être à la racine de l'archive.
                zf.write(path, os.path.relpath(path, EXTENSION))
    return zip_path


if __name__ == "__main__":
    copy_lib()
    print(build_zip())
    sys.exit(0)
