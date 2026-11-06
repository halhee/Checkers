"""Construit le .zip de l'extension, installable dans Blender via
Édition > Préférences > Extensions > (menu ⌄) > Installer depuis le disque.

    python build.py

Équivalent à `blender --command extension build --source-dir checkers_bonsai`
pour qui a Blender sous la main.
"""
import os
import re
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
EXTENSION = os.path.join(HERE, "checkers_bonsai")
DIST = os.path.join(HERE, "dist")


def version() -> str:
    with open(os.path.join(EXTENSION, "blender_manifest.toml"), encoding="utf-8") as f:
        return re.search(r'^version\s*=\s*"([^"]+)"', f.read(), re.M).group(1)


def build() -> str:
    os.makedirs(DIST, exist_ok=True)
    zip_path = os.path.join(DIST, f"checkers_bonsai-{version()}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(EXTENSION):
            dirs[:] = sorted(d for d in dirs if d != "__pycache__")
            for file in sorted(files):
                if not file.endswith(".pyc"):
                    path = os.path.join(root, file)
                    zf.write(path, os.path.relpath(path, EXTENSION))  # manifeste à la racine de l'archive
    return zip_path


if __name__ == "__main__":
    print(build())
