# Checkers pour Bonsai

Extension Blender qui lance le moteur de Checkers (`checkers_core`, le même que l'application web) sur le modèle IFC **ouvert dans Bonsai**, sans export ni serveur.

## Ce que fait le panneau

Vue 3D → barre latérale (`N`) → onglet **Checkers** :

1. **Règles Excel** (même format que l'appli web, multi-corps d'état compris) et/ou **IDS**.
2. **Lot** (par défaut : nom du fichier IFC) et **Corps d'état**. La liste déroulante apparaît si l'Excel contient un onglet `Corps_Etat`. Sur « Auto », seule la correspondance **exacte** du nom de lot est tentée dans Blender, car rapidfuzz et les agents IA n'y sont pas disponibles : choisissez le corps d'état explicitement en cas de doute. Le panneau le signale si le lot se retrouve sans règles.
3. **Lancer l'analyse** : taux de conformité PSet, conformité IDS, bilan carbone et liste des anomalies (agrégées par étage / classe IFC / PSet ou spécification IDS, comme dans l'appli web).
4. Pour chaque anomalie, le bouton de sélection **sélectionne et cadre** les éléments concernés dans la vue 3D.
5. **Colorer** : vert = conforme, rouge = non conforme. Les éléments non contrôlés gardent leur couleur. **Réinitialiser** restaure les couleurs d'origine.
6. **Enregistrer le rapport Excel** : rapport complet (Résumé, Détails, Empreinte carbone, Comparatif, IDS, Fiche de Recette).
7. **Exporter en BCF** : un sujet BCF 2.1 par anomalie, avec les éléments concernés sélectionnés et colorés en rouge, à transmettre aux maîtres d'œuvre et aux entreprises. Le fichier s'ouvre dans BIMcollab, Solibri, Revit/Navisworks (via plugin) ou dans le module BCF de Bonsai. Renseignez **Auteur BCF**, en général une adresse e-mail. Les points de vue n'ont pas de caméra : l'outil de destination cadre lui-même la sélection.

Différences avec l'application web : pas de préconisations IA (CrewAI/Ollama), pas d'historique ni de comparaison avant/après, et un seul lot par analyse (le modèle ouvert).

## Installation

```bash
python bonsai_extension/build.py
# -> bonsai_extension/dist/checkers_bonsai-0.1.0.zip
```

Dans Blender : **Édition → Préférences → Extensions → menu ⌄ (en haut à droite) → Installer depuis le disque…**, choisir le `.zip`. Bonsai doit être installé et activé.

Aucune dépendance n'est embarquée : Bonsai fournit déjà `ifcopenshell`, `ifctester`, `pandas` et `openpyxl`. Le script copie seulement les modules de Checkers utiles au moteur (`checkers_core/`, `carbon.py`, `corps_etat.py`, `ids_check.py`, `excel_safety.py`, `site_location.py`) dans `checkers_bonsai/lib/`. Ce dossier est généré et non versionné : relancez le build après toute modification du moteur.

## Compatibilité

Développé pour Blender 5.2 / Bonsai 0.9 (Python 3.13). Bonsai 0.9 embarque IfcOpenShell 0.9, alors que l'application web reste figée en 0.8.5 (voir `requirements.txt`). Le moteur n'utilise que des API stables (`by_type`, `util.element.get_psets`, `util.element.get_material`, `geom.create_shape`…), mais le contrôle IDS (`ids_check.py`, via `ifctester`) et le calcul de géométrie du bilan carbone doivent être validés sur de vrais modèles avec cette version.

L'analyse s'exécute dans le fil principal de Blender : l'interface reste figée pendant le calcul. Sur un gros modèle, le bilan carbone (géométrie de chaque élément) domine le temps de calcul, comme dans l'application web.

## Structure

```
checkers_bonsai/
├── blender_manifest.toml   # manifeste d'extension Blender (4.2+)
├── __init__.py             # enregistrement
├── engine.py               # lien avec Bonsai et checkers_core (sans bpy, testable hors Blender)
├── operators.py            # analyse, sélection, coloration, export Excel et BCF
├── props.py                # propriétés de scène (fichiers, résultats, anomalies)
├── ui.py                   # panneau et liste des anomalies
└── lib/                    # généré par build.py
```
