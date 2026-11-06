# Checkers pour Bonsai – recette de maquettes IFC

Extension Blender pour [Bonsai](https://bonsaibim.org/) qui fait la **recette d'une maquette IFC** directement sur le modèle ouvert : contrôle des exigences (IDS et/ou règles Excel), verdict, **fiche de recette à signer** et **export BCF** des anomalies vers la maîtrise d'œuvre et les entreprises.

Conçue pour fonctionner sur un poste sans accès Internet :
- **Entièrement hors ligne.** L'extension ne fait aucun accès réseau et ne demande pas la permission « network » à Blender (vérifié par les tests). Les IDS sont validés avec les schémas embarqués par Bonsai.
- **Sans dépendance ajoutée.** Elle utilise ifcopenshell, ifctester et openpyxl, fournis par Bonsai.
- **Sans IA ni service tiers.**

## Utilisation

Ouvrez un projet IFC dans Bonsai, puis, dans la vue 3D, ouvrez la barre latérale (`N`) et l'onglet **Checkers** :

1. **Référentiels**
   - **IDS** : les exigences du cahier des charges BIM, au format buildingSMART. Le bouton à droite du champ propose deux IDS d'exemple : exploitation-maintenance (DOE / GMAO) et conception minimale.
   - **Règles Excel** (facultatif) : format historique de Checkers. Si le fichier déclare des corps d'état, choisissez celui du lot.
2. **Recette** : lot, phase (ESQ… DOE, exploitation), personne qui recette, **mention de confidentialité** facultative (texte libre) et seuils de verdict (98 % / 90 % par défaut). Puis **Lancer la recette**.
3. **Résultats** : verdict, taux de conformité IDS et Excel, et liste des anomalies regroupées par niveau, classe IFC et exigence. Le bouton de sélection d'une anomalie sélectionne et cadre les éléments concernés. **Colorer** passe les éléments contrôlés en vert ou rouge, les autres ne changent pas ; **Réinitialiser** restaure les couleurs d'origine.
4. **Livrables**
   - **Fiche de recette (Excel)** :
     - onglet *Fiche de recette*, prêt à imprimer et signer : identification, résultat et verdict, anomalies ouvertes, traçabilité, zone de décision et de visa ;
     - onglets *IDS* et *Détail PSet* : le détail élément par élément.
   - **Anomalies (BCF)** : un sujet BCF 2.1 par anomalie, avec les éléments sélectionnés et colorés. Il s'ouvre dans BIMcollab, Solibri, Revit/Navisworks (via plugin) ou dans le module BCF de Bonsai.

### Verdict

Chaque contrôle (IDS, Excel) donne un verdict selon son taux de conformité : *validé* au-dessus du premier seuil, *validé avec réserves* au-dessus du second, *rejeté* en dessous. Une spécification IDS obligatoire qui ne trouve aucun élément entraîne le rejet. Le verdict global est le plus sévère des deux.

### Traçabilité

La fiche indique, pour pouvoir prouver ce qui a été recetté :
- l'empreinte SHA-256 de la maquette et des référentiels ;
- le schéma IFC ;
- l'application auteur, l'auteur et l'organisation, et la date d'export, tels que les donne l'en-tête IFC.

Le contrôle porte sur le modèle ouvert dans Bonsai, alors que l'empreinte est celle du fichier enregistré. **Enregistrez la maquette avant la recette** pour que les deux coïncident.

### Mention de confidentialité

La mention saisie figure :
- en tête et en pied de chaque page du rapport Excel ;
- en bandeau sur la fiche ;
- en étiquette et en tête de description de chaque sujet BCF.

Elle ne remplace pas les règles de confidentialité propres au projet.

### Niveaux

Les anomalies sont regroupées par étage (`IfcBuildingStorey`) ou, pour les infrastructures en IFC 4.3, par partie d'ouvrage (`IfcFacilityPart` : quai, partie de pont…). Un élément rangé dans un local est rattaché à l'étage du local.

## IDS d'exemple

Dans `checkers_bonsai/examples/` (exemples à adapter, ce ne sont pas des référentiels officiels) :
- `exploitation_maintenance.ids` : pour les équipements techniques, repère (`Tag`), désignation, fabricant et modèle, numéro de série, fin de garantie et classification patrimoniale ; pour les locaux, numéro et désignation. Il couvre IFC4 / IFC4.3, avec une variante IFC2X3.
- `conception_minimale.ids` : étages nommés, locaux identifiés, murs et dalles (porteur / extérieur), portes (degré coupe-feu).

## Règles Excel

Le format est celui de l'application web : onglet `Element_Types` (colonne `IFC_Class`), un onglet par PSet (`IFC_Class`, `Parametre`, `Type` ∈ String, Int, Float, Number, Bool) et une colonne `Corps_Etat` facultative, avec l'onglet `Corps_Etat` (`Code`, `Nom`).

Différences avec l'application web :
- une règle sur une classe s'applique aussi à ses sous-classes, par exemple `IfcWall` à `IfcWallStandardCase` en IFC2X3 ;
- une valeur vide compte comme absente ;
- un type inconnu dans le fichier est signalé à l'ouverture, au lieu d'être compté en non-conformité sur chaque élément.

## Installation

```bash
python bonsai_extension/build.py      # -> bonsai_extension/dist/checkers_bonsai-<version>.zip
```

Dans Blender : **Édition → Préférences → Extensions → menu ⌄ → Installer depuis le disque…** Bonsai doit être installé et activé. Développé pour Blender 5.2 / Bonsai 0.9 (Python 3.13), avec Blender 4.2 minimum.

## Tests

```bash
python -m unittest discover tests
```

Les tests s'exécutent sans Blender ni IfcOpenShell, grâce aux doublures de `tests/fakes/`. Ils couvrent :
- les règles, les niveaux, le contrôle PSet et le contrôle IDS ;
- le verdict, le rapport (mention, formules neutralisées) et le BCF ;
- la garantie hors ligne et les opérateurs du panneau.

Pour valider aussi le BCF contre les schémas officiels, définissez `BCF_XSD_DIR` (voir l'en-tête du fichier de test). La validation finale se fait dans Bonsai, sur de vraies maquettes.

## Structure

```
checkers_bonsai/
├── blender_manifest.toml   # manifeste (Blender 4.2+), sans permission réseau
├── __init__.py             # enregistrement
├── bonsai_link.py          # accès au modèle ouvert et aux objets Bonsai
├── props.py / operators.py / ui.py
├── examples/               # IDS d'exemple
└── core/                   # moteur, sans dépendance à Blender
    ├── recette.py          # orchestration et verdict
    ├── ids.py              # contrôle IDS sécurisé (ifctester)
    ├── rules.py, psets.py  # règles Excel et contrôle des PSet
    ├── spatial.py          # niveau : étage ou partie d'ouvrage
    ├── trace.py            # empreintes, en-tête IFC
    ├── report.py           # fiche de recette Excel
    ├── bcf.py              # export BCF 2.1
    └── excel_safety.py     # neutralisation des formules injectées
```
