# Checkers – recette de maquettes IFC dans Bonsai

Extension Blender pour [Bonsai](https://bonsaibim.org/) qui contrôle la maquette IFC ouverte et produit les livrables de recette :

- contrôle des exigences **IDS** (buildingSMART) et/ou d'un fichier **Excel de PSet requis** ;
- **verdict** (validé, avec réserves, rejeté) selon des seuils réglables ;
- anomalies listées par niveau, classe IFC et exigence, **sélectionnables dans la vue 3D** ; coloration conforme / non conforme ;
- **fiche de recette Excel** à signer, avec traçabilité (empreintes SHA-256 de la maquette et des référentiels) ;
- **export BCF 2.1** des anomalies pour la maîtrise d'œuvre et les entreprises.

L'extension fonctionne **entièrement hors ligne** : aucun accès réseau, aucune dépendance à installer (Bonsai fournit ifcopenshell, ifctester et openpyxl).

## Installation

1. Construire l'archive : `python build.py`, qui produit `dist/checkers_bonsai-<version>.zip`.
2. Dans Blender : **Édition → Préférences → Extensions → menu ⌄ → Installer depuis le disque…**, puis choisir le `.zip`.
3. Vérifier que **Bonsai** et **Checkers - recette IFC** sont activés.

Blender 4.2 minimum ; développé pour Blender 5.2 / Bonsai 0.9.

## Utilisation

Ouvrir un projet IFC dans Bonsai, puis, dans la vue 3D, ouvrir la barre latérale (`N`) et l'onglet **Checkers**.

1. **Référentiels**
   - **IDS** : les exigences du cahier des charges BIM.
   - **Règles Excel** : facultatif. Si le fichier déclare des corps d'état, choisir celui du lot.
   - Le bouton à droite de chaque champ charge un exemple fourni avec l'extension.
2. **Recette**
   - Renseigner le lot, la phase (suggestions ESQ à DOE, exploitation), la personne qui recette et une mention de confidentialité facultative.
   - Ajuster si besoin les seuils de verdict : 98 % et 90 % par défaut.
   - Cliquer sur **Lancer la recette**.
3. **Résultats**
   - Verdict, taux de conformité et avertissements.
   - Liste des anomalies : le bouton de chaque ligne sélectionne et cadre les éléments concernés.
   - **Colorer** : vert = conforme, rouge = non conforme. **Réinitialiser** restaure les couleurs d'origine.
4. **Livrables**
   - **Fiche de recette (Excel)** : onglets *Fiche de recette* (à imprimer et signer), *IDS* et *Détail PSet* (détail élément par élément).
   - **Anomalies (BCF)** : un sujet par anomalie, avec les éléments sélectionnés. Le fichier s'ouvre dans BIMcollab, Solibri, Revit/Navisworks (plugin) ou dans Bonsai.

Enregistrer la maquette avant la recette : l'empreinte notée sur la fiche est celle du fichier sur disque. Les résultats restent en mémoire jusqu'à la fermeture de Blender ; ils ne sont pas enregistrés dans le `.blend`.

## Référentiels

### IDS

Les fichiers IDS sont contrôlés avant lecture :
- taille limitée et encodage UTF-8 obligatoire ;
- DOCTYPE et ENTITY refusés ;
- motifs regex dangereux refusés ;
- validation par le schéma XSD officiel.

Les spécifications qui visent une autre version IFC que celle de la maquette sont ignorées.

Exemples fournis dans `checkers_bonsai/examples/`, à adapter au projet :
- `exploitation_maintenance.ids` : équipements (repère, désignation, fabricant, modèle, n° de série, fin de garantie, classification) et locaux (numéro, désignation) ;
- `conception_minimale.ids` : étages nommés, locaux identifiés, murs et dalles porteurs, portes coupe-feu.

### Règles Excel

Exemple : `checkers_bonsai/examples/regles_exemple.xlsx`.

- Onglet `Element_Types` : colonne `IFC_Class`, les classes contrôlées, sous-classes comprises.
- Un onglet par PSet : colonnes `IFC_Class`, `Parametre`, `Type` (String, Int, Float, Number ou Bool).
- Colonne facultative `Corps_Etat` : vide = tous les corps d'état ; `ARC, STR` = règle partagée. L'onglet `Corps_Etat` (`Code`, `Nom`) déclare les choix proposés dans le panneau.

Un paramètre absent, vide ou de type incorrect est une non-conformité.

### Verdict

Chaque contrôle donne *validé* au-dessus du premier seuil, *validé avec réserves* au-dessus du second, *rejeté* en dessous. Une spécification IDS obligatoire sans aucun élément concerné entraîne le rejet. Le verdict global est le plus sévère.

Les anomalies sont regroupées par étage ou, en IFC 4.3, par partie d'ouvrage (`IfcFacilityPart`). Un élément placé dans un local est rattaché à l'étage du local.

## Développement

```bash
python -m unittest discover tests     # sans Blender ni IfcOpenShell (doublures dans tests/fakes/)
python build.py                       # archive de l'extension
```

Pour valider aussi le BCF contre les schémas buildingSMART, définir `BCF_XSD_DIR` (voir `tests/test_bonsai_extension.py`).

```
checkers_bonsai/
├── blender_manifest.toml   # manifeste de l'extension (sans permission réseau)
├── __init__.py             # enregistrement
├── bonsai_link.py          # accès au modèle et aux objets Bonsai
├── props.py, operators.py, ui.py
├── examples/               # IDS et règles Excel d'exemple
└── core/                   # moteur, indépendant de Blender
    ├── recette.py          # orchestration et verdict
    ├── ids.py              # contrôle IDS (ifctester)
    ├── rules.py, psets.py  # règles Excel et contrôle des PSet
    ├── spatial.py          # niveau : étage ou partie d'ouvrage
    ├── trace.py            # empreintes et en-tête IFC
    ├── report.py           # fiche de recette Excel
    ├── bcf.py              # export BCF 2.1
    └── excel_safety.py     # neutralisation des formules
```

## Licence et contact

Licence MIT.

ZKHCHICHE – khchiche.zakaria@gmail.com
ABOUZNIR – ali.bouznir@gmail.com
