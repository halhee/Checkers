# Checkers – recette de maquettes IFC dans Bonsai

Extension Blender pour [Bonsai](https://bonsaibim.org/) qui fait la recette de la maquette IFC ouverte, de la réception à l'exploitation :

- **contrôles** : structure du modèle, exigences **IDS** (buildingSMART), **PSet requis** (fichier Excel), **relevé terrain** ;
- **verdict** (validé, avec réserves, rejeté) selon des seuils réglables ;
- anomalies **sélectionnables dans la vue 3D**, coloration conforme / non conforme ;
- **suivi entre livraisons** : anomalies levées, persistantes et nouvelles depuis la recette précédente ;
- **fiche de recette Excel** à signer, avec traçabilité (empreintes SHA-256) ;
- **export BCF 2.1** des anomalies pour la maîtrise d'œuvre et les entreprises ;
- **inventaire d'exploitation** et **fiche de relevé terrain** : équipements et locaux de la maquette ;
- **tableau de suivi** : état de tous les lots à partir des journaux de recette.

L'extension fonctionne **entièrement hors ligne** : aucun accès réseau, aucune dépendance à installer (Bonsai fournit ifcopenshell, ifctester et openpyxl).

## Installation

1. Construire l'archive : `python build.py`, qui produit `dist/checkers_bonsai-<version>.zip`.
2. Dans Blender : **Édition → Préférences → Extensions → menu ⌄ → Installer depuis le disque…**, puis choisir le `.zip`.
3. Vérifier que **Bonsai** et **Checkers - recette IFC** sont activés.
4. Dans le panneau Checkers, section **Installation**, cliquer sur **Vérifier l'installation** : chaque bibliothèque attendue et les IDS d'exemple sont contrôlés sur le poste.

Blender 4.2 minimum ; développé pour Blender 5.2 / Bonsai 0.9.

## Utilisation

Dans la vue 3D, ouvrir la barre latérale (`N`) et l'onglet **Checkers**. Les sections « Suivi de projet » et « Installation » sont disponibles sans maquette ; les autres demandent un projet IFC ouvert dans Bonsai.

1. **Contrôles**
   - **Contrôler la structure du modèle** : coché par défaut, sans référentiel.
   - **IDS** : les exigences du cahier des charges BIM.
   - **Règles Excel** : facultatif. Si le fichier déclare des corps d'état, choisir celui du lot.
   - **Relevé terrain** : facultatif, la fiche de relevé remplie sur site (voir plus bas).
   - Le bouton à droite des champs IDS et Règles Excel charge un exemple fourni avec l'extension.
2. **Recette**
   - Renseigner le lot, la phase (suggestions ESQ à DOE, exploitation), la personne qui recette et une mention de confidentialité facultative.
   - **Recette précédente** : le journal `.json` de la livraison précédente, pour suivre l'évolution.
   - Ajuster si besoin les seuils de verdict : 98 % et 90 % par défaut.
   - Cliquer sur **Lancer la recette**.
3. **Résultats**
   - Verdict, résultat de chaque contrôle, évolution et avertissements.
   - Liste des anomalies : le bouton de chaque ligne sélectionne et cadre les éléments concernés.
   - **Colorer** : rouge = au moins une anomalie, vert = contrôlé sans anomalie. **Réinitialiser** restaure l'affichage d'origine.
4. **Livrables de recette**
   - **Fiche de recette (Excel) + journal** : la fiche à imprimer et signer, les onglets de détail, et le journal `.json` de même nom, à conserver pour la livraison suivante.
   - **Anomalies (BCF)** : un sujet par anomalie, avec les éléments sélectionnés. Le fichier s'ouvre dans BIMcollab, Solibri, Revit/Navisworks (plugin) ou dans Bonsai.
5. **Exploitation et terrain**
   - **Inventaire d'exploitation (Excel)** et **Fiche de relevé terrain (Excel)** : ne demandent pas de recette, seulement une maquette ouverte.
6. **Suivi de projet**
   - **Tableau de suivi (Excel)** : consolide les journaux de recette d'un dossier.

Enregistrer la maquette avant la recette : l'empreinte notée sur la fiche est celle du fichier sur disque. Si Bonsai signale des modifications non enregistrées, la fiche le mentionne.

## Contrôles

### Structure du modèle

| Code | Contrôle | Gravité |
|---|---|---|
| S01 | Projet unique | Bloquant |
| S02 | Identifiants (GlobalId) uniques | Bloquant pour des objets ou des types, majeur pour des jeux de propriétés ou des relations seuls |
| S03 | Unité de longueur déclarée | Majeur |
| S04 | Site, bâtiment ou ouvrage présent | Majeur |
| S05 | Bâtiments découpés en étages | Majeur |
| S06 | Éléments rattachés à la structure spatiale | Majeur |
| S07 | Objets génériques (`IfcBuildingElementProxy`) | Mineur, majeur au-delà de 10 % des éléments |
| S08 | Étages nommés, sans doublon dans un bâtiment | Mineur |
| S09 | Éléments nommés | Mineur |
| S10 | Éléments avec géométrie | Mineur |
| S11 | Géoréférencement renseigné | Mineur |

Un constat bloquant entraîne le rejet ; un constat majeur, ou un contrôle qui n'a pas pu s'exécuter, des réserves ; un constat mineur est signalé sans effet sur le verdict. Ces gravités sont des choix par défaut, définis en tête de `core/structure.py`.

### IDS

Les fichiers IDS sont contrôlés avant lecture :
- taille limitée et encodage UTF-8 obligatoire ;
- DOCTYPE et ENTITY refusés ;
- motifs regex dangereux refusés ;
- validation par le schéma XSD officiel.

Les spécifications qui visent une autre version IFC que celle de la maquette sont ignorées. Une spécification obligatoire qu'aucun élément ne concerne apparaît comme une anomalie portant sur la maquette entière. Le libellé des exigences est celui d'ifctester, en anglais.

Exemples fournis dans `checkers_bonsai/examples/`, à adapter au projet :
- `exploitation_maintenance.ids` : équipements (repère, désignation, fabricant, modèle, n° de série, fin de garantie, classification) et locaux (numéro, désignation) ;
- `conception_minimale.ids` : étages nommés, locaux identifiés, murs et dalles porteurs, portes coupe-feu.

### Règles Excel

Exemple : `checkers_bonsai/examples/regles_exemple.xlsx`.

- Onglet `Element_Types` : colonne `IFC_Class`, les classes contrôlées, sous-classes comprises.
- Un onglet par PSet : colonnes `IFC_Class`, `Parametre`, `Type` (String, Int, Float, Number ou Bool).
- Colonne facultative `Corps_Etat` : vide = tous les corps d'état ; `ARC, STR` = règle partagée. L'onglet `Corps_Etat` (`Code`, `Nom`) déclare les choix proposés dans le panneau.

Un paramètre absent, vide ou de type incorrect est une non-conformité. Le type est vérifié strictement : un nombre saisi comme texte (« 2.5 ») n'est pas un nombre, un texte n'est pas un booléen. Une règle qui porte sur une classe absente de `Element_Types` ne serait jamais appliquée : elle est signalée par un avertissement.

### Relevé terrain

Pour confronter la maquette à ce qui existe réellement sur site :

1. **Fiche de relevé terrain (Excel)** produit la liste des équipements de la maquette, par niveau et par local, avec des colonnes à remplir sur site : *Présent sur site*, *Repère relevé*, *N° de série relevé*, *Observation*. La fiche s'imprime ou se remplit sur tablette.
2. Sur site, un équipement trouvé mais absent de la liste s'ajoute en bas, sans GlobalId.
3. De retour, la fiche remplie se choisit dans **Contrôles → Relevé terrain**, puis on lance la recette.

Les écarts deviennent des anomalies de source « Terrain » : équipement absent sur site, repère différent, n° de série différent, équipement non modélisé. Elles se sélectionnent dans la vue 3D, figurent sur la fiche de recette et dans le BCF, et sont suivies d'une livraison à l'autre.

La comparaison se fait avec les valeurs actuelles de la maquette. Un équipement laissé vide dans la fiche est compté « non vérifié », pas « conforme ».

### Verdict

L'IDS, les règles Excel et le relevé terrain donnent chacun *validé* au-dessus du premier seuil, *validé avec réserves* au-dessus du second, *rejeté* en dessous. Une spécification IDS obligatoire sans aucun élément concerné entraîne le rejet. Le verdict global est le plus sévère des contrôles réalisés.

Les anomalies sont regroupées par étage ou, en IFC 4.3, par partie d'ouvrage (`IfcFacilityPart`). Un élément placé dans un local est rattaché à l'étage du local.

## Suivi

### Entre deux livraisons

Chaque fiche de recette est enregistrée avec un journal `.json`. À la livraison suivante, ce journal fourni comme **recette précédente** permet de comparer, élément par élément et exigence par exigence :

- les anomalies **levées**, en distinguant celles qui disparaissent parce que l'élément a été supprimé ;
- les anomalies **persistantes** ;
- les anomalies **nouvelles** ;
- les anomalies **non recontrôlées**, quand un contrôle de la recette précédente n'a pas été relancé : elles ne sont pas comptées comme levées.

La fiche reçoit une section et un onglet « Évolution », et l'empreinte du journal précédent figure dans la traçabilité. Un avertissement signale un référentiel, un lot ou un projet différent de la recette précédente, car la comparaison n'aurait alors pas le même sens.

Dans le BCF, une anomalie garde le même identifiant de sujet d'une recette à l'autre : l'outil du destinataire met le sujet à jour au lieu de le dupliquer, et les anomalies entièrement levées lui arrivent comme sujets clos.

### Sur tout le projet

**Tableau de suivi (Excel)** lit tous les journaux d'un dossier, sous-dossiers compris, sans ouvrir aucune maquette :

- **Synthèse** : une ligne par lot, avec la date, la phase et le verdict de sa dernière recette, le nombre d'anomalies ouvertes et l'écart avec la recette précédente ;
- **Recettes** : l'historique complet, avec le verdict et le taux de chaque contrôle.

Il suffit de ranger les journaux de recette dans un même dossier, par exemple un sous-dossier par lot.

## Inventaire d'exploitation

Classeur Excel de trois onglets, à reprendre dans un outil de maintenance ou de gestion de patrimoine :

- **Synthèse** : taux de renseignement de chaque information, qui mesure si la maquette est exploitable ;
- **Équipements** : repère, désignation, classe et type, niveau, local, système, fabricant, modèle, n° de série, fin de garantie, classification, GlobalId. Les équipements sont les éléments de distribution technique, hors tronçons et raccords de réseaux, et les appareils de transport ;
- **Locaux** : numéro, désignation, niveau, surface, classification, GlobalId.

Pour un export normalisé COBie, Bonsai propose le module *Facility Management* (IfcFM).

## Sécurité et confidentialité

- **Aucun accès réseau** : l'extension n'importe aucune bibliothèque réseau, ne lance aucun programme externe et ne demande pas la permission « network » à Blender. Un test le vérifie.
- **Fichiers reçus traités comme non fiables** : IDS, classeurs (règles, relevé terrain ; y compris contre les bombes de décompression) et journaux de recette sont contrôlés avant lecture.
- **Classeurs et BCF produits** : les textes venus de la maquette ne peuvent pas devenir des formules Excel, et les caractères de contrôle qui rendraient un fichier illisible sont retirés.
- **Fichier `.blend`** : il garde les réglages du panneau et le résumé de la dernière recette, mais aucun identifiant d'élément. Le résumé est effacé à l'ouverture d'un fichier.
- **Mention de confidentialité** : portée sur chaque page des classeurs et dans chaque sujet BCF. Elle ne remplace pas les règles de confidentialité du projet.

## Limites connues

- L'extension n'a pas de test automatique dans Blender : les tests couvrent le moteur, avec des doublures puis avec les vraies bibliothèques IfcOpenShell et ifctester, mais le panneau lui-même doit être validé dans Bonsai à chaque version.
- La fiche de recette et les journaux sont des fichiers modifiables : la valeur probante vient du visa et des empreintes SHA-256, pas d'une signature électronique.
- L'empreinte est celle du fichier enregistré, alors que le contrôle porte sur le modèle ouvert.
- Les points de vue BCF n'ont pas de caméra : l'outil de destination cadre la sélection. L'extension écrit des BCF, elle ne relit pas les réponses.
- Le contrôle anti-ReDoS des motifs regex d'un IDS est une heuristique : n'ouvrir que des IDS de provenance connue.
- La recette s'exécute dans le fil principal de Blender, dont l'interface reste figée pendant le calcul.
- Un seul modèle par recette : pas de contrôle de cohérence entre maquettes, pas de détection de conflits entre lots (voir IfcClash dans Bonsai), pas de modèles liés.
- L'inventaire ne retient que la distribution technique et les appareils de transport : un équipement modélisé en objet générique n'y figure pas.

## Développement

```bash
python -m unittest discover tests               # logique de Checkers, sans Blender ni IfcOpenShell
python -m unittest discover tests_integration   # avec les vraies bibliothèques (pip install ifcopenshell ifctester openpyxl)
python build.py                                 # archive de l'extension
```

- `tests/` utilise des doublures (`tests/fakes/`) : rapide, sans dépendance, mais ne prouve pas l'accord avec les bibliothèques réelles.
- `tests_integration/` construit une maquette avec IfcOpenShell et déroule une recette complète avec ifctester. Ces tests sont sautés si les bibliothèques ne sont pas installées.
- L'intégration continue (`.github/workflows/tests.yml`) lance les deux, sous Python 3.11 et 3.13, et valide les BCF contre les schémas buildingSMART.

Pour valider le BCF en local, définir `BCF_XSD_DIR` (voir `tests/support.py`).

```
checkers_bonsai/
├── blender_manifest.toml   # manifeste de l'extension (sans permission réseau)
├── __init__.py             # enregistrement
├── bonsai_link.py          # accès au modèle et aux objets Bonsai
├── props.py, operators.py, ui.py
├── examples/               # IDS et règles Excel d'exemple
└── core/                   # moteur, indépendant de Blender
    ├── recette.py          # orchestration et verdict
    ├── structure.py        # contrôles de structure
    ├── ids.py              # contrôle IDS (ifctester)
    ├── rules.py, psets.py  # règles Excel et contrôle des PSet
    ├── survey.py           # relevé terrain
    ├── journal.py          # journal de recette et comparaison entre livraisons
    ├── tracking.py         # tableau de suivi
    ├── inventory.py        # inventaire d'exploitation
    ├── spatial.py          # niveau : étage ou partie d'ouvrage
    ├── trace.py            # empreintes et en-tête IFC
    ├── report.py           # fiche de recette Excel
    ├── bcf.py              # export BCF 2.1
    ├── diagnostic.py       # vérification de l'installation
    └── excel_safety.py     # neutralisation des formules et des caractères de contrôle
```

Historique des versions : [CHANGELOG.md](CHANGELOG.md).

## Licence et contact

Licence MIT.

ZKHCHICHE – khchiche.zakaria@gmail.com
ABOUZNIR – ali.bouznir@gmail.com
