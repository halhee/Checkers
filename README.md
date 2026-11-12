# Checkers – recette de maquettes IFC dans Bonsai

Extension Blender pour [Bonsai](https://bonsaibim.org/) qui fait la recette de la maquette IFC ouverte, de la réception à l'exploitation :

- **contrôles** : structure du modèle, exigences **IDS** (buildingSMART), **PSet requis** (fichier Excel), **relevé terrain** ;
- **coordination entre lots** : cohérence de référence avec les maquettes des autres lots (origine, étages, géoréférencement…) et **conflits géométriques** avec IfcClash, selon des règles réglables ;
- **verdict** (validé, avec réserves, rejeté) selon des seuils réglables, et **dérogations** motivées pour les anomalies admises ;
- anomalies **sélectionnables dans la vue 3D**, coloration conforme / non conforme ;
- **suivi entre livraisons** : anomalies levées, persistantes et nouvelles depuis la recette précédente ;
- **fiche de recette Excel** à signer, avec traçabilité (empreintes SHA-256) ;
- **export BCF 2.1** des anomalies pour la maîtrise d'œuvre et les entreprises ;
- **inventaire d'exploitation** et **fiche de relevé terrain** : équipements et locaux de la maquette ;
- **tableau de suivi** : état de tous les lots à partir des journaux de recette.

L'extension fonctionne **entièrement hors ligne** : aucun accès réseau, aucune dépendance à installer (Bonsai fournit ifcopenshell, ifctester, IfcClash et openpyxl).

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
   - **Coordination entre lots** : facultatif, les maquettes des autres lots et les règles de conflits (voir plus bas).
   - Le bouton à droite des champs IDS et Règles Excel charge un exemple fourni avec l'extension.
2. **Recette**
   - Renseigner le lot, la phase (suggestions ESQ à DOE, exploitation), la personne qui recette et une mention de confidentialité facultative.
   - **Recette précédente** : le journal `.json` de la livraison précédente, pour suivre l'évolution.
   - **Dérogations** : facultatif, le fichier des anomalies admises (voir plus bas). Créé au premier « Admettre » s'il n'est pas choisi.
   - Ajuster si besoin les seuils de verdict : 98 % et 90 % par défaut.
   - Cliquer sur **Lancer la recette**.
3. **Résultats**
   - Verdict, résultat de chaque contrôle, évolution et avertissements.
   - Liste des anomalies : le premier bouton de chaque ligne sélectionne et cadre les éléments concernés, le second admet l'anomalie par dérogation.
   - Pour un conflit : la vue se centre sur le point de conflit et l'élément de l'autre lot s'affiche en filaire ; les flèches passent d'un conflit à l'autre.
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

### Coordination entre lots

Dans la section **Coordination entre lots**, ajouter avec **+** les maquettes IFC des autres lots (structure, CVC, plomberie…), chacune avec un nom court. Ces fichiers sont seulement lus.

**Cohérence de référence** (case cochée par défaut) : la maquette ouverte est comparée à chaque lot, sans géométrie.

| Code | Contrôle | Gravité |
|---|---|---|
| C01 | Même schéma IFC | Majeur |
| C02 | Même projet (GlobalId de l'`IfcProject`) | Majeur ; mineur si seul le GlobalId diffère |
| C03 | Même unité de longueur | Mineur |
| C04 | Même point d'origine et même orientation du bâtiment (1 cm, 0,01°) | Majeur |
| C05 | Étages de même nom à la même altitude (1 cm) | Majeur, par étage |
| C06 | Étages de même nom des deux côtés | Mineur, par étage |
| C07 | Même géoréférencement (`IfcMapConversion`, ou en IFC2X3 les jeux de propriétés `ePSet_MapConversion` et `ePSet_ProjectedCRS` du projet) | Majeur |
| C08 | Même nord géographique | Majeur |
| C09 | Aucun élément de même GlobalId dans les deux lots | Bloquant, par élément |

Un C04 ou un C05 en échec signifie que les maquettes ne se superposent pas : la fédération est fausse et les conflits géométriques entre ces lots n'ont pas de sens. La recette l'indique alors dans ses avertissements. Les entités spatiales (site, bâtiment, étages) peuvent être partagées ; un élément physique ne doit appartenir qu'à un lot.

**Conflits géométriques** (case à cocher, avec IfcClash fourni par Bonsai) : chaque règle confronte une sélection de la maquette ouverte (A) à une sélection d'un lot ou de la maquette ouverte elle-même (B) :

- **Sélections** : requêtes IfcOpenShell, par exemple `IfcDuctSegment, IfcPipeSegment` ou `IfcBeam, IfcColumn, IfcSlab`. Vide : tous les éléments physiques, hors réservations.
- **Mode** : *intersection* (pénétration au-delà de la tolérance), *collision* (contact compris), *dégagement* (distance minimale exigée, par exemple pour l'accès maintenance).
- **Distance** et **gravité** propres à chaque règle.
- **Une règle par lot** crée une règle d'intersection à 1 cm avec chaque lot.

Chaque conflit devient une anomalie de source « Conflit », regroupée par étage, classe et règle. La sélection 3D porte sur l'élément de la maquette ouverte ; dans le BCF, l'élément de l'autre lot est aussi sélectionné (sans couleur) et une caméra est placée sur le point de conflit. L'onglet « Conflits » de la fiche liste chaque paire avec sa distance et ses coordonnées.

Gravité : un conflit bloquant rejette la maquette, un conflit majeur ou une règle qui n'a pas pu s'exécuter (sélection vide, géométrie non calculable) donne des réserves, un conflit mineur est seulement signalé.

**Voir un conflit dans la vue 3D** : le bouton de sélection d'une anomalie « Conflit » sélectionne les éléments de la maquette ouverte, place le curseur 3D et un repère sur le point de conflit, centre la vue dessus et affiche en filaire, par-dessus la maquette, les éléments de l'autre lot (calculés à partir de son fichier IFC). Les flèches **◀ ▶** passent d'un conflit à l'autre, un par un. Ces objets d'affichage sont temporaires : **Effacer** les retire, et ils ne sont jamais enregistrés dans le fichier `.blend`. Si Bonsai a décalé la scène (maquettes aux grandes coordonnées), le point et les éléments suivent ce décalage.

Le suivi d'un conflit d'une recette à l'autre (journal, sujets BCF, dérogations) ne dépend pas de la distance de la règle : ajuster une tolérance ne fait pas passer les conflits pour nouveaux. Changer de mode, si.

### Dérogations

Certaines anomalies sont voulues ou acceptées : une gaine dans sa réservation, un écart validé par la maîtrise d'ouvrage, un ouvrage existant conservé. Plutôt que de les retrouver à chaque recette, on les **admet** :

1. Dans la liste des anomalies, bouton **Admettre** de la ligne.
2. Saisir le **motif** (obligatoire), choisir la portée : tous les éléments de l'anomalie, ou seulement ceux sélectionnés dans la vue 3D ; cocher *Valable pour tous les lots* si besoin.
3. Le verdict, les taux et la coloration sont recalculés aussitôt, sans relancer les contrôles.

Une dérogation porte sur des éléments précis (GlobalId) pour une exigence donnée : un nouvel élément présentant le même défaut reste une anomalie. Elle est enregistrée dans le **fichier des dérogations** (`.json`, par défaut `<maquette>.derogations.json` à côté de la maquette), avec le motif, la personne qui recette et la date, et vaut pour les recettes suivantes. **Retirer** (liste « Admises par dérogation ») rouvre l'anomalie.

Effets :
- les anomalies admises sortent du verdict : un élément dont toutes les anomalies d'un contrôle sont admises compte comme conforme dans le taux de ce contrôle ; pour la structure, la cohérence et les conflits, le verdict est recalculé sans elles ;
- la fiche les liste dans l'onglet **Dérogations**, avec le motif, ainsi que les dérogations **sans objet** (défaut corrigé, élément supprimé) qui peuvent être retirées du fichier ; l'empreinte du fichier figure dans la traçabilité ;
- le BCF les transmet comme **sujets clos**, avec le motif en commentaire ;
- le journal les garde à part : elles ne sont comptées ni comme levées ni comme nouvelles ; le tableau de suivi a une colonne « Anomalies admises ».

### Verdict

L'IDS, les règles Excel et le relevé terrain donnent chacun *validé* au-dessus du premier seuil, *validé avec réserves* au-dessus du second, *rejeté* en dessous. Une spécification IDS obligatoire sans aucun élément concerné entraîne le rejet. La cohérence entre lots et les conflits suivent leurs gravités (comme la structure). Le verdict global est le plus sévère des contrôles réalisés. Les anomalies admises par dérogation n'y entrent pas.

Les anomalies sont regroupées par étage ou, en IFC 4.3, par partie d'ouvrage (`IfcFacilityPart`). Un élément placé dans un local est rattaché à l'étage du local.

## Suivi

### Entre deux livraisons

Chaque fiche de recette est enregistrée avec un journal `.json`. À la livraison suivante, ce journal fourni comme **recette précédente** permet de comparer, élément par élément et exigence par exigence :

- les anomalies **levées**, en distinguant celles qui disparaissent parce que l'élément a été supprimé ;
- les anomalies **persistantes** ;
- les anomalies **nouvelles** ;
- les anomalies **non recontrôlées**, quand un contrôle de la recette précédente n'a pas été relancé : elles ne sont pas comptées comme levées ;
- les anomalies **admises** depuis, par dérogation.

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

- L'extension n'a pas de test automatique dans Blender : les tests couvrent le moteur, avec des doublures puis avec les vraies bibliothèques IfcOpenShell, ifctester et IfcClash, mais le panneau lui-même doit être validé dans Bonsai à chaque version.
- La fiche de recette et les journaux sont des fichiers modifiables : la valeur probante vient du visa et des empreintes SHA-256, pas d'une signature électronique.
- L'empreinte est celle du fichier enregistré, alors que le contrôle porte sur le modèle ouvert.
- Seuls les points de vue BCF des conflits ont une caméra ; ailleurs l'outil de destination cadre la sélection. L'extension écrit des BCF, elle ne relit pas les réponses.
- Le contrôle anti-ReDoS des motifs regex d'un IDS est une heuristique : n'ouvrir que des IDS de provenance connue.
- La recette s'exécute dans le fil principal de Blender, dont l'interface reste figée pendant le calcul.
- Dérogations : le fichier est un simple JSON ; le motif et l'auteur sont déclaratifs. Une dérogation sur un constat bloquant est possible : la fiche la rend visible, c'est au visa de la recette d'en juger.
- Vue 3D d'un conflit : au plus 20 éléments de l'autre lot affichés à la fois, et au plus 500 conflits parcourables par anomalie.
- Coordination : la recette porte sur la maquette ouverte, comparée aux autres lots ; les conflits entre deux lots tiers ne sont pas cherchés. Checkers signale les écarts, il ne recale pas les maquettes.
- Conflits : le calcul géométrique peut prendre plusieurs minutes sur de grosses maquettes, l'interface restant figée ; restreindre les sélections. Les résultats bruts d'IfcClash comportent des conflits admis (gaine dans sa réservation, élément posé sur un autre…) : régler mode, tolérance et sélections plutôt que tout tester contre tout. Au-delà de 50 000 conflits, le détail est tronqué.
- Mode *intersection* : lors des essais, une gaine traversant une poutre de part en part, affleurant sa face supérieure, n'a pas été détectée. En cas de doute, ajouter une règle en mode *collision* sur les mêmes sélections.
- L'inventaire ne retient que la distribution technique et les appareils de transport : un équipement modélisé en objet générique n'y figure pas.

## Développement

```bash
python -m unittest discover tests               # logique de Checkers, sans Blender ni IfcOpenShell
python -m unittest discover tests_integration   # avec les vraies bibliothèques (pip install ifcopenshell ifctester ifcclash openpyxl)
python build.py                                 # archive de l'extension
```

- `tests/` utilise des doublures (`tests/fakes/`) : rapide, sans dépendance, mais ne prouve pas l'accord avec les bibliothèques réelles.
- `tests_integration/` construit une maquette avec IfcOpenShell et déroule une recette complète avec ifctester, puis confronte deux lots avec IfcClash. Ces tests sont sautés si les bibliothèques ne sont pas installées.
- L'intégration continue (`.github/workflows/tests.yml`) lance les deux, sous Python 3.11 et 3.13, et valide les BCF contre les schémas buildingSMART.

Pour valider le BCF en local, définir `BCF_XSD_DIR` (voir `tests/support.py`).

```
checkers_bonsai/
├── blender_manifest.toml   # manifeste de l'extension (sans permission réseau)
├── __init__.py             # enregistrement
├── bonsai_link.py          # accès au modèle et aux objets Bonsai
├── viewer.py               # affichage d'un conflit dans la vue 3D
├── props.py, operators.py, ui.py
├── examples/               # IDS et règles Excel d'exemple
└── core/                   # moteur, indépendant de Blender
    ├── recette.py          # orchestration et verdict
    ├── structure.py        # contrôles de structure
    ├── ids.py              # contrôle IDS (ifctester)
    ├── rules.py, psets.py  # règles Excel et contrôle des PSet
    ├── federation.py       # cohérence de référence entre lots
    ├── clashes.py          # conflits géométriques (IfcClash)
    ├── survey.py           # relevé terrain
    ├── derogations.py      # fichier des dérogations
    ├── scene.py            # coordonnées IFC vers la scène, géométrie d'un élément
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
