# Historique des versions

## 0.5.0

Nouveautés :
- Relevé terrain : fiche de relevé à remplir sur site, puis contrôle des écarts avec la maquette (équipement absent, repère ou n° de série différent, équipement non modélisé). Les écarts sont des anomalies comme les autres : sélection 3D, fiche, BCF, suivi.
- Tableau de suivi : consolidation des journaux de recette d'un dossier, avec l'état de la dernière recette de chaque lot et la tendance.
- Vérification de l'installation depuis le panneau : bibliothèques attendues, IDS d'exemple, modèle ouvert.
- Tests d'intégration avec les vraies bibliothèques IfcOpenShell et ifctester, lancés par l'intégration continue sous Python 3.11 et 3.13.

Corrections :
- IDS : une spécification obligatoire qu'aucun élément ne concerne entraînait un rejet sans aucune anomalie visible. Elle apparaît maintenant dans la liste, la fiche et le BCF.
- IDS : l'identifiant des spécifications (ex. « EXP-01 »), qu'ifctester ne relit pas, est repris du fichier.
- Suivi : les anomalies d'un contrôle non relancé (ex. recette précédente avec IDS, la suivante sans) étaient comptées comme levées. Elles sont comptées à part, « non recontrôlées ».
- Structure : deux bâtiments ayant chacun un « RDC » n'est plus un doublon d'étages ; un GlobalId en double sur de simples jeux de propriétés est un constat majeur et non bloquant ; un contrôle qui n'a pas pu s'exécuter interdit de valider sans réserve.
- Règles Excel : colonne `IFC_Class` absente signalée par une erreur claire ; règles portant sur une classe non déclarée signalées au lieu d'être ignorées en silence.
- Livrables : un caractère de contrôle dans un nom d'élément empêchait d'enregistrer la fiche et produisait un BCF illisible ; un « & » dans la mention de confidentialité était lu comme un code de mise en forme dans les en-têtes Excel.
- Panneau : la coloration passe la vue en mode solide, seul mode où elle se voit ; un objet hors du calque de vue ne fait plus échouer la sélection ; toute erreur d'enregistrement est rapportée sans trace Python.
- Journal de recette : structure vérifiée plus strictement à la relecture ; son empreinte figure dans la traçabilité de la fiche suivante.

## 0.4.0

Nouveautés :
- Contrôles de structure du modèle (S01 à S11), sans référentiel, avec gravité bloquante, majeure ou mineure.
- Suivi entre livraisons : journal `.json` enregistré avec la fiche, comparaison avec la recette précédente (anomalies levées, persistantes, nouvelles), section et onglet « Évolution » dans la fiche.
- Inventaire d'exploitation : équipements et locaux, avec le taux de renseignement de chaque information.
- BCF : identifiants de sujets stables d'une recette à l'autre, sujets clos pour les anomalies levées, priorité selon la gravité, phase reportée dans le champ *Stage*.
- Coloration en vert des éléments conformes à l'IDS, et plus seulement aux règles Excel.

Corrections et durcissement :
- Manifeste : justification de la permission « files » ramenée sous la limite de 64 caractères imposée par Blender. Les archives précédentes pouvaient être refusées à l'installation.
- Types des règles Excel vérifiés strictement. Auparavant, `Bool` acceptait n'importe quelle valeur et un nombre saisi comme texte passait pour un nombre.
- Classeur de règles : archive contrôlée avant ouverture (bombe de décompression).
- Les identifiants des éléments ne sont plus écrits dans le fichier `.blend`.
- Résultats invalidés quand un autre modèle est ouvert ; résumé effacé à l'ouverture d'un fichier.
- Avertissement quand le modèle contient des modifications non enregistrées.
- Seuils de verdict incohérents refusés.
- Règle Excel la plus précise appliquée quand deux classes déclarées couvrent le même élément.
- Message clair si une bibliothèque fournie par Bonsai est introuvable.

## 0.3.0

- Dépôt réduit à l'extension Bonsai ; fichier de règles d'exemple accessible depuis le panneau.

## 0.2.x

- Extension autonome : contrôle IDS et règles Excel, verdict, fiche de recette avec traçabilité SHA-256, export BCF 2.1, niveaux IFC 4.3.
