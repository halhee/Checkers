# Historique des versions

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
