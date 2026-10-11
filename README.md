# Checkers - Validation IFC

Checkers est une application web permettant de valider des fichiers IFC (Industry Foundation Classes) en vérifiant la présence et la conformité des Property Sets (PSet) et des paramètres requis.

<div align="center">
  <img src="static/Sanstitre.png" alt="Checkers Logo" width="600">
</div>

## Fonctionnalités

- Upload de plusieurs fichiers IFC (un par lot/discipline -- Architecture, Structure, CVC...), analysés ensemble et comparés
- Validation automatique des PSet et paramètres (fichier Excel) **et/ou** contrôle d'un fichier **IDS** (Information Delivery Specification, standard buildingSMART -- voir `ids_check.py`)
- Dashboard interactif avec statistiques en temps réel
- Graphique en donut pour la visualisation des résultats
- Bilan carbone basé sur les matériaux réels du modèle IFC (voir `carbon.py`), avec indicateur de fiabilité (mesuré / estimé / non calculable)
- Préconisations générées par 3 agents experts (carbone, qualité des données, métier/risques projet) + une synthèse croisée, en option (voir `agents/`)
- Visionneuse 3D du modèle IFC intégrée (onglet "Maquette 3D")
- Géolocalisation du site sur une carte (onglet "Géoréférencement", à partir de `IfcSite.RefLatitude`/`RefLongitude`)
- Génération de rapports détaillés au format Excel, avec une **Fiche de Recette** prête à signer (verdict par lot, anomalies ouvertes, zone de signature)
- **Suivi des anomalies et historique des analyses** (voir `history.py`) : chaque analyse est enregistrée, comparable à une analyse précédente pour voir ce qui a été corrigé/est nouveau/persiste
- Support des fichiers volumineux (jusqu'à 300 Mo)
- Interface utilisateur moderne et intuitive

## Analyse multilot

Un projet réel a un modèle par lot (discipline), pas un seul IFC. Checkers accepte plusieurs fichiers IFC en une seule analyse -- chacun nommé librement à l'upload (ex. "Architecture", "Structure", "CVC"). Chaque lot garde ses propres statistiques (visibles dans l'onglet Excel "Comparatif Lots" et le tableau du dashboard) tout en alimentant des totaux combinés (Résumé, bilan carbone global). **Pas de détection de clash spatial entre lots** (aucune intersection de géométrie vérifiée) -- uniquement des statistiques agrégées et comparées côte à côte ; les préconisations par agents IA restent volontairement lot-agnostiques (elles raisonnent sur les totaux combinés).

### Excel multi-corps d'état

Par défaut, tous les lots sont validés contre le **même** fichier Excel de paramètres requis. Pour donner des règles différentes par discipline dans **un seul** fichier Excel (voir `corps_etat.py`), on raisonne par PSet -- pas par corps d'état : mêmes onglets qu'un fichier à un seul jeu de règles (`Element_Types` + un onglet par PSet), avec juste une colonne **`Corps_Etat`** en plus sur chaque ligne. Une cellule vide = la ligne s'applique à tous les corps d'état ; plusieurs codes séparés par une virgule (`ARC, STR`) = une règle partagée entre eux, sans avoir à la dupliquer. Un onglet **`Corps_Etat`** (colonnes `Code`, `Nom`) déclare les disciplines disponibles, ex. `ARC | Architecture`, `STR | Structure` -- utilisé pour faire correspondre le nom du lot saisi à l'upload. Un exemple complet (`exemple_parametres_corps_etat.xlsx`) est disponible à la racine du projet.

À chaque analyse, le nom du lot saisi à l'upload est automatiquement associé à un corps d'état -- correspondance exacte, puis floue (`rapidfuzz`), puis un agent IA en dernier recours si c'est ambigu (choix strictement limité aux corps d'état déclarés, jamais d'invention). Si rien ne correspond avec confiance, le lot est quand même inclus dans le bilan carbone (sur tous ses éléments) mais **sans contrôle PSet**, avec un avertissement clair dans le rapport plutôt qu'un blocage. Sans onglet `Corps_Etat` dans le fichier, comportement inchangé (un seul jeu de règles partagé).

Dès que le fichier Excel choisi déclare un onglet `Corps_Etat`, un petit sélecteur apparaît à côté de chaque lot pour choisir explicitement son corps d'état (peuplé via la route `/corps_etat_registry`, appelée automatiquement au choix du fichier). Laissé sur "Auto" (par défaut), le pipeline de correspondance automatique ci-dessus s'applique comme d'habitude ; un choix explicite le court-circuite entièrement pour ce lot (visible dans le rapport comme "... (choix manuel)").

## Contrôle IDS (Information Delivery Specification)

L'IDS est le format ouvert buildingSMART pour exprimer ce qu'un modèle IFC doit contenir (classes, PSet, propriétés, attributs, valeurs). C'est l'équivalent normalisé de l'Excel de règles : on peut fournir **l'un, l'autre, ou les deux** (au moins un). Le contrôle s'appuie sur `ifctester`, le validateur de référence d'IfcOpenShell.

- Un seul fichier IDS, partagé par tous les lots (comme l'Excel). Un exemple (`exemple_ids.ids`) est fourni à la racine.
- **IDS seul** : les lots passent par le chemin « bilan carbone + IDS » (aucun contrôle PSet Excel) ; le verdict de recette vient de la conformité IDS. **Excel + IDS** : le verdict d'un lot est le plus sévère des deux.
- Les spécifications dont la version IFC (`ifcVersion`) ne vise pas le modèle sont **ignorées** (comportement normatif), pas comptées en échec.
- Les non-conformités IDS sont des anomalies comme les autres (regroupées par lot/étage/classe IFC/spécification, libellé `IDS : <nom>`) : elles alimentent l'historique, la comparaison avant/après, la Fiche de Recette et les agents IA.
- Résultats : section « Conformité IDS » du dashboard, colonne dans le tableau des lots, et onglet Excel **IDS** (synthèse par spécification + détail élément par élément, plafonné à 5 000 lignes par lot ; les compteurs restent exacts).

Le fichier IDS est traité comme **non fiable** (il vient souvent d'un tiers) : refus de toute déclaration `DOCTYPE`/`ENTITY` (XXE, expansion d'entités), UTF-8 obligatoire, 10 Mo maximum, validation XSD stricte, rejet des motifs regex à quantificateurs imbriqués ou trop longs (heuristique anti-ReDoS : elle écarte la forme classique du blocage mais ne prétend pas être exhaustive).

**Version** : `ifctester` doit rester alignée sur `ifcopenshell` (0.8.5 ↔ 0.8.5) ; la 0.9.x appelle des méthodes absentes d'IfcOpenShell 0.8.x.

## Recette fonctionnelle, anomalies et historique

Chaque analyse génère, en plus des onglets existants, un onglet Excel **"Fiche de Recette"** : verdict par lot (Validé / Validé avec réserves / Rejeté, selon le taux de validité -- seuils indicatifs réglables dans `Checkers.py::VERDICT_THRESHOLD_OK`/`VERDICT_THRESHOLD_RESERVE`), la liste des anomalies ouvertes, et une zone de signature (testeur/date/visa) -- le document attendu à l'issue d'une recette, prêt à imprimer ou transmettre tel quel.

Une **anomalie** est une non-conformité agrégée par (lot, étage, classe IFC, PSet) -- pas par élément individuel : à l'échelle réelle (des centaines d'éléments concernés par une même règle manquante), énumérer chaque élément rend le suivi illisible ; le détail reste disponible dans l'onglet "Détails" pour qui veut descendre au niveau de l'élément.

Chaque analyse est **persistée** (SQLite, `history.py`, fichier local `data/checkers_history.db` -- jamais commité, voir `.gitignore`) : date, repère (saisi à l'upload ou nom du fichier Excel par défaut), taux de validité, nombre d'anomalies, empreinte carbone. Le dashboard affiche cet historique en permanence (section "Historique des analyses", visible même avant toute analyse dans la session en cours) avec une courbe de tendance (taux de validité / nombre d'anomalies dans le temps) dès que deux analyses ou plus existent.

Deux analyses peuvent être **comparées** (sélecteurs "Avant"/"Après", ou route `GET /compare/<id_avant>/<id_après>`) : la comparaison distingue les anomalies **corrigées** (présentes avant, absentes après), **nouvelles** (absentes avant, présentes après -- ex. régression) et **persistantes** (dans les deux, avec le nombre d'éléments concernés de chaque côté pour voir une amélioration partielle). Typiquement utilisé entre deux livraisons successives d'une même maquette par un corps d'état, après correction des anomalies signalées.

## Bilan carbone

Le matériau de chaque élément est lu en priorité dans les données du fichier IFC lui-même (associations matériau, y compris les parois multicouches). Le mapping par classe IFC (ex. "un mur est en béton") n'est utilisé qu'en dernier recours, quand le fichier ne porte aucune information matériau, et c'est indiqué dans le rapport. Les facteurs d'émission utilisés (`carbon.py::MATERIAL_CARBON_FACTORS`) sont des **valeurs indicatives** (ordres de grandeur), pas une extraction d'une base FDES/INIES réelle -- éditez cette table si vous disposez de données réelles. Chaque élément est classé "mesuré" (quantité issue du modèle), "estimé" (calculée depuis la géométrie) ou "non calculable" (jamais remplacé par une valeur inventée) ; le rapport Excel et le dashboard affichent la répartition.

## Préconisations par agents IA

Optionnel, désactivé par défaut dans l'image Docker (`ENABLE_AI_RECOMMENDATIONS=false`). Basé sur CrewAI, 3 experts qui se prononcent chacun sur une facette distincte des mêmes faits déjà calculés, puis une synthèse qui croise leurs constats :
- **Expert carbone** : priorise des pistes de réduction à partir de scénarios de substitution déjà chiffrés (`carbon.py::compute_substitution_scenarios` -- quantité réelle × facteur de la table, comparaison d'ordre de grandeur, pas une préconisation d'ingénierie prête à l'emploi).
- **Expert qualité des données** : PSet/paramètres manquants + fiabilité du bilan carbone lui-même (% mesuré/estimé/non calculable).
- **Expert métier / risques projet** : croise les deux jeux de faits précédents pour prioriser ce qui compte le plus sur le terrain (`agents/recommendations.py::compute_business_priorities` -- types d'éléments à la fois mal renseignés et carbonés).

Un modèle local (Ollama) tourne en priorité pour chaque étape, avec escalade vers un modèle cloud (Anthropic) plafonnée en dépense si le résultat local échoue sa vérification. Les agents ne font qu'expliquer/prioriser les chiffres déjà calculés par `Checkers.py`/`carbon.py` -- ils ne recalculent et n'inventent jamais de chiffre. Les 4 étapes tournent séquentiellement (mesuré empiriquement que le parallélisme n'apporte rien avec Ollama, qui traite les requêtes une par une côté serveur).

Pour l'activer : copier `.env.example` en `.env`, s'assurer qu'un service Ollama est joignable (`OLLAMA_BASE_URL`, `ollama pull llama3.2:3b`), et éventuellement fournir `CLOUD_ANTHROPIC_API_KEY` pour l'escalade cloud (plafonnée par `CLOUD_BUDGET_EUR`, 10 € par défaut, suivi de façon persistante dans `data/.cloud_spend.json`).

En Docker, Ollama ne tourne pas dans le conteneur Checkers : `OLLAMA_BASE_URL` doit pointer vers un service Ollama accessible depuis le conteneur (ex. `http://host.docker.internal:11434`, ou un service `ollama` dédié dans `docker-compose.yml`).

Sur Vercel (serverless), CrewAI + Ollama local ne fonctionnent pas (pas de process persistant, pas d'accès à un Ollama local) -- `ENABLE_AI_RECOMMENDATIONS=false` est configuré par défaut dans `vercel.json`. Le bilan carbone corrigé, lui, fonctionne normalement sur Vercel (pur Python, sans dépendance lourde supplémentaire).

## Prérequis

- Docker
- Docker Compose

## Installation

1. Clonez le dépôt :
```bash
git clone https://github.com/halhee/Checkers.git
cd Checkers
```

2. Démarrez l'application avec Docker Compose :
```bash
docker-compose up --build
```

L'application sera accessible à l'adresse : http://localhost:8080

## Structure du Projet

```
.
├── Checkers.py           # Application Flask principale
├── carbon.py             # Bilan carbone (matériaux réels, facteurs, onglet Excel)
├── site_location.py      # Géoréférencement du site (IfcSite.RefLatitude/RefLongitude)
├── corps_etat.py         # Registre des corps d'état et correspondance lot -> discipline
├── history.py            # Historique persistant des analyses + comparaison (anomalies)
├── ids_check.py          # Contrôle IDS (ifctester) : chargement sécurisé, validation par lot, onglet Excel
├── excel_safety.py       # Neutralisation de l'injection de formule dans les rapports Excel
├── agents/               # Préconisations par agents IA (CrewAI, optionnel)
├── Dockerfile           # Configuration Docker
├── docker-compose.yml   # Configuration Docker Compose
├── requirements.txt     # Dépendances Python
├── .env.example         # Configuration des agents IA (à copier en .env)
├── static/             # Fichiers statiques (CSS, images)
├── templates/          # Templates HTML
├── uploads/           # Dossier pour les fichiers uploadés
├── temp/              # Dossier pour les fichiers temporaires
└── data/              # Budget cloud des agents IA + historique des analyses (SQLite)
```

## Utilisation

1. Accédez à l'interface web via http://localhost:8080
2. Ajoutez un ou plusieurs lots (nom + fichier IFC), et un fichier Excel de paramètres requis partagé
3. L'application analysera automatiquement tous les lots et affichera :
   - Les statistiques globales et le comparatif par lot
   - La répartition des éléments par étage
   - Les PSet et paramètres manquants
4. Téléchargez le rapport détaillé au format Excel

## Technologies Utilisées

- Python 3.12
- Flask 3.1
- IfcOpenShell 0.8
- Pandas 3.0
- OpenPyXL 3.1
- Gunicorn 26
- CrewAI (préconisations IA, optionnel)
- Docker
- Bootstrap
- Chart.js

## Développement

Pour le développement local sans Docker :

1. Créez un environnement virtuel :
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

2. Installez les dépendances :
```bash
pip install -r requirements.txt
```

3. Lancez l'application :
```bash
python Checkers.py
```

## Contribution

1. Fork le projet
2. Créez une branche pour votre fonctionnalité (`git checkout -b feature/AmazingFeature`)
3. Committez vos changements (`git commit -m 'Add some AmazingFeature'`)
4. Push vers la branche (`git push origin feature/AmazingFeature`)
5. Ouvrez une Pull Request

## Licence

Ce projet est sous licence MIT. Voir le fichier `LICENSE` pour plus de détails.

## Contact

ZKHCHICHE - khchiche.zakaria@gmail.com
ABOUZNIR - ali.bouznir@gmail.com 

Lien du projet : [https://github.com/halhee/Checkers](https://github.com/halhee/Checkers)
