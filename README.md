# Checkers - Validation IFC

Checkers est une application web permettant de valider des fichiers IFC (Industry Foundation Classes) en vérifiant la présence et la conformité des Property Sets (PSet) et des paramètres requis.

![Checkers Logo](static/sans titre.png)

## Fonctionnalités

- Upload de fichiers IFC
- Validation automatique des PSet et paramètres
- Dashboard interactif avec statistiques en temps réel
- Graphique en donut pour la visualisation des résultats
- Génération de rapports détaillés au format Excel
- Support des fichiers volumineux (jusqu'à 300 Mo)
- Interface utilisateur moderne et intuitive

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
├── Dockerfile           # Configuration Docker
├── docker-compose.yml   # Configuration Docker Compose
├── requirements.txt     # Dépendances Python
├── static/             # Fichiers statiques (CSS, images)
├── templates/          # Templates HTML
├── uploads/           # Dossier pour les fichiers uploadés
└── temp/              # Dossier pour les fichiers temporaires
```

## Utilisation

1. Accédez à l'interface web via http://localhost:8080
2. Téléchargez votre fichier IFC
3. L'application analysera automatiquement le fichier et affichera :
   - Les statistiques globales
   - La répartition des éléments par étage
   - Les PSet et paramètres manquants
4. Téléchargez le rapport détaillé au format Excel

## Technologies Utilisées

- Python 3.10
- Flask 2.3.3
- IfcOpenShell 0.7.0
- Pandas 2.1.0
- OpenPyXL 3.1.2
- Gunicorn 21.2.0
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

Lien du projet : [https://github.com/halhee/Checkers](https://github.com/halhee/Checkers)
