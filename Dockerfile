FROM python:3.10-slim

# Installation des dépendances système nécessaires
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Création et définition du répertoire de travail
WORKDIR /app

# Copie des fichiers du projet
COPY . /app/

# Mise à jour de pip et installation des outils de base
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir wheel setuptools

# Configuration de pip pour plus de fiabilité
ENV PIP_DEFAULT_TIMEOUT=100 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    PIP_NO_CACHE_DIR=1

# Installation de numpy en premier pour éviter les conflits
RUN pip install --no-cache-dir --timeout 100 numpy==1.24.3

# Installation des dépendances principales
RUN pip install --no-cache-dir --timeout 100 \
    Flask==2.3.3 \
    Werkzeug==2.3.7 \
    python-dotenv==1.0.0

# Installation des dépendances pour le traitement des données
RUN pip install --no-cache-dir --timeout 100 \
    pandas==2.1.0 \
    openpyxl==3.1.2

# Installation d'ifcopenshell
RUN pip install --no-cache-dir --timeout 100 ifcopenshell==0.7.0.240627

# Installation de gunicorn pour la production
RUN pip install --no-cache-dir --timeout 100 gunicorn==21.2.0

# Création des répertoires nécessaires
RUN mkdir -p uploads temp

# Exposition du port
EXPOSE 5050

# Configuration pour la production
ENV FLASK_ENV=production \
    FLASK_APP=Checkers.py \
    PORT=8080

# Création d'un utilisateur non-root pour la sécurité
RUN useradd -m myuser
RUN chown -R myuser:myuser /app
USER myuser

# Commande pour démarrer l'application avec gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "Checkers:app"]
