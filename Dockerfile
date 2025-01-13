FROM python:3.10-slim

# Installation des dépendances système nécessaires
RUN apt-get update && apt-get install -y \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Création et définition du répertoire de travail
WORKDIR /app

# Copie des fichiers du projet
COPY . /app/

# Mise à jour de pip et installation des dépendances Python
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir --timeout 100 \
    numpy==1.24.3 \
    Flask==2.3.3 \
    Werkzeug==2.3.7 \
    python-dotenv==1.0.0 \
    pandas==2.1.0 \
    openpyxl==3.1.2 \
    ifcopenshell==0.7.0.240627 \
    gunicorn==21.2.0

# Création des répertoires nécessaires
RUN mkdir -p uploads temp

# Configuration des variables d'environnement
ENV FLASK_ENV=production \
    FLASK_APP=Checkers.py \
    PORT=8080

# Création d'un utilisateur non-root pour la sécurité
RUN useradd --create-home --shell /bin/bash myuser
RUN chown -R myuser:myuser /app
USER myuser

# Exposition du port 5050
EXPOSE 5050

# Commande pour démarrer l'application avec Gunicorn
CMD ["gunicorn", "--workers", "3", "--threads", "2", "--timeout", "60", "--bind", "0.0.0.0:8080", "Checkers:app"]
