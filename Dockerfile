# pandas>=3.0 (requis par crewai, voir requirements.txt) exige Python >=3.11.
FROM python:3.12-slim

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
    pip install --no-cache-dir --timeout 100 -r requirements.txt

# Création des répertoires nécessaires (data/ : suivi persistant du budget
# cloud des agents IA, voir agents/budget.py)
RUN mkdir -p uploads temp data

# Configuration des variables d'environnement
ENV FLASK_ENV=production \
    FLASK_APP=Checkers.py \
    PORT=8080

# Préconisations IA désactivées par défaut dans l'image : CrewAI/Ollama local
# ne sont pas installés dans ce conteneur (voir README -- Ollama doit tourner
# en dehors, OLLAMA_BASE_URL doit pointer vers ce service). Surchargez cette
# variable (ou passez-la via docker-compose/.env) pour les activer.
ENV ENABLE_AI_RECOMMENDATIONS=false

# Création d'un utilisateur non-root pour la sécurité
RUN useradd --create-home --shell /bin/bash myuser
RUN chown -R myuser:myuser /app
USER myuser

# Exposition du port 5050
EXPOSE 5050

# Commande pour démarrer l'application avec Gunicorn
CMD ["gunicorn", "--workers", "3", "--threads", "2", "--timeout", "60", "--bind", "0.0.0.0:8080", "Checkers:app"]
