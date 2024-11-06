# Utiliser l'image Python officielle
FROM python:3.9

# Définir le répertoire de travail
WORKDIR /app

# Copier les fichiers de l'application
COPY . /app

# Installer les dépendances
RUN pip install --no-cache-dir flask pandas openpyxl ifcopenshell werkzeug

# Exposer le port de Flask
EXPOSE 5000

# Commande pour lancer l'application
CMD ["python", "Checkers.py"]
