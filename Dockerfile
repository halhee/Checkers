FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .

# Installation des dépendances système nécessaires
RUN apt-get update && apt-get install -y \
    gcc \
    python3-dev \
    && rm -rf /var/lib/apt/lists/*

RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Create necessary directories
RUN mkdir -p uploads temp

EXPOSE 8080

CMD ["python", "app.py"]
