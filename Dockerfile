FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Create necessary directories
RUN mkdir -p uploads temp

EXPOSE 8080

# Run with gunicorn
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "4", "--timeout", "120", "--log-level", "debug", "app:app"]
