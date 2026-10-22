FROM python:3.11-slim

WORKDIR /app

# Install dipendenze
COPY requirements_backend.txt .
RUN pip install --no-cache-dir -r requirements_backend.txt

# Copia codice sorgente
COPY . .

# Crea database SQLite se non esiste
RUN python -c "from app import app, db; import os; os.makedirs('data', exist_ok=True)"

EXPOSE 10000

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000", "--workers", "2"]