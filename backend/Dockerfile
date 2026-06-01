# ──────────────────────────────────────────────────────────────────────────────
# Imagen de producción para Railway — FastAPI + Uvicorn
# ──────────────────────────────────────────────────────────────────────────────

FROM python:3.13-slim

# Directorio de trabajo dentro del contenedor
WORKDIR /app

# Copiar primero solo requirements para aprovechar la caché de capas de Docker
COPY requirements.txt .

# Instalar dependencias de Python (sin caché local para imagen más liviana)
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el resto del código fuente
COPY . .

# Puerto expuesto (Railway lo mapea automáticamente)
EXPOSE 8000

# Comando de arranque — Railway también puede sobreescribirlo vía railway.json
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
