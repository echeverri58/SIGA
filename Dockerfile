# SIGA · Aplicación + conversión Excel -> PDF (todo en uno) para la nube.
# Construye una imagen con LibreOffice, que es quien convierte el Excel a PDF.
FROM node:20-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        fonts-liberation \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY . .

ENV PORT=3000
EXPOSE 3000

CMD ["node", "backend/server.js"]
