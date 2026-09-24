# Backend de conversión Excel -> PDF (LibreOffice) para SIGA
FROM node:20-slim

# LibreOffice Calc + fuentes (LibreOffice necesita fuentes reales para el PDF)
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        fonts-liberation \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY server.js .

ENV PORT=3000
EXPOSE 3000

CMD ["node", "server.js"]
