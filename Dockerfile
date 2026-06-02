# Microservicio PDF — Node + LibreOffice (Calc) para convertir xlsx -> pdf
FROM node:22-slim

# LibreOffice Calc (incluye 'soffice' en el PATH) + fuentes para render correcto
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      libreoffice-calc \
      fonts-dejavu \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

COPY . .

# Railway inyecta PORT; el server usa process.env.PORT
CMD ["node", "server.js"]
