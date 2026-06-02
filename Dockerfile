# Microservicio PDF — Node + LibreOffice (Calc) para convertir xlsx -> pdf
FROM node:22-slim

# LibreOffice Calc (incluye 'soffice' en el PATH) + fuentes.
# fonts-crosextra-carlito = compatible métrico con Calibri (la fuente del Excel).
# fonts-crosextra-caladea = compatible métrico con Cambria.
# Sin estas fuentes, LibreOffice sustituye Calibri por otra más ancha y las
# columnas se desbordan → el presupuesto sale en 6 páginas en vez de 3.
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      libreoffice-calc \
      fonts-dejavu \
      fonts-crosextra-carlito \
      fonts-crosextra-caladea \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

COPY . .

# Railway inyecta PORT; el server usa process.env.PORT
CMD ["node", "server.js"]
