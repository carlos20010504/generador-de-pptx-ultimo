# syntax=docker/dockerfile:1
#
# Imagen única que corre la app completa: el servidor Next.js (Node) que lanza
# el pipeline Python (socya_pipeline) y rasteriza previews con LibreOffice.
# Pensada para Render (servicio web Docker, una instancia persistente).

# ---- Stage 1: build del frontend Next.js ----
FROM node:20-bookworm-slim AS build
WORKDIR /app

# Instalación reproducible de deps JS (incluye devDeps: TS, tailwind, eslint,
# babel-plugin-react-compiler — necesarias para `next build`).
COPY package.json package-lock.json ./
RUN npm ci

# Resto del código y build de producción (webpack, como en package.json).
COPY . .
RUN npm run build

# ---- Stage 2: runtime (Node + Python + LibreOffice) ----
FROM node:20-bookworm-slim AS runtime
WORKDIR /app

ENV NODE_ENV=production \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    SOCYA_CACHE_DIR=/tmp/socya_cache \
    MPLBACKEND=Agg \
    MPLCONFIGDIR=/tmp/mpl \
    PATH="/opt/venv/bin:${PATH}"

# Dependencias de sistema:
# - python3 + venv: ejecutar el pipeline (socya_pipeline)
# - libreoffice-impress + poppler-utils: PPTX -> PDF -> PNG (previews + export)
# - fuentes: reducir la sustitución de tipografías al rasterizar previews
RUN apt-get update && apt-get install -y --no-install-recommends \
        python3 \
        python3-venv \
        python3-pip \
        libreoffice-impress \
        poppler-utils \
        fonts-liberation \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

# Deps Python en un venv aislado. Al estar en el PATH, `spawn('python')` y
# `spawn('python3')` del route Node resuelven a este intérprete con las libs.
COPY requirements.txt ./
RUN python3 -m venv /opt/venv \
    && /opt/venv/bin/pip install --upgrade pip \
    && /opt/venv/bin/pip install -r requirements.txt

# Artefactos de la app. `next start` (no standalone) corre desde /app, así que
# la resolución de rutas Python por process.cwd() encuentra socya_pipeline,
# organizer.py y la plantilla en la raíz — sin tocar código.
# node_modules se copia completo a propósito (sin prune): el costo en tamaño se
# cambia por robustez, ya que el build no se prueba localmente (no hay Docker
# en la máquina de dev). Optimizar a standalone/prune es una mejora futura.
COPY --from=build /app/.next ./.next
COPY --from=build /app/node_modules ./node_modules
COPY --from=build /app/public ./public
COPY package.json next.config.ts ./
COPY socya_pipeline ./socya_pipeline
COPY organizer.py ./organizer.py
COPY ["Plantilla_Presentacion_Socya (1) (1).pptx", "./"]

# Render inyecta $PORT y `next start` lo respeta (default 3000, hostname
# 0.0.0.0). EXPOSE es informativo; el puerto real lo fija el entorno.
EXPOSE 3000

CMD ["npm", "run", "start"]
