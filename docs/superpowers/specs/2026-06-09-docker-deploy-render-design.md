# Despliegue Docker en Render — Diseño

Fecha: 2026-06-09
Estado: aprobado por el usuario (brainstorming)

## Objetivo

Empaquetar la app tal cual (Next.js 16 que lanza el pipeline Python `socya_pipeline`)
en **una sola imagen Docker** y desplegarla en **Render** como servicio web persistente
de una instancia, con **generación de PPTX + previews PNG + export PDF funcionando en
Linux**. Sin reescribir la arquitectura.

## Decisiones tomadas

- **Plataforma:** Render (contenedor persistente de una instancia → encaja con el token
  en memoria `globalThis.__SOCYA_PENDING__` y los temporales en `/tmp` con TTL de 30 min).
- **Alcance previews:** completo — previews PNG **y** export PDF en Linux vía LibreOffice
  ("que todo funcione").
- **Empaquetado:** `next start` (NO `output: standalone`) desde `/app`, para que la
  resolución de rutas Python por `process.cwd()` siga funcionando sin tocar código.
- **Plantilla:** se fuerza a git (`git add -f`), 3.5 MB directo sin LFS.

## Arquitectura del contenedor (multi-stage)

- Base: `node:20-bookworm-slim` (Next 16 requiere Node ≥20; confirmado en
  `node_modules/next/dist/docs/.../next.md`: `next start` respeta `PORT`, hostname 0.0.0.0).
- **Stage build:** `npm ci` + `npm run build` + `npm prune --omit=dev`.
- **Stage runtime:** Node 20 + Python en venv `/opt/venv` (en PATH, así `spawn('python')`
  y `'python3'` resuelven al venv con las deps) + LibreOffice Impress + poppler-utils +
  fuentes. Copia: `.next`, `node_modules` (prod), `public`, `package.json`, `next.config.ts`,
  `socya_pipeline/`, `organizer.py`, plantilla.

## Cambios de código (mínimos)

1. `package.json` → `"start": "next start"` (sin `-p 3001`; honra `$PORT`).
2. `requirements.txt` → agregar `pdf2image` (puro Python; poppler es dep de sistema).
3. `socya_pipeline/preview.py` → rama no-Windows (LibreOffice):
   - `_find_soffice()`: resuelve `soffice`/`libreoffice` en PATH o ruta absoluta Debian.
   - `_soffice_convert_to_pdf()`: `soffice --headless --convert-to pdf:impress_pdf_Export
     --outdir <tmp> <pptx>`, con `-env:UserInstallation=file:///tmp/lo_<unico>` para
     soportar requests concurrentes (evita el lock single-instance de LibreOffice).
   - `_render_with_libreoffice()`: PPTX→PDF (1 sola conversión) y rasteriza cada página a
     PNG con `pdf2image` (poppler) → un PNG por slide (`soffice` solo exporta la 1ª).
   - `_export_pdf_libreoffice()`: misma conversión, mueve el PDF al destino pedido.
   - `generate_previews` y `export_to_pdf` despachan por `sys.platform`: win32 →
     PowerPoint COM (Windows local intacto), resto → LibreOffice. Se conserva el
     contrato fail-soft `{ok, renderer, slides/path, error}`.

## Artefactos nuevos

- `Dockerfile` (multi-stage).
- `.dockerignore` (excluye `node_modules`, `.next`, `.git`, `.env*`, `*_previews/`, logs,
  backups; **mantiene** `package-lock.json`, `socya_pipeline/`, `organizer.py`, plantilla).
- `render.yaml` (Blueprint: web service Docker, healthCheck `/api/health`, env vars).
- `.env.example` (documenta claves sin valores).

## Variables de entorno (en Render, como secretos)

`GROQ_API_KEY`, `OPENROUTER_API_KEY`, `CEREBRAS_API_KEY` (opcional),
`SOCYA_CACHE_DIR=/tmp/socya_cache`. `PORT` lo provee Render.

## Caveats honestos

- **Fuentes en previews:** LibreOffice sustituye fuentes al rasterizar → el PNG de preview
  puede verse ligeramente distinto a PowerPoint. **El PPTX generado NO se afecta** (solo la
  imagen de previsualización). Se instalan fuentes comunes; si la plantilla usa una fuente
  de marca en `.ttf`, se mete a la imagen y queda idéntico.
- **Memoria:** LibreOffice + pandas + matplotlib + Node pueden superar 512 MB (Free/Starter
  de Render) al renderizar decks grandes → riesgo de OOM. Para fiabilidad plena, el plan
  Standard (2 GB) es lo seguro. Decisión de costo del usuario; el `render.yaml` lo deja claro.
- **Docker no está instalado en la máquina del usuario** → el build no se prueba localmente;
  ocurre en Render. Se compensa con: `pytest` + `npm test` locales y un test del renderer
  LibreOffice con `subprocess`/`pdf2image` mockeados (corre en Windows sin LibreOffice).

## Verificación

- Test nuevo del renderer (mocked) + `pytest` + `npm test` verdes.
- En Render tras el primer deploy: subir un Excel real → PPTX descargable, previews PNG
  visibles, PDF exportable; `/api/health` 200.

## Fuera de alcance (YAGNI)

`output: standalone`, multi-instancia/escalado, cola de trabajos, CI/CD, usuario no-root en
la imagen, Kubernetes. La instancia única persistente + SSE (mantiene viva la conexión)
cubren el estado en memoria y las requests largas.
