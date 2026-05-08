# Generador PPTX Socya

Aplicacion Next.js para transformar archivos Excel en presentaciones `.pptx` y reportes auditables con apoyo de `xlsx` y un backend Python (`socya_pipeline`).

## Requisitos

- Node.js compatible con `Next 16`
- Python accesible desde `PATH`
- Paquete `socya_pipeline` disponible en la raiz del proyecto con sus dependencias (`pandas`, `python-pptx`, `matplotlib`, `requests`)

## Ejecucion local

Instala dependencias:

```bash
npm install
```

Levanta el servidor:

```bash
npm run dev
```

La aplicacion queda disponible en [http://localhost:3001](http://localhost:3001).

## Flujo principal (rediseno Hermes-driven)

Pipeline en 5 pasos, IA actua como planificador unico:

1. **Parser** (`socya_pipeline/parser.py`) — Lee `.xlsx`/`.xls`/`.xlsm` con pandas. Devuelve dataclasses con shape, tipos, samples y stats. Sin inferencia de negocio.
2. **Inventario** (`socya_pipeline/inventory.py`) — Cataloga bloques disponibles (tablas, KPIs, distribuciones categoricas, series temporales) con procedencia (sheet, columnas, filas) y `quality_flags` (low_fill_ratio, single_dominant_category, all_zero, etc.).
3. **Planificador IA** (`socya_pipeline/planner.py`) — Una sola llamada a Hermes 3 405B (free, via OpenRouter). Cadena de 4 modelos free como fallback automatico: Hermes -> Llama 3.3 70B -> Mistral Small 24B -> Qwen 2.5 72B. Cache por hash de contenido (`.cache/plans/{sha256}.json`). Payload destilado de hasta 6K tokens.
4. **Validador** (`socya_pipeline/validator.py`) — Drop estricto de slides cuyas referencias no existen, cuyos quality_flags son malos, o cuyos bullets contienen numeros/nombres que no aparecen en el Excel real (anti datos fantasmas).
5. **Renderer** (`socya_pipeline/renderer.py`) — Ejecuta-only con `python-pptx` + `matplotlib`. Sin fallbacks de plantilla. Si el plan dice "tabla con bloque T1", saca exactamente esos datos.

CLI: `python -m socya_pipeline {plan|generate} ...`

## Rutas API

- `POST /api/excel-intelligence` — Analisis del Excel + sugerencias de prompt para el panel UI (perfil `fast`, ~25s timeout)
- `POST /api/advanced-generate` — Devuelve plan validado + audit (perfil `patient`)
- `POST /api/generate-pptx` — **Streaming SSE** con 5 fases. Devuelve `data: { downloadToken, audit }` en evento `done`. El binario se descarga via `GET /api/generate-pptx?token=...`
- `GET /api/health` — Verifica que Python + `socya_pipeline` + plantilla esten disponibles

## Variables de entorno

- `OPENROUTER_API_KEY` — Requerida para que el planificador IA funcione (tambien se busca en `.env`/`.env.local`)
- `SOCYA_AI_PROFILE` — `fast` (default) o `patient`. Controla el timeout y la profundidad de la cadena de fallback

## Notas operativas

- Las rutas API aceptan archivos Excel `.xlsx`, `.xls` y `.xlsm`
- El backend aplica validaciones de tamano, sanitizacion de nombre y timeout de procesos
- Si Python o el paquete `socya_pipeline` no estan disponibles, la generacion fallara
- Puedes verificar el estado del runtime en `GET /api/health`

## Checklist de despliegue

- Confirmar que `python --version` responde correctamente en el servidor
- Confirmar que `python -c 'import socya_pipeline'` funciona sin errores
- Confirmar que `socya_pipeline/` y `Plantilla_Presentacion_Socya (1) (1).pptx` existen en la raiz del proyecto desplegado
- Verificar `GET /api/health` antes de habilitar trafico real
- Probar al menos un flujo completo con un Excel pequeno antes de pasar a produccion
