# Generador PPTX Socya

Este proyecto es una aplicación **Next.js 16** diseñada para transformar archivos Excel (.xlsx, .xls, .xlsm) en presentaciones PowerPoint (.pptx) corporativas. Utiliza un backend híbrido donde la orquestación y la API residen en Node.js, mientras que el análisis de datos y la generación de diapositivas se delegan al paquete **Python** `socya_pipeline`.

## Arquitectura y Componentes Principales

El flujo de trabajo está dividido en 5 pasos ejecutados por `socya_pipeline`:

1. **Parser** (`socya_pipeline/parser.py`)
   * Lee `.xlsx`/`.xls`/`.xlsm` con `pandas`.
   * Devuelve dataclasses con shape, tipos de columna, samples y estadísticas básicas.
   * No realiza ninguna inferencia de negocio en este paso.

2. **Inventario** (`socya_pipeline/inventory.py`)
   * Cataloga bloques disponibles: tablas, KPIs, distribuciones categóricas, series temporales.
   * Registra procedencia (sheet, columnas, filas) y `quality_flags` (low_fill_ratio, single_dominant_category, all_zero, etc.) que el validador usa después.

3. **Planificador IA** (`socya_pipeline/planner.py`)
   * Realiza **una sola llamada** al modelo primario: **Hermes 3 405B** (free, vía OpenRouter).
   * Cadena de fallback automático: Hermes 3 405B → Llama 3.3 70B → Mistral Small 24B → Qwen 2.5 72B.
   * Cache por hash de contenido en `.cache/plans/{sha256}.json` para evitar llamadas repetidas.
   * Payload destilado de hasta 6K tokens para minimizar latencia y costo.

4. **Validador** (`socya_pipeline/validator.py`)
   * Drop estricto de slides cuyas referencias de bloques no existen en el inventario.
   * Descarta slides con `quality_flags` problemáticos.
   * Verifica que bullets y cifras mencionadas en el plan aparezcan en los datos reales del Excel (anti datos fantasma).

5. **Renderer** (`socya_pipeline/renderer.py`)
   * Ejecuta-only: toma el plan validado y produce el `.pptx` con `python-pptx` + `matplotlib`.
   * Sin fallbacks de plantilla. Si el plan indica "tabla con bloque T1", extrae exactamente esos datos.
   * La plantilla corporativa (`Plantilla_Presentacion_Socya (1) (1).pptx`) aporta la capa visual (layouts, colores, tipografía), no el contenido.

CLI disponible: `python -m socya_pipeline {plan|generate} ...`

## Tecnologías Clave

* **Frontend/API:** Next.js 16 (TypeScript), Tailwind CSS.
* **Backend de Datos:** Python 3.x, `pandas`, `numpy`, `python-pptx`, `matplotlib`.
* **IA:** Hermes 3 405B como planificador principal (vía OpenRouter) + 3 modelos free de fallback automático (Llama 3.3 70B, Mistral Small 24B, Qwen 2.5 72B).

## Rutas de API Críticas

* `POST /api/excel-intelligence`: Análisis del Excel + sugerencias de prompt para el panel UI (perfil `fast`, ~25s timeout).
* `POST /api/advanced-generate`: Devuelve plan validado + audit (perfil `patient`).
* `POST /api/generate-pptx`: **Streaming SSE** con 5 fases. Devuelve `data: { downloadToken, audit }` en el evento `done`. El binario se descarga vía `GET /api/generate-pptx?token=...`.
* `GET /api/health`: Verifica que Python, `socya_pipeline` y la plantilla estén disponibles y operativos.

## Comandos de Desarrollo

* **Instalación:** `npm install`
* **Entorno de Desarrollo:** `npm run dev` (disponible en http://localhost:3001)
* **Construcción:** `npm run build`
* **Linting:** `npm run lint`
* **Pruebas:** `npm test`

## Convenciones y Notas de Desarrollo

* **Runtime:** El backend requiere que `python` esté en el `PATH` y que el paquete `socya_pipeline` esté instalado o disponible en la raíz del proyecto con sus dependencias (`pandas`, `python-pptx`, `matplotlib`, `requests`). Verificar con `python -c 'import socya_pipeline'`.
* **Gestión de Archivos:** Las APIs utilizan directorios temporales para procesar los archivos subidos y los eliminan al finalizar.
* **Validaciones:** Se aplican límites estrictos de tamaño de archivo y timeouts de ejecución para los scripts de Python (configurados en `utils/excel-ai-panel.cjs` y `utils/server-runtime.ts`).
* **Plantilla:** `Plantilla_Presentacion_Socya (1) (1).pptx` en la raíz es la capa visual (layouts y colores corporativos). El contenido lo define el planificador IA, no la plantilla.
* **Codificación:** Los scripts de Python están configurados para usar `UTF-8` para evitar errores de caracteres especiales en Windows.
* **Variables de entorno:** `OPENROUTER_API_KEY` es requerida. `SOCYA_AI_PROFILE` (`fast` | `patient`) controla el timeout y profundidad de la cadena de fallback.

## Estructura de Carpetas

* `app/api/`: Endpoints de la aplicación (`excel-intelligence`, `advanced-generate`, `generate-pptx`, `health`).
* `components/`: Componentes React (Uploader, Paneles de IA).
* `utils/`: Utilidades para manejo de archivos, integridad de presentaciones y validación de runtime.
* `socya_pipeline/`: Paquete Python con el pipeline de 5 pasos (`parser.py`, `inventory.py`, `planner.py`, `validator.py`, `renderer.py`) y utilidades de soporte (`ai_chain.py`, `ai_cache.py`, `errors.py`, `extractor.py`).
* `Raíz`: Contiene `organizer.py` (wrapper delgado de compatibilidad), la plantilla base y el CLI entry-point del paquete.
