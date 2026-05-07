# Generador PPTX Socya

Este proyecto es una aplicación **Next.js 16** diseñada para transformar archivos Excel (.xlsx, .xls, .xlsm) en presentaciones PowerPoint (.pptx) corporativas. Utiliza un backend híbrido donde la orquestación y la API residen en Node.js, mientras que el análisis de datos y la generación de diapositivas se delegan a scripts de **Python**.

## Arquitectura y Componentes Principales

El flujo de trabajo se divide en dos fases principales:

1.  **Análisis y Organización (`organizer.py`):**
    *   Procesa el Excel usando `pandas`.
    *   Detecta automáticamente KPIs, tablas y gráficos.
    *   Utiliza etiquetas especiales (opcionales) en el Excel como `TITLE:`, `SUBTITLE:` y `TYPE:` para guiar la generación.
    *   Integra inteligencia artificial (vía OpenRouter/Hermes) para generar insights y resúmenes ejecutivos si el usuario proporciona un "prompt".
2.  **Generación de PowerPoint (`generate_template_presentation.py`):**
    *   Utiliza la librería `python-pptx` para manipular una plantilla corporativa (`Plantilla_Presentacion_Socya (1) (1).pptx`).
    *   Genera gráficos usando `matplotlib`.
    *   Aplica lógica de paginación para tablas largas y límites visuales para mantener la legibilidad.

## Tecnologías Clave

*   **Frontend/API:** Next.js 16 (TypeScript), Tailwind CSS (inferido por `globals.css` y `postcss.config.mjs`).
*   **Backend de Datos:** Python 3.x, `pandas`, `numpy`, `python-pptx`, `matplotlib`.
*   **IA:** Integración con modelos de lenguaje a través de OpenRouter (prioridad: Hermes-3-Llama-3.1-405b).

## Rutas de API Críticas

*   `POST /api/advanced-generate`: Organiza el contenido del Excel y devuelve una estructura de diapositivas sugerida.
*   `POST /api/generate-pptx`: Genera y descarga el archivo `.pptx` final.
*   `GET /api/health`: Verifica que Python, los scripts y la plantilla estén disponibles y operativos.

## Comandos de Desarrollo

*   **Instalación:** `npm install`
*   **Entorno de Desarrollo:** `npm run dev` (disponible en http://localhost:3001)
*   **Construcción:** `npm run build`
*   **Linting:** `npm run lint`
*   **Pruebas:** `npm test`

## Convenciones y Notas de Desarrollo

*   **Runtime:** El backend requiere que `python` esté en el `PATH` y tenga instaladas las dependencias necesarias (`pandas`, `python-pptx`, `matplotlib`, `requests`).
*   **Gestión de Archivos:** Las APIs utilizan directorios temporales para procesar los archivos subidos y los eliminan al finalizar.
*   **Validaciones:** Se aplican límites estrictos de tamaño de archivo y timeouts de ejecución para los scripts de Python (configurados en `utils/excel-ai-panel.cjs` y `utils/server-runtime.ts`).
*   **Plantilla:** La generación depende críticamente de la existencia de `Plantilla_Presentacion_Socya (1) (1).pptx` en la raíz.
*   **Codificación:** Los scripts de Python están configurados para usar `UTF-8` para evitar errores de caracteres especiales en Windows.

## Estructura de Carpetas

*   `app/api/`: Endpoints de la aplicación.
*   `components/`: Componentes React (Uploader, Paneles de IA).
*   `utils/`: Utilidades para manejo de archivos, integridad de presentaciones y validación de runtime.
*   `scripts/` (en `.agents`, `.claude`, `.trae`): Herramientas adicionales para manipulación de PPTX (limpieza, miniaturas, validación).
*   `Raíz`: Contiene los scripts principales de Python (`organizer.py`, `generate_template_presentation.py`) y la plantilla base.
