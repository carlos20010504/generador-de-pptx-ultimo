# GEMINI.md - Contexto de Instrucción del Proyecto

## Información General del Proyecto
Este es el **Generador PPTX Socya**, una aplicación web moderna (Next.js v15/16) diseñada para transformar archivos Excel complejos en presentaciones de PowerPoint (.pptx) profesionales y auditables. Combina la agilidad de una interfaz web con el poder de procesamiento de datos de Python y capacidades de IA generativa (Gemini).

### Arquitectura
- **Frontend:** Next.js (App Router) con TypeScript. Interfaz orientada a la experiencia del usuario con animaciones y flujos guiados.
- **API (Backend Node.js):** Actúa como orquestador, gestionando la carga de archivos, validaciones y la ejecución de subprocesos Python.
- **Motor de Procesamiento (Python):** Scripts especializados que realizan el análisis estadístico, la organización de datos y la renderización final del PPTX.
- **Integración de IA:** Utiliza la familia de modelos Gemini (Flash 1.5/2.0) para generar resúmenes ejecutivos e insights avanzados directamente desde los datos del Excel.

## Comandos Clave y Configuración

### Requisitos Previos
- **Node.js:** Versión compatible con Next.js 15+.
- **Python 3.x:** Debe estar disponible en el `PATH` como `python`.
- **Dependencias de Python:** `pandas`, `numpy`, `python-pptx`, `matplotlib`, `google-generativeai`.

### Desarrollo y Ejecución
- **Instalación:** `npm install`
- **Servidor de Desarrollo:** `npm run dev` (disponible en `http://localhost:3001` por defecto).
- **Verificación de Salud:** `GET /api/health` permite verificar si el entorno de Python y los scripts necesarios están correctamente configurados.

## Estructura de Archivos Críticos
- `app/api/`: Contiene las rutas de la API (`generate-pptx`, `advanced-generate`, `health`).
- `organizer.py`: Lógica principal de análisis de Excel y preparación de datos para los slides.
- `generate_template_presentation.py`: Renderizador de PowerPoint usando `python-pptx`.
- `utils/server-runtime.ts`: Gestiona la detección del runtime de Python y la disponibilidad de scripts.
- `Plantilla_Presentacion_Socya (1) (1).pptx`: La plantilla base obligatoria para la generación.

## Convenciones de Desarrollo
- **Idiomas:** Código y comentarios mayoritariamente en español/inglés. Los mensajes de usuario y logs de negocio están en español.
- **Manejo de Python:** La comunicación entre Node.js y Python se realiza mediante `child_process.execFile`, pasando datos a través de archivos temporales o argumentos de línea de comandos.
- **Validaciones:** Se aplican estrictos límites de tamaño y timeouts para evitar bloqueos del servidor durante el procesamiento de Excels pesados.
- **Estilo de Código:** Sigue las convenciones estándar de Next.js para el frontend y PEP 8 para los scripts de Python.

## Notas de Operación
- El sistema utiliza un sistema de prioridad de modelos Gemini (`flash-lite`, `flash`, etc.) con mecanismos de enfriamiento (cooldown) para manejar las cuotas de la API gratuita.
- Existen archivos de caché en `.cache/` para los resultados de la IA, optimizando costos y tiempos de respuesta.
