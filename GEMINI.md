# 🚀 Generador PPTX Socya - GEMINI.md

Este proyecto es un sistema inteligente de generación de presentaciones de PowerPoint (`.pptx`) a partir de archivos Excel, diseñado para automatizar reportes de auditoría y comisiones con un enfoque en calidad visual, accesibilidad (WCAG) y procesamiento eficiente de datos.

## 📌 Visión General del Proyecto

La arquitectura combina la potencia de **Python (Pandas)** para el análisis de datos masivos y **TypeScript (Next.js/PptxGenJS)** para la orquestación y el renderizado final de las diapositivas.

### 🛠 Tecnologías Principales
- **Frontend:** Next.js 16 (App Router), React 19, TailwindCSS 4.
- **Procesamiento de Datos:** Python 3 + Pandas (vía `organizer.py`).
- **Generación PPTX:** `pptxgenjs` (vía `AdvancedPptxGenerator`).
- **Parsing de Excel:** `xlsx` (SheetJS).

## 🏗 Arquitectura del Sistema

1.  **Capa de Entrada:** El usuario sube un archivo Excel mediante `ExcelUploader.tsx`.
2.  **Capa de Análisis (Python):** La API ejecuta `organizer.py`, que utiliza un motor de priorización (Scoring Algorithm) para:
    *   Identificar KPIs globales (totales, promedios, conteos).
    *   Generar distribuciones (por estado, ciudad, centro de costos).
    *   Extraer tablas críticas (Hallazgos, COSO, Oportunidades de Mejora).
    *   Aplicar "Compresión Inteligente" (agrupa datos masivos en "Top 5 + Otros").
3.  **Capa de Orquestación:** `app/api/advanced-generate/route.ts` recibe el JSON de Python y define la estructura de las diapositivas.
4.  **Capa de Renderizado:** `AdvancedPptxGenerator` (en `utils/`) aplica temas corporativos, valida contrastes WCAG y genera el archivo final.

## 🚀 Comandos Clave

| Comando | Descripción |
|---------|-------------|
| `npm run dev` | Inicia el servidor de desarrollo en `localhost:3000`. |
| `npm run build` | Compila la aplicación para producción. |
| `npm run start` | Inicia la aplicación compilada. |
| `npm run lint` | Ejecuta el linter (ESLint). |
| `python organizer.py <path_excel>` | Ejecuta el análisis de datos de forma independiente (útil para pruebas). |

> **Nota:** Se requiere tener instalado **Python** con la librería **pandas** en el entorno de ejecución para que la generación de reportes funcione correctamente.

## 🧠 Convenciones de Desarrollo

- **Tipado Estricto:** Se prefiere el uso exhaustivo de TypeScript para interfaces de datos y reportes de optimización (`OptimizationReport`).
- **WCAG 2.1:** Toda inyección de color debe validarse mediante `getContrastRatio` en el motor de generación para asegurar legibilidad.
- **Sanitización:** Los strings inyectados en PPTX deben pasar por limpieza Regex (`replace(/[^\w\s-]/gi, '')`) para evitar corrupción del XML de PowerPoint.
- **Paginación Inteligente:** Las tablas extensas deben paginarse automáticamente (máximo 12 filas por slide) o derivarse a un **Apéndice Interactivo** si exceden el límite de diapositivas (25 slides).

## 📂 Archivos Críticos

- `app/api/advanced-generate/route.ts`: Punto de entrada de la lógica de negocio.
- `utils/advanced-pptx-generator.ts`: El "motor" de renderizado PPTX.
- `organizer.py`: El "cerebro" analítico en Python.
- `utils/socya-renderer.ts`: Definiciones estéticas específicas para la marca Socya.
- `README_ADVANCED_PPTX.md`: Documentación técnica detallada sobre el motor de scoring y troubleshooting.
