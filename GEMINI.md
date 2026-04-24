# 🚀 Generador PPTX Socya - GEMINI.md

Este proyecto es un sistema inteligente de generación de presentaciones de PowerPoint (`.pptx`) a partir de archivos Excel, diseñado para automatizar reportes de auditoría y comisiones con un enfoque en calidad visual, fidelidad al dato y procesamiento eficiente.

## 📌 Visión General del Proyecto

La arquitectura combina **Python** para el análisis del Excel y la generación final del `.pptx`, junto con **TypeScript (Next.js)** para la experiencia web y la orquestación del flujo.

### 🛠 Tecnologías Principales
- **Frontend:** Next.js 16 (App Router), React 19, TailwindCSS 4.
- **Procesamiento de Datos:** Python 3 + Pandas (vía `organizer.py`).
- **Generación PPTX:** Python + `python-pptx` + `matplotlib` (vía `generate_template_presentation.py`).
- **Validación de archivos:** `xlsx` en el frontend y backend.

## 🏗 Arquitectura del Sistema

1.  **Capa de Entrada:** El usuario sube un archivo Excel mediante `ExcelUploader.tsx`.
2.  **Capa de Análisis (Python):** La API ejecuta `organizer.py`, que utiliza un motor de priorización (Scoring Algorithm) para:
    *   Identificar KPIs globales (totales, promedios, conteos).
    *   Generar distribuciones (por estado, ciudad, centro de costos).
    *   Extraer tablas críticas (Hallazgos, COSO, Oportunidades de Mejora).
    *   Aplicar "Compresión Inteligente" (agrupa datos masivos en "Top 5 + Otros").
3.  **Capa de Orquestación UI/API:** `app/api/advanced-generate/route.ts` prepara el Excel organizado y `app/api/generate-pptx/route.ts` dispara la generación final.
4.  **Capa de Renderizado:** `generate_template_presentation.py` aplica la plantilla corporativa, genera gráficos/tablas y guarda el archivo final.

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

- **Tipado Estricto:** La capa web usa TypeScript para validar entradas, tamaños y estados de UI.
- **Fidelidad al Dato:** El flujo productivo siempre depende del Excel subido en la solicitud actual; no se permite fallback a archivos legados.
- **Sanitización:** Los nombres de archivo y el texto serializado deben limpiarse para evitar errores en la generación y en la descarga.
- **Paginación Inteligente:** Las tablas extensas se fragmentan y priorizan para conservar legibilidad en la plantilla final.

## 📂 Archivos Críticos

- `app/api/advanced-generate/route.ts`: Punto de entrada para organización guiada del Excel.
- `app/api/generate-pptx/route.ts`: Punto de entrada para la generación final del PowerPoint.
- `organizer.py`: El "cerebro" analítico en Python.
- `generate_template_presentation.py`: Motor principal de renderizado y composición del `.pptx`.
- `components/ExcelUploader.tsx`: Flujo guiado de carga, organización y generación.
