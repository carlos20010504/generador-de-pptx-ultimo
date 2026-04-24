# 🚀 Sistema Inteligente de Generación de Presentaciones (PPTX AI)

## 📌 Arquitectura del Sistema
El flujo actual se apoya en `organizer.py` para analizar el Excel y en `generate_template_presentation.py` para renderizar la presentación final con plantilla corporativa.

### Componentes Principales
1. **Carga y validación**: El frontend valida tipo y tamaño del Excel antes de enviarlo al backend.
2. **Organización del Excel**: `organizer.py` detecta KPIs, tablas y gráficos candidatos a partir del archivo subido.
3. **Generación guiada**: `app/api/advanced-generate/route.ts` devuelve una versión organizada para el paso 1 del flujo.
4. **Renderizado final**: `generate_template_presentation.py` usa `python-pptx` y `matplotlib` para construir el `.pptx` final con la plantilla institucional.

---

## 🛠️ Guía de Troubleshooting (Errores Comunes)

| Error de PowerPoint | Causa Raíz | Solución Implementada |
|---------------------|------------|-----------------------|
| *"PowerPoint encontró un problema con el contenido"* | Texto corrupto, valores inconsistentes o elementos inválidos en el XML del archivo. | El generador limpia texto, normaliza valores y limita bloques antes de escribir en la plantilla. |
| *"PowerPoint no pudo leer algún contenido y tuvo que quitarlo"* | Etiquetas o tablas demasiado extensas para el layout. | El sistema resume categorías, pagina tablas y recorta contenido no legible. |
| *Crashes al abrir el archivo* | Dependencias faltantes en el runtime o plantilla ausente. | La app valida `python`, `organizer.py`, `generate_template_presentation.py` y la plantilla desde `GET /api/health`. |

---

## 🧠 Manual de Usuario: Sistema de Compresión Inteligente

Cuando suba un archivo Excel masivo (ej. > 1,000 filas o decenas de hojas), el sistema actuará de la siguiente manera:
1. **Filtrado Top N**: Los gráficos muestran solo las categorías más relevantes y consolidan el resto cuando hace falta.
2. **Priorización visual**: El generador favorece primero KPIs, gráficos y tablas legibles dentro del espacio real de la plantilla.
3. **Fragmentación segura**: Si una tabla es muy larga, se divide en varias diapositivas en lugar de saturar una sola.

---

## 📊 Certificado de Calidad y Rendimiento
- **Validación operativa:** `npm run lint`, `npm test` y `npm run build` deben pasar antes del despliegue.
- **Dependencias críticas:** Python accesible, plantilla `.pptx` disponible y scripts raíz presentes.
- **Fidelidad de salida:** La presentación se genera exclusivamente a partir del archivo enviado en la solicitud actual.
- **Manejo de errores:** Las rutas API aplican límites de tamaño, timeouts y mensajes de error explícitos para evitar salidas inconsistentes.
