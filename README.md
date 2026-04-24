# Generador PPTX Socya

Aplicacion Next.js para transformar archivos Excel en presentaciones `.pptx` y reportes auditables con apoyo de `xlsx` y un backend Python (`organizer.py` + `generate_template_presentation.py`).

## Requisitos

- Node.js compatible con `Next 16`
- Python accesible desde `PATH`
- Dependencias de Python necesarias para `organizer.py`, en especial `pandas`

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

## Flujo principal

- `app/api/advanced-generate/route.ts`: organiza y resume el Excel para el paso guiado de preparacion
- `app/api/generate-pptx/route.ts`: genera el `.pptx` final
- `app/api/health/route.ts`: expone el estado operativo del backend y sus dependencias
- `organizer.py`: analisis y organizacion del Excel
- `generate_template_presentation.py`: motor principal de generacion PowerPoint

## Notas operativas

- Las rutas API aceptan archivos Excel `.xlsx`, `.xls` y `.xlsm`
- El backend aplica validaciones de tamano, sanitizacion de nombre y timeout de procesos
- Si Python o `pandas` no estan disponibles, la generacion fallara
- Puedes verificar el estado del runtime en `GET /api/health`

## Checklist de despliegue

- Confirmar que `python --version` responde correctamente en el servidor
- Confirmar que `organizer.py`, `generate_template_presentation.py` y `Plantilla_Presentacion_Socya (1) (1).pptx` existen en la raiz del proyecto desplegado
- Verificar `GET /api/health` antes de habilitar trafico real
- Probar al menos un flujo completo con un Excel pequeno antes de pasar a produccion
