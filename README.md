# Generador PPTX Socya

Aplicacion Next.js para transformar archivos Excel en presentaciones `.pptx` y reportes auditables con apoyo de `pptxgenjs`, `xlsx` y un analizador Python (`organizer.py`).

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

- `app/api/generate-pptx/route.ts`: genera el `.pptx` final
- `app/api/advanced-generate/route.ts`: genera la estructura avanzada de diapositivas en JSON
- `app/api/health/route.ts`: expone el estado operativo del backend y sus dependencias
- `generate_excel_org_presentation.js`: motor principal de generacion
- `organizer.py`: analisis y organizacion del Excel

## Notas operativas

- Las rutas API aceptan archivos Excel `.xlsx`, `.xls` y `.xlsm`
- El backend aplica validaciones de tamano, sanitizacion de nombre y timeout de procesos
- Si Python o `pandas` no estan disponibles, la generacion fallara
- Puedes verificar el estado del runtime en `GET /api/health`

## Checklist de despliegue

- Confirmar que `python --version` responde correctamente en el servidor
- Confirmar que `organizer.py` y `generate_excel_org_presentation.js` existen en la raiz del proyecto desplegado
- Verificar `GET /api/health` antes de habilitar trafico real
- Probar al menos un flujo completo con un Excel pequeno antes de pasar a produccion
