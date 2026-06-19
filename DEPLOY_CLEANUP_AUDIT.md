# Deploy Cleanup Audit

## Resumen

Esta limpieza se hizo para preparar el proyecto para produccion sin tocar el flujo activo:

- UI web en `app/` y `components/`
- APIs en `app/api/advanced-generate`, `app/api/generate-pptx` y `app/api/health`
- Backend Python productivo en `organizer.py` y `generate_template_presentation.py`
- Plantilla activa `Plantilla_Presentacion_Socya (1) (1).pptx`

## Backup previo

- Backup completo creado antes de eliminar archivos:
  - `C:\Users\cpinzon\Desktop\generador-pptx-socya-main-ultimo-backup-20260423-095747`

## Criterio de eliminacion segura

Cada elemento eliminado cumplio al menos una de estas condiciones:

- No tenia referencias en el codigo activo, APIs, UI ni configuracion de despliegue.
- Era un artefacto generado por pruebas, revisiones visuales o ejecuciones manuales.
- Era un asset por defecto o un recurso estatico no referenciado.
- Era una pila legacy aislada del flujo actual de produccion.

## Ajustes de codigo y configuracion

- `package.json`
  - Se elimino la dependencia `pptxgenjs`, que quedo sin importaciones reales tras retirar la pila legacy.
  - Se actualizo `next` y `eslint-config-next` a `16.2.4` para corregir una vulnerabilidad con fix disponible.
- `package-lock.json`
  - Se sincronizo el lockfile con las versiones finales instaladas para evitar inconsistencias en despliegue.
- `utils/presentation-integrity.cjs`
  - Se actualizo la referencia permitida desde el generador legacy JS hacia `generate_template_presentation.py`.
- `__tests__/safe-mode.test.js`
  - Se reemplazaron nombres de archivos de muestra por nombres genericos no dependientes de archivos reales en el repo.
- `next.config.ts`
  - Se elimino `utils/presentation-integrity.cjs` del tracing de `/api/generate-pptx` porque no participa en runtime productivo.
- `app/layout.tsx`
  - Se reemplazo metadata obsoleta de Create Next App por metadata real del producto.
- `README.md`
  - Se corrigio la documentacion del flujo real de produccion.
- `GEMINI.md`
  - Se actualizo la arquitectura documentada al generador Python actual.
- `README_ADVANCED_PPTX.md`
  - Se reescribio para describir el flujo vigente.
- `Guia_Organizacion_Excel.md`
  - Se quitaron referencias a nombres de archivos de ejemplo eliminados.
- `eslint.config.mjs`
  - Se retiraron referencias a archivos legacy ya eliminados.
- `tsconfig.json`
  - Se amplio el include para aceptar archivos `.d.ts` auxiliares usados en la compatibilidad de tipos.
- `next-module-shims.d.ts`
  - Se agrego un shim de tipos para compatibilidad con imports `.js` generados por `Next 16.2.4` durante el build.

## Archivos de codigo eliminados

### Stack legacy y utilidades fuera de flujo

- `analyze_template.py`
- `build_template_presentation.py`
- `check_analysis.py`
- `convert.js`
- `create_large_multi_sheet_excel.py`
- `create_test_excel.py`
- `extract_colors.py`
- `generate_excel_org_presentation.js`
- `inspect_slide2.py`
- `inspect_slide4.py`
- `inspect_slide5.py`
- `inspect_slide_2.py`
- `inspect_slide_4.py`
- `inspect_slide_5.py`
- `read_template.py`
- `template.zip`
- `test-parser-v2.ts`
- `test_clean_slate.py`
- `test_insert.py`
- `test_organizer.py`
- `test_render.ts`
- `validate-pipeline.js`
- `validate-quick.js`

### Tests y modulos legacy no usados por el flujo actual

- `__tests__/advanced-pptx-generator.test.ts`
- `utils/advanced-pptx-generator.ts`
- `utils/excel-parser.ts`
- `utils/pptx-renderer.ts`
- `utils/socya-renderer.ts`

## Recursos estaticos eliminados

- `public/file.svg`
- `public/globe.svg`
- `public/next.svg`
- `public/vercel.svg`
- `public/window.svg`

## Directorios de artefactos eliminados

- `ppt_preview/`
- `ppt_preview_v2/`
- `ppt_review_antimicro_com/`
- `ppt_review_antimicro_com2/`
- `ppt_review_antimicro_fer/`
- `ppt_review_antimicro_fer2/`
- `ppt_review_antimicro_final_com/`
- `ppt_review_antimicro_final_fer/`
- `ppt_review_antimicro_mapcheck/`
- `ppt_review_comisiones/`
- `ppt_review_comisiones_final/`
- `ppt_review_directivo/`
- `ppt_review_ferreteria/`
- `ppt_review_ferreteria_final/`
- `ppt_review_ferreteria_run1/`
- `ppt_review_ferreteria_run2/`
- `ppt_review_final_comm/`
- `ppt_review_final_universal/`
- `ppt_review_fix_cards_ferreteria/`
- `ppt_review_fix_cards_stacked/`
- `ppt_review_fix_legibility_com/`
- `ppt_review_fix_legibility_com2/`
- `ppt_review_fix_legibility_com3/`
- `ppt_review_fix_legibility_fer/`
- `ppt_review_latest_universal/`
- `ppt_review_layout_comisiones/`
- `ppt_review_layout_ferreteria/`
- `ppt_review_layout_multiblock/`
- `ppt_review_metric_fix_final/`
- `ppt_review_multiblock/`
- `ppt_review_naturaleza/`
- `ppt_review_nocut_comm2/`
- `ppt_review_nocut_uni2/`
- `ppt_review_prod_comisiones/`
- `ppt_review_prod_comisiones2/`
- `ppt_review_prod_ferreteria1/`
- `ppt_review_prod_ferreteria2/`
- `ppt_review_prod_ferreteria3/`
- `ppt_review_softgreen/`
- `ppt_review_tablas_full/`
- `ppt_review_tables_comm/`
- `ppt_review_tables_universal/`
- `ppt_review_textheavy/`
- `ppt_review_v1/`
- `ppt_review_verde/`
- `template_unpacked/`
- `__pycache__/`

## Artefactos y datos de prueba eliminados

### Excels y JSON temporales

- `.socya-layout-history.json`
- `AUDITORIA VIVA.xlsx`
- `auditoria_viva_organizer.json`
- `Comisiones V1.xlsx`
- `ferreteria_demo.xlsx`
- `Muestra final 14 oct 2025 (Autoguardado).xlsx`
- `muestra_final_organizer.json`
- `multi_block_demo.xlsx`
- `solo_texto_demo.xlsx`
- `stacked_complex_demo.xlsx`
- `test_data.xlsx`
- `text_heavy_demo.xlsx`
- `tmp_com_analysis.json`
- `tmp_organizer_output.json`

### Plantillas duplicadas, salidas y revisiones manuales

- `api_test_Comisiones_V1.pptx`
- `Mejora.pptx`
- `Plantilla_Presentacion_Socya (1) (1) - copia.pptx`
- `test_insert.pptx`
- `test_output.audit.json`
- `test_output.pptx`
- `test_output_bi.audit.json`
- `test_output_bi.pptx`
- `test_output_bi_v2.audit.json`
- `test_output_bi_v2.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v10b.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v10b.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v5.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v5.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v5b.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v5b.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v5c.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v5c.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v6.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v6.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v7.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v7.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v8.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v8.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v8b.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v8b.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v9.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v9.pptx`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v9b.audit.json`
- `Presentacion_Plantilla_AUDITORIA_VIVA_boardroom_v9b.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v11.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v11.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v12.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v12.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v13.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v13.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v14.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v14.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v15.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v15.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v16.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v16.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v17.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v17.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v18.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v18.pptx`
- `Presentacion_Plantilla_Comisiones_V1_v19.audit.json`
- `Presentacion_Plantilla_Comisiones_V1_v19.pptx`
- `Presentacion_Plantilla_ferreteria_demo_v10.audit.json`
- `Presentacion_Plantilla_ferreteria_demo_v10.pptx`
- `Presentacion_Plantilla_ferreteria_demo_v5.audit.json`
- `Presentacion_Plantilla_ferreteria_demo_v5.pptx`
- `Presentacion_Plantilla_ferreteria_demo_v6.audit.json`
- `Presentacion_Plantilla_ferreteria_demo_v6.pptx`
- `Presentacion_Plantilla_ferreteria_demo_v7.audit.json`
- `Presentacion_Plantilla_ferreteria_demo_v7.pptx`
- `Presentacion_Plantilla_ferreteria_demo_v8.audit.json`
- `Presentacion_Plantilla_ferreteria_demo_v8.pptx`
- `Presentacion_Plantilla_ferreteria_demo_v8b.audit.json`
- `Presentacion_Plantilla_ferreteria_demo_v8b.pptx`
- `Presentacion_Plantilla_ferreteria_demo_v9b.audit.json`
- `Presentacion_Plantilla_ferreteria_demo_v9b.pptx`

### Artefactos de texto y exportaciones auxiliares

- `colors_out.txt`
- `ejemplo_plantilla.csv`
- `errors.txt`
- `Exceltopptxstructure.txt`
- `javascriptsx.txt`
- `plantilla.txt`
- `sample_tags.csv`
- `test_output.txt`
- `test_output_content.md`
- `test_output_content_v2.md`
- `tmp_excel_content.txt`
- `tmp_pptx_content.md`
- `tmp_pptx_content_utf8.md`
- `validate-output.txt`
- `validate-result.txt`

## Validaciones ejecutadas

- `npm run lint`
- `npm test`
- `npm run build`
- `npm audit --omit=dev --json`

`lint`, `test` y `build` terminaron correctamente despues de la limpieza.

## Dependencias retenidas

No se eliminaron dependencias de `package.json` sin evidencia fuerte de no uso en runtime o cadena de build. En particular:

- `next`, `react`, `react-dom`, `lucide-react`, `xlsx`
- `tailwindcss` y `@tailwindcss/postcss`
- `babel-plugin-react-compiler`

Estas dependencias permanecen porque forman parte del flujo web activo o de la cadena de compilacion validada por `build`.

## Hallazgos de seguridad residuales

- `next`
  - Vulnerabilidad alta reportada por `npm audit` corregida al subir a `16.2.4`.
- `xlsx`
  - `npm audit` sigue reportando 1 vulnerabilidad alta sin `fixAvailable` oficial desde npm.
  - Riesgo residual pendiente de mitigacion aguas arriba o de sustitucion futura de libreria.
