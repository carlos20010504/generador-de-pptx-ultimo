# Rediseño del pipeline Excel → PPTX con Hermes como planificador único

**Fecha:** 2026-05-07
**Autor:** Claude (Sonnet 4.6) en colaboración con rashid.andrade@gmail.com
**Estado:** Diseño aprobado — pendiente revisión final del usuario antes de pasar a implementación
**Branch destino:** `feature/hermes-driven-pipeline`

---

## Contexto y motivación

El generador actual (`Socya PPTX Generator`) convierte Excels en PowerPoints usando un pipeline híbrido Next.js + Python (`organizer.py` + `generate_template_presentation.py`) con Hermes 3 405B Instruct (free) vía OpenRouter para enriquecer textos.

**Problemas reportados por el usuario en uso real:**

1. PPT con datos fantasmas (valores que no existen en el Excel).
2. Gráficas sin sentido (ejes mal elegidos, todo cero, una sola categoría dominante).
3. Tablas con celdas vacías o columnas irrelevantes.
4. Textos genéricos plantilla ("Análisis ejecutivo del archivo", "Conclusiones sobre los datos").
5. El prompt del usuario no afecta la estructura del PPT (solo entra como decoración).
6. Sensación de que la IA no se usa (en realidad: la config `best_effort + 35s deadline + 0 wait` hace que casi siempre caiga al fallback local).

**Diagnóstico raíz:** El pipeline actual rellena huecos con plantillas genéricas cuando algo falta o cuando Hermes no responde, en vez de **omitir** ese contenido. El fallback local (`build_local_executive_summary`, `build_local_briefing_payload`, etc.) es la fuente principal de los textos genéricos. Y como Hermes está semi-desactivado, ese fallback se activa casi siempre.

**Restricciones explícitas del usuario:**

- Mantener la app funcional en todo momento durante el rediseño.
- Hermes (o equivalente IA) **debe leer datos reales del Excel**, no solo metadatos.
- Cero errores de cuota visibles para el usuario en operación normal.
- Preferencia: error claro y accionable antes que PPT generado con basura.
- Soporte para cualquier tipo de Excel (no solo dominios actualmente cubiertos).

---

## Objetivo

Rediseñar el pipeline para que:

1. **Toda la planificación de slides venga de la IA**, dirigida por el prompt del usuario y los datos reales del Excel.
2. **Todo dato en el PPT sea trazable a celdas concretas** del Excel (provenance estricta). Lo no trazable se descarta.
3. **El uso de cuota de OpenRouter se minimice** mediante una sola llamada IA por presentación, caché por hash, cadena de modelos free de fallback y dos perfiles de tolerancia (preview / generación).
4. **La UI comunique honestamente** el estado de la IA, los datos descartados y los pasos de generación.

---

## Sección 1 — Arquitectura del nuevo pipeline

Pipeline en **5 pasos**, con la IA presente en **un solo punto** (paso 3):

```
[1] Parser Excel (Python, sin IA)
        ↓
[2] Inventario de datos (Python, sin IA)
        ↓
[3] Planificador IA (1 sola llamada Hermes/cadena modelos)
        ↓
[4] Validador + extractor (Python, sin IA)
        ↓
[5] Renderer PPTX (Python, sin IA, ejecuta-solo)
```

### 1.1 — Parser Excel

Carga `.xlsx`, `.xls`, `.xlsm` con `pandas` + `openpyxl`. Devuelve estructura cruda de hojas, headers, dtypes, sample rows. **No infiere nada de negocio.**

### 1.2 — Inventario de datos

Construye un catálogo enumerado de "bloques disponibles". Cada bloque tiene:
- `id` único (`T1`, `K1`, `C1`, `S1`...)
- `kind` (`table`, `kpi_candidate`, `categorical_distribution`, `time_series_candidate`)
- `provenance` (sheet, rows, columns, source_cells)
- `quality_flags` (fill_ratio, n_unique, has_nulls, all_zero, single_category, etc.)

Bloques de mala calidad se marcan aquí mismo (no se descartan todavía — esa decisión la toma la IA o el validador).

### 1.3 — Planificador IA

**Una única llamada** por presentación. Entradas:

```json
{
  "user_prompt": "...",
  "audience": "...",
  "language": "...",
  "workbook": {
    "filename": "...",
    "sheets": [
      {
        "name": "...",
        "shape": [rows, cols],
        "fill_ratio": 0.94,
        "columns": [
          { "name": "...", "dtype": "...", "n_unique": N, "fill_ratio": 0.X,
            "samples": [...], "min": ..., "max": ..., "mean": ...,
            "top_values": [["val", count], ...] }
        ],
        "first_rows": [...8 filas representativas en formato compacto...]
      }
    ],
    "blocks": [...IDs precomputados del inventario...]
  }
}
```

Salida obligatoria:

```json
{
  "presentation_meta": { "title": "...", "subtitle": "..." },
  "slides": [
    { "type": "title", "title": "...", "subtitle": "..." },
    { "type": "kpi_row", "title": "...", "block_refs": ["K1","K3","K5"] },
    { "type": "chart", "chart_type": "bar|line|pie", "title": "...",
      "block_ref": "C1", "narrative": "..." },
    { "type": "table", "title": "...", "block_ref": "T1",
      "columns_subset": [...], "max_rows": 12 },
    { "type": "text_bullets", "title": "...", "bullets": [...],
      "supports_block": "T1" }
  ],
  "prompt_suggestions": [...3-5 sugerencias específicas a este Excel...]
}
```

La IA **solo puede referenciar IDs que existen en `blocks`**. Si referencia algo inexistente, ese slide se descarta en el paso 4.

### 1.4 — Validador + extractor

Por cada slide del plan:
- Verifica que `block_ref` / `block_refs` existan en el inventario.
- Verifica que el bloque pase los `quality_flags` críticos.
- Aplica reglas de validación específicas por tipo de slide (Sección 3.2).
- Extrae los datos reales del Excel referenciados, con `provenance` adjunta.
- Filtra NaN/null/literales feos (Sección 3.1 actualizada).
- Si después de validar **quedan ≥1 slides reales**, sigue al renderer. Si 0, devuelve `EXCEL_INSUFFICIENT_DATA`.

### 1.5 — Renderer PPTX

Itera el plan validado y dibuja con `python-pptx` + `matplotlib`. **Sin fallbacks de plantilla.** Si el plan dice "tabla con bloque T1", saca exactamente esos datos. La plantilla corporativa (`Plantilla_Presentacion_Socya (1) (1).pptx`) sigue usándose para layouts/colores, pero no como fuente de contenido.

### 1.6 — Reducción de código existente

- `organizer.py`: se reduce a Pasos 1-2 + Paso 4 (~30% del código actual). Se eliminan heurísticas domain-specific (auditoría, comisiones, etc.) y todos los textos plantilla.
- `generate_template_presentation.py`: se reduce a Paso 5 (renderer puro).
- Aparece un nuevo módulo `planner.py` que hace la llamada IA + cadena de fallback.

---

## Sección 2 — Estrategia anti-cuota + lectura real del Excel

### 2.1 — Una sola llamada por presentación

El planificador único (1.3) reemplaza las 3-5 llamadas actuales (`enrich_textual_blocks_with_ai`, `build_executive_summary_ai`, `build_briefing_ai`, `generate_suggestions_ai`...). **Reducción 5× de consumo de cuota por Excel.**

### 2.2 — Payload destilado (~4-6K tokens)

Hermes recibe muestras reales suficientes para "ver" cada columna sin recibir el Excel completo:
- Por columna: dtype, n_unique, fill_ratio, samples (5-8 valores), min/max/mean/sum si numérica, top_values si categórica.
- Por hoja: 8 primeras filas representativas en formato compacto.
- **Truncamiento adaptativo:** si el Excel tiene >5 hojas o columnas con >50 valores únicos, se envían top-N por frecuencia + summary stats. Target: payload ≤6K tokens siempre.

### 2.3 — Caché persistente por hash de contenido

`cache_key = sha256(file_bytes + user_prompt + audience + language + planner_version)`.

- Hit en caché → 0 llamadas IA. Plan instantáneo.
- Almacén: `.cache/plans/{hash}.json`.
- Nunca expira: el contenido del Excel no cambia para el mismo hash.
- `planner_version` se bumpea cuando cambia el prompt del planner → invalida cache automáticamente.

### 2.4 — Cadena de fallback de modelos free

```python
MODEL_CHAIN = [
    "nousresearch/hermes-3-llama-3.1-405b:free",
    "meta-llama/llama-3.3-70b-instruct:free",
    "google/gemma-2-9b-it:free",
    "qwen/qwen-2.5-72b-instruct:free",
]
```

Cada modelo tiene cuota independiente en OpenRouter. Si el primario devuelve 429 / "high demand" / timeout, se intenta el siguiente automáticamente. La UI muestra qué modelo terminó respondiendo.

### 2.5 — Dos perfiles de tolerancia

| Momento | Perfil | Comportamiento |
|---|---|---|
| Subir Excel → preview en panel IA | `fast` | Timeout 25s, 1 modelo. Si falla, panel muestra resumen estructural sin IA + flag "Vista previa sin IA". No bloquea. |
| Click "Generar PPTX" | `patient` | Hasta 4 min, 4 ciclos por la cadena, respetando `Retry-After`. UI muestra progreso real. |

### 2.6 — Si todos los modelos saturan (Opción A)

Devolver error claro:

```
"No hay modelos IA disponibles ahora mismo (todos los gratuitos están saturados).
Espera 5-10 min e intenta de nuevo. Tu Excel está validado y listo."
```

UI muestra botón "Reintentar generación" sin requerir re-upload. **No se genera PPT mínimo de relleno.**

---

## Sección 3 — Fidelidad de datos y procedencia estricta

**Regla de oro:** Si un dato no se puede trazar a celdas concretas del Excel, no aparece en el PPT.

### 3.1 — Procedencia obligatoria por slide

Cada slide validado lleva campo `provenance`:

```json
{
  "type": "chart",
  "title": "Ventas por ciudad",
  "block_ref": "C1",
  "provenance": {
    "sheet": "Ventas",
    "rows": [2, 121],
    "columns": ["Ciudad", "Total"],
    "cell_count": 240,
    "non_empty_count": 238,
    "data_hash": "sha256:abc...",
    "extracted_at": "2026-05-07T..."
  }
}
```

Esto se persiste en el `.audit.json` para trazabilidad.

### 3.2 — Reglas de validación duras

**Charts** se descartan si:
- < 2 categorías con valor > 0.
- Todos los valores iguales (sin distribución).
- > 30% NaN/vacíos.
- Pie con > 6 categorías o con una > 85%.
- Eje X es ID secuencial o campo único por fila.

**Tablas** se filtran:
- Columna con `fill_ratio < 0.3` → fuera.
- Fila con `fill_ratio < 0.5` → fuera.
- < 2 filas o < 2 columnas restantes → tabla entera fuera.
- Headers `Unnamed: N` o vacíos → columna fuera.
- Literales feos (`nan`, `None`, `null`, `NaT`, `???`, `—`, `s/d`, `n/a`) → reemplazados por string vacío. Si fila queda con < 50% celdas llenas tras la sustitución → fila fuera.

**KPIs** se descartan si:
- Valor 0 sin contexto explícito.
- Placeholder string.
- Cálculo ambiguo (mixto texto/número).

**Bullets / textos narrativos:**
- Se extraen automáticamente todos los números, fechas, nombres propios y porcentajes con regex.
- Cada extraído se busca en el bloque `supports_block` del Excel (tolerancia ±0.5% numérico, fuzzy match con `rapidfuzz` para nombres).
- Si NO existe → bullet descartado entero (no se "corrige").
- Bullets puramente cualitativos sin números/nombres específicos → también descartados (son texto genérico).
- Si quedan 0 bullets válidos → slide entero descartado.

### 3.3 — Eliminación total de plantillas genéricas

Lista de funciones / textos que se eliminan:

| Origen actual | Acción |
|---|---|
| `build_local_executive_summary` | Eliminar |
| `build_local_briefing_payload` | Eliminar |
| Conclusiones plantilla ("Análisis ejecutivo del archivo", "Los datos muestran patrones relevantes", "Se recomienda profundizar el análisis") | Eliminar |
| `buildPromptFocusPhrase` con fallback "datos reales del Excel" | Eliminar |
| Slide cierre forzado "Siguiente decisión sobre..." | Eliminar (solo aparece si IA lo pide) |
| KPIs metainfo (Hoja base / Filas / Columnas / Columnas numéricas) | Eliminar |
| Títulos receta tipo "Detalle adicional para..." | Eliminar |

Reduce ~1,500 líneas de `organizer.py`.

### 3.4 — Umbral mínimo

≥1 slide validado → genera PPT. 0 slides → `EXCEL_INSUFFICIENT_DATA` con mensaje accionable.

### 3.5 — Audit.json mejorado

```json
{
  "model_used": "nousresearch/hermes-3-llama-3.1-405b:free",
  "cache_hit": false,
  "fallback_chain_steps": [],
  "slides_planned": 9,
  "slides_validated": 7,
  "slides_dropped": [
    { "type": "chart", "reason": "all_zero_values", "block_ref": "C2" },
    { "type": "text_bullets", "reason": "all_bullets_failed_provenance", "supports_block": "T3" }
  ],
  "bullets_dropped": 4,
  "provenance_per_slide": [...]
}
```

---

## Sección 4 — Cambios en UI

### 4.1 — Indicador dinámico de modelo IA

En header del `AIControlPanel.tsx` (reemplaza string estático "Hermes 3 · OpenRouter"):

```
✦ Hermes 3 405B · listo
✦ Hermes saturado → usando Llama 3.3 70B
✦ Resultado en caché (sin consumo IA)
✦ Vista previa sin IA — Hermes saturado, reintentar
✦ Esperando IA disponible: 1m 23s restantes…
✦ ⚠ Todos los modelos saturados — espera 5-10 min
```

Estado viene de `ai_status: { model, source, wait_seconds, message }` en cada respuesta.

### 4.2 — Progreso real (5 fases) durante generación

`/api/generate-pptx` se reescribe a **streaming SSE** emitiendo eventos `{phase, step, message}`:

```
📊 [1/5] Leyendo Excel…
🗂 [2/5] Construyendo inventario…
🧠 [3/5] Consultando IA (Hermes)…       (con countdown si espera)
✓ [4/5] Validando datos…
🎨 [5/5] Renderizando PPTX…
```

### 4.3 — Banner de calidad post-generación

```
✓ Presentación generada con Hermes 3 · 7 slides validados · 2 slides descartados (datos insuficientes)
   [Ver detalles]
```

Modal "Ver detalles" muestra audit en lenguaje natural (no JSON crudo): modelo usado, slides descartados con razón legible, bullets descartados.

### 4.4 — Sugerencias de prompt dinámicas

`PLACEHOLDER_SUGGESTIONS` hardcoded → fuera. Las sugerencias vienen 100% del planner IA en su única llamada (`prompt_suggestions: [...]`). Si cache hit → vienen del cache.

### 4.5 — Botón "Reintentar IA" tras saturación

UI muestra error claro de Sección 2.6 + botones [Reintentar] [Cambiar prompt] [Cancelar]. Excel y prompt persisten en memoria del frontend (sin re-upload).

### 4.6 — UX intuitiva

- Iconos amigables por fase (no solo `[N/5]`).
- Tooltips en hover sobre términos técnicos.
- Errores en español llano sin jerga ("OpenRouter saturado" → "Servicio IA con mucha demanda").
- Botón "Generar PPTX" siempre visible y dominante.
- Estados vacíos explicados.
- Modal audit en lenguaje natural.

### 4.7 — Lo que NO cambia

Layout general del panel, focus selector, theme picker, paleta visual, plantilla corporativa.

---

## Sección 5 — Manejo de errores y rollback

### 5.1 — Contrato uniforme de error

Las 3 rutas API responden con shape estándar:

```json
{
  "error": {
    "code": "AI_SATURATED",
    "message": "...",
    "details": "...",
    "user_action": "retry_later",
    "retry_after_seconds": 300
  }
}
```

Códigos cerrados: `EXCEL_INVALID`, `EXCEL_EMPTY`, `EXCEL_INSUFFICIENT_DATA`, `AI_SATURATED`, `AI_RESPONSE_INVALID`, `PLANNER_REJECTED_PROMPT`, `PYTHON_RUNTIME_ERROR`, `TIMEOUT`.

### 5.2 — Comportamiento por tipo de fallo

- **Parser Python crash** → `EXCEL_INVALID` + log server.
- **Inventario vacío** → `EXCEL_EMPTY`.
- **IA JSON inválido** → reintento con `temperature=0`, luego siguiente modelo, luego `AI_RESPONSE_INVALID`.
- **Plan referencia bloques inexistentes** → validador descarta esos slides; si quedan ≥1 sigue, si 0 → `EXCEL_INSUFFICIENT_DATA`.
- **Renderer falla** → `PYTHON_RUNTIME_ERROR` + log con plan completo.
- **Plantilla corporativa ausente** → `/api/health` lo detecta; generación falla con mensaje claro.

### 5.3 — Garantías

1. Nunca PPT con datos no trazables.
2. Nunca cuelgue > 28 min.
3. Nunca silenciar fallback (siempre se reporta en `ai_status`).

### 5.4 — Migración por fases (cero downtime)

| Fase | Entrega | Estado del repo |
|---|---|---|
| **F1** | Capa anti-cuota (cache hash + cadena modelos + perfiles fast/patient) | Hermes empieza a usarse de verdad. PPTs siguen iguales. |
| **F2** | Nuevo `planner.py` + extractor con provenance. Llamado solo desde `/api/advanced-generate`. Pipeline viejo en paralelo. | PPT con plan IA. Si falla, fallback al pipeline viejo (interim). |
| **F3** | Renderer execute-only. Validador estricto. Drop fallbacks plantilla. | Pipeline nuevo único. Pipeline viejo eliminado. |
| **F4** | UI streaming SSE + badges + modal audit. | UX completo. |
| **F5** | Limpieza de código muerto en `organizer.py` y `generate_template_presentation.py`. | Repo limpio. |

Cada fase es deployable. `git revert` por fase si rompe.

### 5.5 — Setup previo a F1

1. Branch `feature/hermes-driven-pipeline` desde `master`.
2. **Commit baseline**: los ~2000 líneas de cambios pendientes en `master` (en la misma línea de los últimos commits "strengthen hermes ppt pipeline" y "improve data fidelity") se commitean primero como WIP estado base.
3. Snapshot del comportamiento actual con 3-4 Excels de prueba (pptx + audit guardados como referencia).
4. Test suite mínima: cargar Excel → `/api/health` OK → genera PPT → tiene >0 slides.

### 5.6 — Riesgos y mitigaciones

| Riesgo | Probabilidad | Mitigación |
|---|---|---|
| Hermes free retirado de OpenRouter | Baja | Cadena 4 modelos cubre |
| Free tier requiere tarjeta | Media | Diseño no asume API key, falla con `AI_SATURATED` claro |
| Excels >50MB / >100K filas | Baja | Inventario detecta, IA recibe muestra estadística, renderer pagina |
| `python-pptx` versión rompe layouts | Baja | Lock en `requirements.txt`, tests regresivos |
| Usuario percibe lentitud | Media | UI streaming + cache hit instantáneo en re-runs |

---

## Sección 6 — Plan de pruebas

### 6.1 — Fixtures (`__tests__/fixtures/`)

| # | Archivo | Caso |
|---|---|---|
| 1 | `ventas_simple.xlsx` | 1 hoja, 100 filas, 5 cols mixed |
| 2 | `comisiones_real.xlsx` | Caso real existente |
| 3 | `auditoria_compleja.xlsx` | Multi-hoja: hallazgos + COSO + checklist |
| 4 | `casi_vacio.xlsx` | 3 filas + mucho NaN |
| 5 | `enorme.xlsx` | 50K filas, 20 cols |
| 6 | `corrupto.xlsx` | Archivo dañado |
| 7 | `dominio_raro.xlsx` | RRHH (no auditoría/ventas) |

### 6.2 — Tests automáticos

**Python (`pytest`):**
- `test_parser.py` — los 7 fixtures parsean sin excepción.
- `test_inventory.py` — fixture #1 produce bloques esperados con provenance.
- `test_validator.py` — `block_ref` inexistente se descarta.
- `test_provenance.py` — bullet con número inexistente se filtra.
- `test_renderer.py` — plan validado produce `.pptx` válido.

**Node (`jest`):**
- `test_ai_chain.test.js` — mock 429 en Hermes → cae a Llama → success.
- `test_cache.test.js` — segunda llamada con mismo `cache_key` no toca red.
- `test_error_contract.test.js` — cada `code` retorna shape estándar.
- `test_sse_stream.test.js` — `/api/generate-pptx` emite eventos en orden.

**E2E manual (`npm run smoke`):** levanta dev server, sube fixture #1, valida pptx + audit.

### 6.3 — Checklist manual pre-merge

- [ ] Fixture #1 → PPT con 5+ slides, abre en PowerPoint, revisión visual OK.
- [ ] Fixture #2 con prompt "solo gráficas, máx 6 slides" → plan respeta prompt.
- [ ] Fixture #4 → mensaje de error claro, no PPT basura.
- [ ] Fixture #1 dos veces → segunda instantánea (cache hit).
- [ ] `SOCYA_FORCE_AI_ERROR=429` → UI muestra fallback a Llama.
- [ ] Ningún PPT muestra `nan`, `None`, `null`, `???` en celdas.
- [ ] Audit modal legible en español.

### 6.4 — Criterio "listo"

1. 7 fixtures pasan tests automáticos.
2. Checklist 6.3 pasa 100%.
3. Inspección manual del usuario en 3 PPTs confirma desaparición de problemas reportados.
4. `npm run lint` y `npm test` verdes.
5. `/api/health` reporta capacidades OK.

Tope: si tras 3 iteraciones de feedback del usuario no aprueba calidad → rollback completo al pipeline anterior.

---

## Decisiones tomadas (alineación usuario ↔ Claude)

| Decisión | Elegido |
|---|---|
| Enfoque (1/2/3) | **Enfoque 1** — Plan dirigido por IA + procedencia estricta |
| Soporte de dominios | Cualquier Excel (genérico, sin priors de dominio) |
| Si todos los modelos saturan | **Opción A** — error claro con botón "Reintentar" |
| Truncamiento payload (~6K tokens) | Aprobado |
| Cadena de 4 modelos free | Aprobado |
| Umbral mínimo de slides | ≥1 slide real (lenient) + cero literales feos en celdas |
| Streaming progreso | **SSE** sobre Next.js 16 |
| Cambios pendientes en `master` | **Opción A** — commit baseline antes de empezar |

---

## Próximo paso

Pasar a `superpowers:writing-plans` para producir el plan de implementación detallado fase por fase, con checkpoints de revisión humana entre F1 → F2 → F3 → F4 → F5.
