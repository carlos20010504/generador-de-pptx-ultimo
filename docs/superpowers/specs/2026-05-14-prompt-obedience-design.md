# Prompt Obedience + UX Cleanup — Design

**Date:** 2026-05-14
**Status:** Draft → approved (verbal)
**Author:** rashid.andrade@gmail.com (con Claude)
**Branch base:** `feature/hermes-driven-pipeline`

## Problema

El usuario reporta cuatro síntomas concretos cuando genera un PowerPoint desde el Excel de comisiones:

1. **El AI ignora el prompt.** "Hazme 9 slides con Riesgos Core" → devuelve 7 sin Riesgos Core.
2. **Slides pobres.** Las slides `text_bullets` salen con 2-3 bullets cortos y se sienten vacías.
3. **"Analizando estructura del Excel..." se queda colgado** en algunos archivos.
4. **Dos lugares para el prompt** (PreparePanel "Refinar prompt y contexto" + AIControlPanel drawer "Panel avanzado"). El usuario se pierde y no sabe dónde escribir.

Además quiere poder eliminar slides en la vista previa **antes** de generar el PPTX (esto ya existe en `PreparePanel`, pero no es descubrible).

## Objetivo

Hacer que el deck generado obedezca lo que el usuario pide en el prompt — cantidad de slides exacta, hojas mencionadas siempre incluidas — y que el flujo UI tenga **un solo lugar obvio** donde escribir esa instrucción.

No-objetivos:
- Rediseñar el pipeline Hermes (es otro proyecto, ya hay spec aparte).
- Cambiar el sistema de validación de provenance.
- Agregar tipos nuevos de slide.

## Decisiones tomadas

| Decisión | Opción elegida | Por qué |
|---|---|---|
| Cantidad de slides | **Estricta**: pides 9 → exactamente 9 | El usuario controla el tamaño del deck con precisión |
| Hojas mencionadas | **Garantía**: mención de hoja → mínimo 1 slide de esa hoja | Predecible; si falta, warning explícito |
| Densidad bullets | 5-7 por slide `text_bullets`, cada uno con dato concreto | Responde "se ven pobres con solo 3 líneas" |
| Lugar del prompt | **Solo PreparePanel**, abierto por default | Eliminar duplicación que confunde |

## Arquitectura

### Capas afectadas

```
┌────────────────────────────────────┐
│ UI: PreparePanel.tsx               │ ← prompt único + intent banner
│     AIControlPanel.tsx             │ ← se elimina textarea
└────────────────┬───────────────────┘
                 │ /api/preview-plan
                 ▼
┌────────────────────────────────────┐
│ socya_pipeline/prompt_intent.py    │ ← NUEVO: parsea slide_count + sheets
│ socya_pipeline/planner.py          │ ← enforcement post-AI + bullets
│ socya_pipeline/ai_chain.py         │ ← unchanged
└────────────────────────────────────┘
```

### Componente nuevo: `prompt_intent.py`

Módulo puro (sin I/O), una función pública:

```python
def extract(user_prompt: str, available_sheet_names: list[str]) -> PromptIntent:
    """Parsea la intención del usuario y devuelve estructura usable por el planner.

    Returns
    -------
    PromptIntent con:
      - requested_slide_count: int | None  (regex: "9 slides", "deck de 12")
      - required_sheets: list[SheetMatch]  (fuzzy contra available_sheet_names)
      - topic_hints: list[str]             (keywords sueltas, lowercase)
    """
```

Donde `SheetMatch` es:

```python
@dataclass
class SheetMatch:
    requested: str          # texto literal del prompt: "riesgos core"
    matched: str | None     # nombre real de la hoja: "Riesgos Core" o None
    closest: str | None     # si no matchea, la hoja más parecida
    ratio: float            # confianza del fuzzy match
```

**Regex slide count** (orden importa, primer hit gana):
- `(\d+)\s*(?:slides?|diapositivas?|láminas?)`
- `deck\s*de\s*(\d+)`
- `(\d+)\s*(?:slides?|diapositivas?)\s*(?:exactas?|exactamente)`

**Fuzzy match hojas**: `difflib.SequenceMatcher.ratio()` ≥ 0.65, normalizando ambos lados (lowercase, sin tildes, sin guiones).

### Cambios en `planner.py`

#### 1. Construcción del payload

`build_payload()` ahora recibe el `PromptIntent` y lo inyecta en el JSON que va a la AI:

```python
payload = {
    "user_prompt": effective_prompt,
    "user_intent": {
        "slide_count": intent.requested_slide_count,   # null si no se pidió
        "required_sheets": [m.matched for m in intent.required_sheets if m.matched],
    },
    ...
}
```

#### 2. PROMPT_TEMPLATE — reglas nuevas

Se agregan dos reglas al template, justo antes de la regla de "Volumen":

> **0a. Cantidad de slides**: si `user_intent.slide_count` está presente y no es null, debes producir EXACTAMENTE ese número de slides totales (incluyendo portada). NO menos, NO más.
>
> **0b. Hojas obligatorias**: si `user_intent.required_sheets` no está vacío, cada hoja listada debe aparecer como `provenance.sheet` en al menos UNA slide. Si no podés cubrir alguna por falta de datos válidos, OMÍTELA y al final del JSON añade `"_skipped_required_sheets": ["nombre"]`.

La regla "Volumen" actual ("entre 7 y 11 slides") se vuelve fallback solo cuando `slide_count` es null.

#### 3. Enforcement post-AI

Después de parsear el JSON del AI:

```python
def _enforce_intent(plan: dict, intent: PromptIntent, blocks, wb) -> dict:
    """Garantiza que el plan honre el intent. Idempotente."""

    # 1. Slide count
    if intent.requested_slide_count is not None:
        target = intent.requested_slide_count
        actual = len(plan["slides"])
        if actual > target:
            # Recortar las slides menos prioritarias (mantener portada,
            # required-sheet slides, y bloques con quality_flags vacíos primero)
            plan["slides"] = _trim_to(plan["slides"], target, intent)
        elif actual < target:
            # Inyectar slides extras desde bloques no usados (chart > table > bullets)
            plan["slides"] = _pad_to(plan["slides"], target, blocks, wb, intent)

    # 2. Required sheets
    used_sheets = {s.get("provenance", {}).get("sheet")
                   for s in plan["slides"] if "provenance" in s}
    skipped = []
    for match in intent.required_sheets:
        if match.matched and match.matched not in used_sheets:
            injected = _inject_sheet_slide(match.matched, blocks, wb)
            if injected:
                plan["slides"].append(injected)
                # Si rompió el target, recortar otra slide menos prioritaria
                if intent.requested_slide_count is not None:
                    plan["slides"] = _trim_to(plan["slides"],
                                               intent.requested_slide_count, intent)
            else:
                skipped.append(match.requested)

    plan["_intent_report"] = _build_intent_report(intent, plan, skipped)
    return plan
```

#### 4. Densidad de bullets

`_build_fallback_bullets()` cap sube de 3 a 6. Se agregan dos generadores nuevos:
- `min` de la columna monetaria principal: "El monto mínimo registrado fue $X"
- Top 2 categorías de la primer cat column: "'X' representa N (Y%)"

En `PROMPT_TEMPLATE`, regla 4 cambia de "bullets específicos" a:

> **4. Bullets específicos y abundantes**: cada slide `text_bullets` debe tener **5-7 bullets**, cada uno con al menos una cifra concreta o nombre propio del bloque. PROHIBIDO bullets genéricos.

### Cambios en API: `/api/preview-plan/route.ts`

Response shape extendido:

```ts
{
  meta: { title, subtitle },
  slides: [...],
  ai_status: { model, cache_hit },
  intent_report: {
    requested_slide_count: number | null,
    actual_slide_count: number,
    count_honored: boolean,
    required_sheets: Array<{
      requested: string,
      matched: string | null,
      closest: string | null,
      slide_indices: number[],
    }>,
  } | null  // null si no se pudo parsear intent
}
```

### Cambios UI

#### `PreparePanel.tsx`

1. **Sección "Refinar prompt" se reordena al TOP** (encima del listado de slides).
2. **Abierto por default**, no colapsado.
3. Label cambia de "Refinar prompt y contexto" → **"Tu instrucción para la IA"** con tipografía más grande.
4. Placeholder cambia a algo accionable: `"ej: hazme 9 slides incluyendo Riesgos Core y Riesgos Acciones"`.
5. **Banner intent_report** sobre el listado:
   - Verde si todo honored: `✓ Detecté: 9 slides + Riesgos Core (slide 5) + Riesgos Acciones (slide 7)`
   - Amarillo con detalles si algo no se pudo cumplir: `⚠ "Riesgos Acciones" no se encontró — la hoja más cercana es 'Acciones'. Reintenta con ese nombre.`
6. Tip explícito sobre toggles de slides: `"Click en cualquier slide para activarla/desactivarla. Arrastra para reordenar. Click en el título para editarlo."` — más visible.

#### `AIControlPanel.tsx`

1. **Eliminar el textarea del prompt** (líneas ~425-520) y el botón "Aplicar".
2. **Eliminar las sugerencias de prompt** (las que ya están duplicadas en PreparePanel).
3. Mantener: audiencia, idioma, tema, focus.
4. El botón que abre el drawer en `ExcelUploader.tsx:565` se renombra de "Panel avanzado" → **"Audiencia y tema"**.

#### `ExcelUploader.tsx`

- El estado `userPrompt` (línea 95) sigue siendo el source-of-truth. No cambia.
- El drawer ya no edita ese estado.

### "Analizando estructura del Excel..." stuck

**Investigación primero, fix después.** Hipótesis a verificar:

1. Cold-start Python lento → el pre-warm de `ExcelUploader` ya dispara `?warmup=1` pero quizás solo a una ruta. Verificar que llegue a `/api/quick-summary?warmup=1` también.
2. Archivos grandes → quick-summary no tiene streaming. Si el archivo es grande, el cliente solo ve "Analizando..." sin progreso. Plan B: añadir progreso por hoja vía un endpoint SSE separado, **solo si el bug persiste después del warmup fix**.
3. Timeout backend < timeout cliente → ya están alineados (60s + 70s buffer). No es eso.

Decisión: empezar con el fix del warmup (1 línea probablemente). Si en testing el bug sigue, escalamos a SSE de progreso.

## Plan de validación (review CEO)

Después de implementar, en el Excel de comisiones reales:

**Test 1 — Cantidad estricta:**
Prompt: `"Hazme un deck de 9 slides para el comité ejecutivo."`
Esperado: exactamente 9 slides, sin warnings de intent.

**Test 2 — Hoja específica:**
Prompt: `"Incluye toda la hoja de Riesgos Core."`
Esperado: ≥1 slide con `provenance.sheet = "Riesgos Core"`, banner verde.

**Test 3 — Combinado:**
Prompt: `"Hazme 11 slides con Riesgos Core, Riesgos Acciones, y los 3 KPIs principales."`
Esperado: 11 slides, ambas hojas representadas, banner verde con ambos hits.

**Test 4 — Hoja inexistente:**
Prompt: `"Incluye la hoja de Proyecciones 2027."`
Esperado: banner amarillo `⚠ "Proyecciones 2027" no se encontró`, plan se genera igual con el resto.

**Test 5 — Densidad:**
Cualquier deck generado: cada slide `text_bullets` tiene ≥5 bullets, cada bullet tiene un número o nombre propio.

**Test 6 — Render CEO:**
Render PNG de cada slide del Test 3. Reviso uno por uno como si fuera CEO de empresa: ¿se entiende sin contexto?, ¿hay un dato concreto?, ¿el insight es accionable?

Cualquier slide que falle en ese review → identificar la causa (prompt template, extractor, renderer) y arreglarla antes de marcar el proyecto completo.

## Archivos tocados

| Archivo | Tipo |
|---|---|
| `socya_pipeline/prompt_intent.py` | NUEVO |
| `socya_pipeline/planner.py` | edit |
| `app/api/preview-plan/route.ts` | edit (response shape) |
| `components/PreparePanel.tsx` | edit (UX + banner) |
| `components/AIControlPanel.tsx` | edit (eliminar textarea) |
| `components/ExcelUploader.tsx` | edit (renombrar botón) |
| `__tests__/python/test_prompt_intent.py` | NUEVO |
| `__tests__/python/test_planner_enforcement.py` | NUEVO |

## Riesgos

- **Cache invalidation**: el `compute_cache_key` actual incluye `user_prompt` así que cambios en el intent invalidan cache automáticamente. ✓
- **Fuzzy match agresivo**: si bajamos `ratio` mucho, "riesgos" matchearía "Riesgos Operativos" cuando el user pidió "Riesgos Core". Mantener `ratio ≥ 0.65` y al empate preferir exact-substring match.
- **Padding sin bloques**: si el user pide 15 slides y el Excel solo tiene material para 8, `_pad_to` se queda corto. Solución: respeta el techo natural y reporta en `intent_report.count_honored = false` con mensaje "tu Excel solo tiene material para N slides".
- **Eliminar textarea del drawer rompe usuarios que ya lo usaban**: nadie lo usa así (es un proyecto pre-launch). No es un riesgo real.
