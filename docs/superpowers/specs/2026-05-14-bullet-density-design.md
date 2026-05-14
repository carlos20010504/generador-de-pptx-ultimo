# Bullet Density + Insight Bullets — Design

**Date:** 2026-05-14
**Status:** Approved (verbal)
**Author:** rashid.andrade@gmail.com (con Claude)
**Branch base:** `feature/hermes-driven-pipeline`

## Problema

User ya tiene el deck obediente (cantidad + hojas honradas), pero las slides de `text_bullets` siguen viéndose pobres en su Excel real:
1. A veces salen con solo 2-3 viñetas — el AI no respeta "5-7 bullets" del prompt o el validator dropea por provenance.
2. Hay UNA sola slide de hallazgos en todo el deck — el resto son charts/tablas.
3. Los bullets son facts ("Total: $1.5B", "1852 registros") cuando podrían ser insights ("Top 3 concentra 78% — riesgo de dependencia").

Quote: "quiero powerpoint de alto impacto visual y de alto impacto informativo".

## Objetivo

Garantizar que cada slide de hallazgos del deck final tenga **≥5 bullets con insights** (no solo cifras crudas), y que el deck contenga **≥1 slide de hallazgos por hoja importante** intercalada con sus charts/tablas.

No-objetivos:
- Cambiar el renderer (cómo se dibujan visualmente los bullets en PPTX) — eso es otra discusión.
- Tocar el validator (sigue siendo strict provenance).
- Generar más texto de relleno — preferimos menos bullets que padding genérico.

## Decisiones tomadas

| Decisión | Opción | Por qué |
|---|---|---|
| Min bullets/slide | **5**, max 6 | Balance entre densidad y readability |
| Source de extra bullets | **Deterministic top-up** desde `_build_fallback_bullets` enriquecido | Pasa validator garantizado, no depende de AI |
| Cantidad de hallazgos slides | **1 por hoja importante** cubierta en el deck | Deck "respira" insights entre charts |
| Posición | **Sandwich**: hallazgos VA después de su chart/tabla referente | Narrativa fluida |
| Calidad de bullets | **Mix de ángulos**: total + distribución + outlier + Pareto + brecha | Cada slide se siente con substance |

## Arquitectura

### Capas afectadas

```
plan from AI
    ↓
validate_plan() — sin cambios
    ↓
extract_for_render() — top-up bullets si <5  ← NUEVO
    ↓
auto_complete_slides() — ahora también añade text_bullets ← NUEVO PASO
    ↓
rendered → renderer
```

### Cambios concretos

#### `socya_pipeline/insights.py` — sin cambios

Reusamos `pareto_share`, `iqr_outliers`, `trend_acceleration`, `seasonality_hint` que ya existen.

#### `socya_pipeline/planner.py`

Refactor de `_build_fallback_bullets`:
- Renombre a `_build_insight_bullets(wb, block, max_bullets=8)` — devuelve hasta 8 candidatos.
- Cada candidato es `(angle: str, bullet: str)`. `angle` ∈ `{"total", "count", "top_cat", "max", "min", "second_cat", "pareto", "outlier_high", "outlier_low", "range"}`.
- El caller filtra para variedad (no dos del mismo angle).

Generators nuevos (todos best-effort, degradan a no-op si no aplica):

```python
def _bullet_pareto(wb, block) -> Optional[tuple[str, str]]:
    """Pareto top-3 sobre la columna numérica principal del block.
    Devuelve ('pareto', 'Top 3 X concentran Y% del total — riesgo de
    dependencia') cuando share ≥ 0.6."""

def _bullet_outlier(wb, block) -> Optional[tuple[str, str]]:
    """Outlier alto vía IQR. Devuelve ('outlier_high', 'X registró Y,
    Zx la mediana — caso atípico')."""

def _bullet_brecha(wb, block) -> Optional[tuple[str, str]]:
    """Brecha entre top y bottom de una columna categórica. Devuelve
    ('range', 'Brecha de N% entre X y Y — concentración geográfica')."""
```

#### `socya_pipeline/extractor.py`

**Pieza 1 — top-up dentro de `extract_for_render` (text_bullets branch):**

```python
elif stype == "text_bullets":
    bullets = slide.get("bullets") or []
    if len(bullets) < MIN_BULLETS:
        bullets = _topup_bullets(bullets, block, wb)
    if bullets:
        rendered.append({**slide, "data": {"bullets": bullets}})
    else:
        dropped.append(...)
```

Donde `_topup_bullets`:
- Llama `_build_insight_bullets(wb, block, max_bullets=8)`
- Dedup contra existentes (lowercase + clave por números/keywords)
- Mix de angles para variedad
- Devuelve hasta `MAX_BULLETS=6` total

**Pieza 2 — `auto_complete_slides` añade hallazgos:**

Después del paso de tablas, antes del return final:

```python
def _add_hallazgos_per_sheet(rendered, inv, wb, blocks_by_id) -> list:
    """Para cada hoja en el deck que tiene chart/table pero NO tiene
    text_bullets propia, inserta una slide de hallazgos justo después.
    Sandwich: chart_X → hallazgos_X → chart_Y → hallazgos_Y."""
```

Heurística "hoja importante": aparece en `provenance.sheet` de alguna slide ya rendered, NO redundant_sheet, tiene block tipo `table` o ≥1 chart con datos.

#### Constantes

```python
MIN_BULLETS = 5
MAX_BULLETS = 6
```

### Tests

Nuevo `__tests__/python/test_bullet_density.py`:
- `test_topup_when_ai_returned_too_few` — 2 bullets in → ≥5 bullets out
- `test_topup_dedup_against_existing` — AI bullet "total $1.5B" + fallback bullet "total $1.5B" → only one stays
- `test_topup_mix_angles` — output contains different generators (not 5x "total")
- `test_auto_complete_adds_hallazgos_per_sheet` — deck with 2 sheets covered → hallazgos per sheet
- `test_auto_complete_skips_sheets_with_existing_hallazgos` — no duplicates
- `test_pareto_bullet_emits_when_share_geq_60pct` + `_skipped_below_threshold`
- `test_outlier_bullet_emits_when_iqr_outlier_present`

## Riesgos

- **Bullets duplicados si dedup falla**: test específico evita. Dedup keys por (numeric_value_rounded, top_keyword).
- **Slide overflow visual**: 6 bullets puede no entrar en una slide standard si son largos. Mitigación: cortar bullets a 140 chars, renderer ya maneja line wrap. Si overflow real → followup.
- **Auto-add agrega muchos hallazgos = deck inflado**: cap de hallazgos auto-añadidos = 3 (el AI ya tiene 1 propio). Si Excel tiene 10 hojas, no metemos 10 hallazgos.
- **Validator dropea bullets fallback**: el `_build_insight_bullets` los construye sobre stats del block (sum/mean/min/max/top_values) que YA están en el haystack del validator. Tested.
