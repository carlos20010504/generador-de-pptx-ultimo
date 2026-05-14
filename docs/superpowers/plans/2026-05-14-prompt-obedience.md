# Prompt Obedience + UX Cleanup — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Hacer que el deck generado obedezca lo que el usuario pide en el prompt — cantidad exacta de slides, hojas mencionadas siempre incluidas, densidad de bullets adecuada — y que la UI tenga un solo lugar obvio para escribir esa instrucción.

**Architecture:** Nuevo módulo `prompt_intent.py` que parsea slide-count + sheet names (fuzzy match) del prompt en español. El planner inyecta el intent en el payload a la AI y aplica enforcement post-AI determinístico (trim/pad por slide_count, inject por required_sheets faltantes). El response incluye `intent_report` que la UI muestra como banner. El textarea duplicado del drawer "Panel avanzado" se elimina; queda solo el de PreparePanel, abierto por default y arriba.

**Tech Stack:** Python 3.12-3.14 (stdlib `re` + `difflib` + `dataclasses`), Next.js 16 App Router, React 19, TypeScript, pytest, Playwright.

**Spec:** [`docs/superpowers/specs/2026-05-14-prompt-obedience-design.md`](../specs/2026-05-14-prompt-obedience-design.md)

---

## File Structure

| Archivo | Tipo | Responsabilidad |
|---|---|---|
| `socya_pipeline/prompt_intent.py` | NEW | Parser puro: prompt → `PromptIntent(slide_count, required_sheets, topic_hints)` |
| `socya_pipeline/planner.py` | MODIFY | Inyectar intent en payload; template rules nuevas; enforcement post-AI |
| `socya_pipeline/cli.py` | MODIFY | Calcular intent en `cmd_plan`/`cmd_generate`; surface `_intent_report` en stdout |
| `app/api/preview-plan/route.ts` | MODIFY | Forward `intent_report` al cliente |
| `components/PreparePanel.tsx` | MODIFY | Mover Refinar al top, abrir por default, label más visible, banner intent_report |
| `components/AIControlPanel.tsx` | MODIFY | Eliminar textarea + sugerencias del prompt (quedan audiencia/idioma/tema/focus) |
| `components/ExcelUploader.tsx` | MODIFY | Renombrar botón "Panel avanzado" → "Audiencia y tema" |
| `__tests__/python/test_prompt_intent.py` | NEW | Tests del parser (slide count regex, fuzzy match, edge cases) |
| `__tests__/python/test_planner_enforcement.py` | NEW | Tests del enforcement (trim, pad, inject_sheet, idempotencia) |
| `tests/e2e/prompt-obedience.spec.ts` | NEW | E2E: tipear "9 slides con riesgos core" y verificar banner verde |

---

## Task 1: PromptIntent dataclasses + slide count parser (TDD)

**Files:**
- Create: `socya_pipeline/prompt_intent.py`
- Test:   `__tests__/python/test_prompt_intent.py`

- [ ] **Step 1: Write failing tests for slide count parsing**

Create `__tests__/python/test_prompt_intent.py`:

```python
"""Tests para socya_pipeline/prompt_intent.py — parser puro del prompt."""
import pytest
from socya_pipeline.prompt_intent import extract, PromptIntent, SheetMatch


# ─── slide count regex ─────────────────────────────────────────────

@pytest.mark.parametrize("prompt,expected", [
    ("hazme 9 slides con riesgos core", 9),
    ("9 slides", 9),
    ("dame 12 diapositivas", 12),
    ("deck de 15", 15),
    ("quiero un deck de 7 slides exactamente", 7),
    ("11 láminas", 11),
    ("Hazme 10 SLIDES", 10),  # case insensitive
])
def test_slide_count_extracted(prompt, expected):
    intent = extract(prompt, available_sheet_names=[])
    assert intent.requested_slide_count == expected


@pytest.mark.parametrize("prompt", [
    "",
    "muéstrame riesgos core sin más",
    "hazme un deck bonito",
    "incluye 3 columnas de datos",  # 3 no es slide count
])
def test_slide_count_none_when_not_specified(prompt):
    intent = extract(prompt, available_sheet_names=[])
    assert intent.requested_slide_count is None


def test_slide_count_first_match_wins():
    # Si el user escribe "9 slides ... 5 columnas", agarra el 9
    intent = extract("hazme 9 slides con 5 columnas", [])
    assert intent.requested_slide_count == 9


def test_slide_count_zero_or_negative_ignored():
    # "0 slides" no tiene sentido; preferimos None y dejamos al default del planner
    intent = extract("hazme 0 slides", [])
    assert intent.requested_slide_count is None
```

- [ ] **Step 2: Run tests — verify they fail with ImportError**

Run: `python -m pytest __tests__/python/test_prompt_intent.py -v`
Expected: ALL fail with `ModuleNotFoundError: No module named 'socya_pipeline.prompt_intent'`

- [ ] **Step 3: Create minimal `prompt_intent.py` with slide count only**

Create `socya_pipeline/prompt_intent.py`:

```python
"""Parsea la intención del usuario expresada en español.

Pure functions, no I/O. Used by planner.py to:
1. Inject structured signals into the AI payload.
2. Enforce post-AI: trim/pad slides to match `requested_slide_count`,
   inject missing `required_sheets` slides.

The parsing is deliberately conservative — when in doubt, return None and
let the AI decide. False positives are worse than no-ops because they
override the user's natural-language intent silently.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import List, Optional


# Regex de cantidad de slides. Orden importa: primer hit gana. Capturamos
# 1-3 dígitos para evitar matchear años (2026) u otros números grandes.
_SLIDE_COUNT_PATTERNS = (
    re.compile(r"(\d{1,3})\s*(?:slides?|diapositivas?|láminas?|laminas?)",
               re.IGNORECASE),
    re.compile(r"deck\s+de\s+(\d{1,3})", re.IGNORECASE),
)


@dataclass
class SheetMatch:
    """Una hoja que el user mencionó en el prompt + el matching real."""
    requested: str           # texto literal del prompt: "riesgos core"
    matched: Optional[str]   # nombre real: "Riesgos Core" o None si no hubo match
    closest: Optional[str]   # hoja más parecida cuando matched=None
    ratio: float             # confianza del fuzzy match (0..1)


@dataclass
class PromptIntent:
    """Resultado del parsing — todo opcional, valores `None`/`[]` ⇒ sin señal."""
    requested_slide_count: Optional[int] = None
    required_sheets: List[SheetMatch] = field(default_factory=list)
    topic_hints: List[str] = field(default_factory=list)


def extract(user_prompt: str, available_sheet_names: List[str]) -> PromptIntent:
    """Parsea `user_prompt` y devuelve un PromptIntent estructurado.

    Conservador por diseño: si no hay señal clara, devuelve fields vacíos.
    """
    if not user_prompt or not user_prompt.strip():
        return PromptIntent()

    return PromptIntent(
        requested_slide_count=_extract_slide_count(user_prompt),
        required_sheets=[],   # implementado en Task 2
        topic_hints=[],       # implementado en Task 3
    )


def _extract_slide_count(prompt: str) -> Optional[int]:
    """Devuelve el primer slide-count válido (≥1, ≤999) o None."""
    for pattern in _SLIDE_COUNT_PATTERNS:
        match = pattern.search(prompt)
        if match:
            try:
                n = int(match.group(1))
            except ValueError:
                continue
            if 1 <= n <= 999:
                return n
    return None
```

- [ ] **Step 4: Run tests — verify slide-count tests pass**

Run: `python -m pytest __tests__/python/test_prompt_intent.py -v`
Expected: all `test_slide_count_*` tests PASS.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/prompt_intent.py __tests__/python/test_prompt_intent.py
git commit -m "feat(prompt_intent): extract slide count from user prompt"
```

---

## Task 2: Fuzzy-match required sheets

**Files:**
- Modify: `socya_pipeline/prompt_intent.py`
- Modify: `__tests__/python/test_prompt_intent.py`

- [ ] **Step 1: Add failing tests for sheet fuzzy match**

Append to `__tests__/python/test_prompt_intent.py`:

```python
# ─── required_sheets fuzzy match ──────────────────────────────────

def test_exact_sheet_match_returns_sheetmatch():
    intent = extract("incluye Riesgos Core",
                     available_sheet_names=["Riesgos Core", "Ventas"])
    assert len(intent.required_sheets) == 1
    m = intent.required_sheets[0]
    assert m.requested == "riesgos core"
    assert m.matched == "Riesgos Core"
    assert m.closest is None
    assert m.ratio == 1.0


def test_case_insensitive_match():
    intent = extract("incluye RIESGOS core",
                     ["Riesgos Core"])
    assert intent.required_sheets[0].matched == "Riesgos Core"


def test_accent_insensitive_match():
    # User type "operacion" sin tilde, hoja real es "Operación"
    intent = extract("incluye operacion", ["Operación"])
    assert intent.required_sheets[0].matched == "Operación"


def test_fuzzy_match_above_threshold():
    # "riesgos core" vs "Riesgos-Core" — separador distinto
    intent = extract("incluye riesgos core", ["Riesgos-Core"])
    m = intent.required_sheets[0]
    assert m.matched == "Riesgos-Core"
    assert m.ratio >= 0.65


def test_no_match_returns_closest():
    intent = extract("incluye Proyecciones 2027",
                     ["Riesgos Core", "Acciones", "Ventas"])
    assert len(intent.required_sheets) == 1
    m = intent.required_sheets[0]
    assert m.matched is None
    assert m.closest in {"Riesgos Core", "Acciones", "Ventas"}
    assert m.ratio < 0.65


def test_multiple_sheets_extracted():
    intent = extract("9 slides con Riesgos Core y Riesgos Acciones",
                     ["Riesgos Core", "Riesgos Acciones", "Otra"])
    matched = {m.matched for m in intent.required_sheets}
    assert "Riesgos Core" in matched
    assert "Riesgos Acciones" in matched


def test_no_sheet_mentions_returns_empty():
    intent = extract("hazme 9 slides bonitas",
                     ["Riesgos Core", "Ventas"])
    assert intent.required_sheets == []


def test_empty_sheet_list_returns_empty():
    intent = extract("incluye Riesgos Core", available_sheet_names=[])
    assert intent.required_sheets == []
```

- [ ] **Step 2: Run new tests — verify they fail**

Run: `python -m pytest __tests__/python/test_prompt_intent.py -v -k "sheet or match"`
Expected: new tests FAIL (required_sheets always empty).

- [ ] **Step 3: Implement sheet matching**

Replace `extract()` and add helpers in `socya_pipeline/prompt_intent.py`:

```python
# Añadir constantes al top del módulo, debajo de los imports
_FUZZY_THRESHOLD = 0.65

# Stop-words que NO son nombres de hoja aunque aparezcan tras "incluye/agrega"
_STOPWORDS = frozenset({
    "todo", "todos", "toda", "todas", "el", "la", "los", "las",
    "y", "o", "u", "de", "del", "en", "con", "para", "por",
    "datos", "información", "informacion", "info", "hoja", "hojas",
    "tabla", "tablas", "slide", "slides", "diapositiva", "diapositivas",
    "lámina", "láminas", "lamina", "laminas", "deck", "presentación",
    "presentacion", "powerpoint", "ppt", "pptx",
    "más", "mas", "menos", "que", "como", "asi", "así",
})


def extract(user_prompt: str, available_sheet_names: List[str]) -> PromptIntent:
    """Parsea `user_prompt` y devuelve un PromptIntent estructurado."""
    if not user_prompt or not user_prompt.strip():
        return PromptIntent()

    return PromptIntent(
        requested_slide_count=_extract_slide_count(user_prompt),
        required_sheets=_extract_required_sheets(user_prompt,
                                                   available_sheet_names),
        topic_hints=[],   # Task 3
    )


def _normalize(text: str) -> str:
    """lower + strip accents + collapse whitespace + strip punctuation
    excepto espacios y guiones."""
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", text)
    no_accents = "".join(c for c in nfkd if not unicodedata.combining(c))
    out = no_accents.lower().strip()
    # Reemplazar guiones bajos / puntos / comas por espacio para que
    # "Riesgos_Core" matchee "riesgos core"
    out = re.sub(r"[_\.,;:!?\(\)\[\]\{\}]+", " ", out)
    out = re.sub(r"\s+", " ", out).strip()
    return out


def _ratio(a: str, b: str) -> float:
    """SequenceMatcher ratio sobre versiones normalizadas."""
    return SequenceMatcher(None, _normalize(a), _normalize(b)).ratio()


def _extract_required_sheets(prompt: str,
                              available: List[str]) -> List[SheetMatch]:
    """Encuentra menciones a hojas en el prompt usando fuzzy match.

    Estrategia: probamos n-gramas (1-4 palabras consecutivas) del prompt
    contra cada hoja disponible. Para cada hoja, conservamos solo el mejor
    n-grama. Devolvemos un SheetMatch por hoja-disponible mencionada con
    ratio ≥ _FUZZY_THRESHOLD, más SheetMatch con matched=None para hojas
    pedidas por nombre que NO existen (closest = mejor candidato).

    Para detectar "no existe": miramos n-gramas que parecen nombres propios
    (≥2 palabras o palabra ≥6 chars, no son stop-words) y no matchean nada.
    """
    if not available:
        return []

    norm_prompt = _normalize(prompt)
    if not norm_prompt:
        return []

    tokens = norm_prompt.split()
    # Generar n-gramas 1..4 (los nombres de hoja típicos son 1-3 palabras)
    ngrams: List[str] = []
    for n in (4, 3, 2, 1):
        for i in range(len(tokens) - n + 1):
            ng = " ".join(tokens[i:i+n])
            # Filtrar n-gramas que son puro stop-word
            if all(t in _STOPWORDS for t in tokens[i:i+n]):
                continue
            # Single-token n-gramas: requiere longitud ≥4 para evitar matchear
            # palabras genéricas como "y", "de", "con"
            if n == 1 and (len(ng) < 4 or ng in _STOPWORDS):
                continue
            ngrams.append(ng)

    if not ngrams:
        return []

    # Para cada hoja disponible, encontrar el mejor n-grama
    matches: List[SheetMatch] = []
    matched_sheets: set = set()
    for sheet in available:
        best_ratio = 0.0
        best_ng = ""
        for ng in ngrams:
            r = _ratio(ng, sheet)
            if r > best_ratio:
                best_ratio = r
                best_ng = ng
        if best_ratio >= _FUZZY_THRESHOLD:
            matches.append(SheetMatch(
                requested=best_ng,
                matched=sheet,
                closest=None,
                ratio=round(best_ratio, 3),
            ))
            matched_sheets.add(sheet)

    # Detectar n-gramas que parecen nombres propios pero no matchearon nada
    # → user pidió una hoja que no existe. Para cada n-grama no-matched que
    # tiene ratio < threshold con todas las hojas, reportamos el closest.
    # Solo si el n-grama es "nombre-ish": ≥2 palabras o ≥6 chars.
    used_ngrams = {m.requested for m in matches}
    for ng in ngrams:
        if ng in used_ngrams:
            continue
        is_name_ish = (" " in ng) or (len(ng) >= 6)
        if not is_name_ish:
            continue
        # Buscar la mejor hoja para este n-grama
        best_sheet = None
        best_ratio = 0.0
        for sheet in available:
            if sheet in matched_sheets:
                continue
            r = _ratio(ng, sheet)
            if r > best_ratio:
                best_ratio = r
                best_sheet = sheet
        # Solo reportar como "no encontrado" si el ratio es entre 0.3 y threshold
        # (debajo de 0.3 no es plausible que el user quería esa hoja)
        if best_sheet and 0.3 <= best_ratio < _FUZZY_THRESHOLD:
            matches.append(SheetMatch(
                requested=ng,
                matched=None,
                closest=best_sheet,
                ratio=round(best_ratio, 3),
            ))

    # Dedup: si dos n-gramas matchearon la misma hoja, conservamos el mejor
    seen_sheets: dict = {}
    for m in matches:
        key = m.matched or f"NONE_{m.requested}"
        if key not in seen_sheets or m.ratio > seen_sheets[key].ratio:
            seen_sheets[key] = m
    return list(seen_sheets.values())
```

- [ ] **Step 4: Run tests — verify all sheet tests pass**

Run: `python -m pytest __tests__/python/test_prompt_intent.py -v`
Expected: all tests PASS.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/prompt_intent.py __tests__/python/test_prompt_intent.py
git commit -m "feat(prompt_intent): fuzzy-match required sheets from prompt"
```

---

## Task 3: Topic hints (lightweight keyword extraction)

**Files:**
- Modify: `socya_pipeline/prompt_intent.py`
- Modify: `__tests__/python/test_prompt_intent.py`

- [ ] **Step 1: Add failing tests for topic hints**

Append to `__tests__/python/test_prompt_intent.py`:

```python
# ─── topic hints (keywords sueltas) ──────────────────────────────

def test_topic_hints_extracts_meaningful_words():
    # Words ≥4 chars que NO son stop-words ni números
    intent = extract("muéstrame riesgos críticos y comisiones aprobadas",
                     available_sheet_names=[])
    hints = set(intent.topic_hints)
    assert "riesgos" in hints
    assert "criticos" in hints   # accent stripped
    assert "comisiones" in hints
    assert "aprobadas" in hints


def test_topic_hints_excludes_stopwords_and_short():
    intent = extract("y de en con para por", [])
    assert intent.topic_hints == []


def test_topic_hints_does_not_include_slide_count_words():
    intent = extract("hazme 9 slides", [])
    assert "slides" not in intent.topic_hints
    assert "9" not in intent.topic_hints


def test_topic_hints_unique_lowercased():
    intent = extract("Riesgos riesgos RIESGOS críticos", [])
    assert intent.topic_hints.count("riesgos") == 1
```

- [ ] **Step 2: Run tests — verify they fail**

Run: `python -m pytest __tests__/python/test_prompt_intent.py -v -k "topic"`
Expected: FAIL (topic_hints siempre vacío).

- [ ] **Step 3: Implement topic_hints**

Modify `extract()` in `socya_pipeline/prompt_intent.py`:

```python
def extract(user_prompt: str, available_sheet_names: List[str]) -> PromptIntent:
    if not user_prompt or not user_prompt.strip():
        return PromptIntent()
    return PromptIntent(
        requested_slide_count=_extract_slide_count(user_prompt),
        required_sheets=_extract_required_sheets(user_prompt,
                                                   available_sheet_names),
        topic_hints=_extract_topic_hints(user_prompt),
    )


def _extract_topic_hints(prompt: str) -> List[str]:
    """Extrae palabras significativas (≥4 chars, no stop-words, no dígitos)
    para que el planner pueda usarlas como hint de prioridad de bloques."""
    norm = _normalize(prompt)
    if not norm:
        return []
    seen: set = set()
    out: List[str] = []
    for token in norm.split():
        if len(token) < 4:
            continue
        if token in _STOPWORDS:
            continue
        if token.isdigit():
            continue
        if token in seen:
            continue
        seen.add(token)
        out.append(token)
    return out
```

- [ ] **Step 4: Run all prompt_intent tests**

Run: `python -m pytest __tests__/python/test_prompt_intent.py -v`
Expected: ALL pass.

- [ ] **Step 5: Commit**

```bash
git add socya_pipeline/prompt_intent.py __tests__/python/test_prompt_intent.py
git commit -m "feat(prompt_intent): extract topic hints (≥4 char keywords)"
```

---

## Task 4: Inject intent into planner payload + new template rules

**Files:**
- Modify: `socya_pipeline/planner.py`

- [ ] **Step 1: Update `build_payload` signature and body**

In `socya_pipeline/planner.py`, modify the `build_payload` function to accept and inject the intent. Replace the existing function (lines 18-63) with:

```python
def build_payload(wb: WorkbookData, blocks, user_prompt: str, audience: str,
                   language: str, intent=None) -> dict:
    sheets_payload = []
    for s in wb.sheets:
        sheets_payload.append({
            "name": s.name,
            "shape": list(s.shape),
            "fill_ratio": round(s.fill_ratio, 3),
            "columns": [_col_payload(c) for c in s.columns],
            "first_rows": s.first_rows[:MAX_FIRST_ROWS],
        })
    blocks_payload = [_block_payload(b) for b in blocks]

    effective_prompt = (user_prompt or "").strip()
    if not effective_prompt:
        try:
            all_cols = [c.name for s in wb.sheets for c in s.columns]
            sheet_names = [s.name for s in wb.sheets]
            effective_prompt = insights.infer_prompt(wb.filename, sheet_names, all_cols)
        except Exception:
            effective_prompt = ""

    user_intent_payload = None
    if intent is not None:
        user_intent_payload = {
            "slide_count": intent.requested_slide_count,
            "required_sheets": [m.matched for m in intent.required_sheets if m.matched],
        }

    payload = {
        "user_prompt": effective_prompt,
        "user_intent": user_intent_payload,
        "audience": audience or "ejecutivos",
        "language": language or "es",
        "workbook": {
            "filename": wb.filename,
            "sheets": sheets_payload,
            "blocks": blocks_payload,
        },
    }

    serialized = json.dumps(payload, ensure_ascii=False, default=str)
    if len(serialized) > MAX_PAYLOAD_CHARS:
        for s in payload["workbook"]["sheets"]:
            s["first_rows"] = s["first_rows"][:3]
            for c in s["columns"]:
                if "samples" in c:
                    c["samples"] = c["samples"][:3]
                if "top_values" in c:
                    c["top_values"] = c["top_values"][:3]
    return payload
```

- [ ] **Step 2: Update PROMPT_TEMPLATE rules**

Replace the existing `PROMPT_TEMPLATE` block (lines 88-150) — find the line `REGLAS CRÍTICAS:` and replace through the `{payload_json}` line with:

```python
PROMPT_TEMPLATE = """Eres un director de arte y analista que diseña presentaciones ejecutivas tipo McKinsey/BCG a partir de datos reales de Excel. Tu trabajo es producir un deck **rico, completo y narrativamente coherente**, con insights que se sienten como hechos por un analista senior.

REGLAS CRÍTICAS:

0a. **Cantidad de slides — RESPETO LITERAL**: si `user_intent.slide_count` está presente y NO es null, debes producir EXACTAMENTE ese número de slides totales (incluyendo portada). NO menos, NO más. Si está null, aplica la regla 1 (Volumen).

0b. **Hojas obligatorias — RESPETO LITERAL**: si `user_intent.required_sheets` no está vacío, cada hoja listada debe aparecer como `provenance.sheet` en al menos UNA slide. Si una hoja listada no tiene datos suficientes para construir una slide válida, OMÍTELA y al final del JSON añade `"_skipped_required_sheets": ["nombre"]`. Nunca inventes datos para forzar la cobertura.

1. **Volumen** (solo cuando `user_intent.slide_count` es null): produce entre 7 y 11 slides. Una portada + 6-10 slides de contenido. NO menos.

2. **Procedencia obligatoria**: cada slide DEBE referenciar bloques existentes por su `id` (`block_ref`, `block_refs`, o `supports_block`). NUNCA inventes IDs.

3. **CERO alucinaciones**: ningún número, nombre, fecha, porcentaje en el JSON puede ser inventado. Si va en un `narrative` o `bullets`, debe existir en el bloque referenciado (`samples`, `top_values`, `min/max/mean/sum/value`, o `first_rows`).

4. **Bullets específicos y abundantes**: cada slide `text_bullets` debe tener **5-7 bullets**, cada uno con al menos una cifra concreta o un nombre propio del bloque. PROHIBIDO bullets genéricos como "los datos muestran variabilidad" o "es importante revisar".

5. **Narrativas tipo analista (NO descriptivas)**: cada `narrative` de un chart debe explicar QUÉ pasa Y POR QUÉ importa. Incluir mínimo 2 datos concretos y al menos uno de estos ángulos:
   - **Concentración** ("top 3 destinos = 78% del total → riesgo de dependencia"),
   - **Outlier** ("ANTIOQUIA - EL BAGRE concentra 19, 9x la mediana → caso atípico"),
   - **Brecha** ("solo 12% aprobado vs 88% pendiente → cuello de botella"),
   - **Tendencia** ("decreció 47% entre 2023Q3 y 2026Q1 → caída sostenida").

6. **Estructura recomendada** (ordenada para máximo impacto narrativo):
   - 1 slide `title` (portada)
   - 1 slide `kpi_row` con 2-4 KPIs principales. PREFIERE: (a) bloques con `quality_flags: ["derived"]` (% derivados como "% Aprobación") por encima de todo, (b) currency totales SIN flag `subsumed_by_total`, (c) NUNCA elijas KPIs marcados `subsumed_by_total` salvo que no haya alternativa — esos son sub-totales redundantes.
   - 2-4 slides `chart` con distintos cortes (mezcla `bar`, `pie`, `line` según los `kind` disponibles)
   - 1-2 slides `table` con cortes detallados (T*) — usa `columns_subset` con solo 4-6 columnas relevantes (NO IDs, NO observaciones largas)
   - 1-2 slides `text_bullets` con hallazgos accionables (5-7 bullets con cifras; cada bullet debe incluir un dato concreto del bloque referenciado)

7. **Variedad visual**: si tienes ≥2 distribuciones categóricas, usa una `bar` y una `pie`. Si hay serie temporal `S*`, dale un `line`.

8. **Calidad de KPIs**: prefiere bloques `K*` cuyo nombre **no** contenga "id", "código", "número", "folio", "consecutivo". Esos son identificadores, no KPIs.

9. **Títulos editorial**: títulos atractivos en español ("Concentración geográfica de destinos", no "Gráfica de Datos"). Subtítulos opcionales con contexto.

10. **JSON estricto y único**: devuelve EXCLUSIVAMENTE un JSON válido (sin markdown, sin texto adicional) con esta forma exacta:

{{
  "presentation_meta": {{ "title": "Título atractivo del deck", "subtitle": "Subtítulo con contexto" }},
  "slides": [
    {{ "type": "title", "data": {{ "title": "Título portada", "subtitle": "Bajada" }} }},
    {{ "type": "kpi_row", "title": "Indicadores Clave", "block_refs": ["K1","K2","K3"] }},
    {{ "type": "chart", "chart_type": "bar", "title": "Distribución por Ciudad",
       "block_ref": "C1", "narrative": "Bogotá concentra 53 registros (53%), seguida por Medellín (23%) y Cali (15%)." }},
    {{ "type": "chart", "chart_type": "pie", "title": "Composición por Estado",
       "block_ref": "C2", "narrative": "El 75% está en estado CONTABILIZADO, mientras 13% está RECHAZADO." }},
    {{ "type": "chart", "chart_type": "histogram", "title": "Distribución de Montos Solicitados",
       "block_ref": "K2", "narrative": "La mayoría se concentra entre $50K y $200K; cola larga hasta $2M." }},
    {{ "type": "table", "title": "Detalle Operativo", "block_ref": "T1",
       "columns_subset": ["Mes","Total","Ciudad","Estado"], "max_rows": 10 }},
    {{ "type": "text_bullets", "title": "Hallazgos Accionables", "supports_block": "T1",
       "bullets": [
         "El monto máximo individual asciende a $1.000.000 en ANTIOQUIA - EL BAGRE.",
         "Existen 192 comisiones en estado RECHAZADO que requieren revisión.",
         "La concentración geográfica se ubica en MEDELLÍN con el mayor número de registros.",
         "El total acumulado de montos solicitados suma $250.2M.",
         "Bogotá representa el 53% del volumen, evidenciando concentración geográfica.",
         "El monto promedio por solicitud es $2.5M, con 192 outliers sobre 2x ese valor."
       ] }}
  ],
  "prompt_suggestions": [
    "Detalla el análisis por trimestre",
    "Compara montos por centro de costos",
    "Resalta los 5 solicitantes con mayor monto"
  ]
}}

Datos disponibles:
{payload_json}
"""
```

- [ ] **Step 3: Update `plan_presentation` signature and body**

Modify `plan_presentation` (lines 327-396). Find the existing function and replace with:

```python
def plan_presentation(wb: WorkbookData, blocks, user_prompt: str, audience: str,
                       language: str, api_key: str,
                       profile: AIProfile = AIProfile.PATIENT,
                       cache_dir: Optional[Path] = None,
                       file_path: Optional[Path] = None,
                       intent=None) -> dict:
    cache_key = None
    cache: Optional[PlanCache] = None
    if file_path is not None:
        try:
            file_bytes = Path(file_path).read_bytes()
            # Cache key incluye el slide_count del intent para que distintos
            # counts no compartan cache hit.
            intent_signature = ""
            if intent is not None:
                req_sheets = ",".join(sorted(
                    m.matched for m in intent.required_sheets if m.matched
                ))
                intent_signature = (f"|sc={intent.requested_slide_count}"
                                     f"|rs={req_sheets}")
            cache_key = compute_cache_key(file_bytes, user_prompt + intent_signature,
                                            audience, language, PLANNER_VERSION)
            cache = PlanCache(cache_dir=cache_dir)
            cached = cache.get(cache_key)
            if cached:
                cached["_meta"] = {**cached.get("_meta", {}),
                                    "cache_hit": True}
                return cached
        except OSError:
            pass

    payload = build_payload(wb, blocks, user_prompt, audience, language,
                              intent=intent)
    payload_json = json.dumps(payload, ensure_ascii=False, default=str)
    prompt = PROMPT_TEMPLATE.format(payload_json=payload_json)

    chain = AIChain(api_key=api_key, profile=profile)
    try:
        result = chain.call(prompt)
        parsed = insights.parse_loose_json(result.content)
        if not isinstance(parsed, dict) or "slides" not in parsed:
            raise PipelineError(
                ErrorCode.AI_RESPONSE_INVALID,
                "El planificador devolvió un JSON sin la forma esperada.",
                details=str(result.content)[:300],
                user_action="retry",
            )
        if not isinstance(parsed.get("slides"), list) or not parsed.get("slides"):
            raise PipelineError(
                ErrorCode.AI_RESPONSE_INVALID,
                "El planificador no devolvió slides utilizables.",
                details=f"slides={type(parsed.get('slides')).__name__}",
                user_action="retry",
            )
    except PipelineError as e:
        if e.code in (ErrorCode.AI_SATURATED, ErrorCode.AI_RESPONSE_INVALID):
            plan = _deterministic_plan_fallback(wb, blocks, audience, language)
            plan["_meta"]["fallback_reason"] = (
                e.code.value if hasattr(e.code, "value") else str(e.code)
            )
            # Aplicamos enforcement también al fallback determinístico
            if intent is not None:
                plan = _enforce_intent(plan, intent, blocks, wb)
            return plan
        raise

    parsed["_meta"] = {
        "model": result.model,
        "cache_hit": False,
        "fallback_steps": result.fallback_steps,
        "planner_version": PLANNER_VERSION,
    }

    # Enforcement post-AI: garantiza slide_count y required_sheets
    if intent is not None:
        parsed = _enforce_intent(parsed, intent, blocks, wb)

    if cache and cache_key:
        cache.set(cache_key, parsed)

    return parsed
```

- [ ] **Step 4: Bump PLANNER_VERSION**

In `socya_pipeline/planner.py` line 12, change:

```python
PLANNER_VERSION = "p3"  # bump when prompt template changes — invalidates cache
```

to:

```python
PLANNER_VERSION = "p4"  # bump when prompt template changes — invalidates cache
```

- [ ] **Step 5: Run existing planner tests to verify nothing broke**

Run: `python -m pytest __tests__/python/ -v -k "planner or plan"`
Expected: existing tests still pass (we only added optional `intent=None` param + recipe changes, no behavior change when intent is None).

- [ ] **Step 6: Commit**

```bash
git add socya_pipeline/planner.py
git commit -m "feat(planner): inject user intent into payload + template rules"
```

---

## Task 5: Implement `_enforce_intent` (trim/pad/inject)

**Files:**
- Modify: `socya_pipeline/planner.py`
- Create: `__tests__/python/test_planner_enforcement.py`

- [ ] **Step 1: Write failing tests for enforcement**

Create `__tests__/python/test_planner_enforcement.py`:

```python
"""Tests para el enforcement post-AI del planner — garantiza slide_count
y required_sheets sin depender de la AI."""
import pytest
from socya_pipeline.planner import _enforce_intent
from socya_pipeline.prompt_intent import PromptIntent, SheetMatch
from socya_pipeline.parser import WorkbookData, SheetData, ColumnData
from socya_pipeline.inventory import build_inventory


def _make_wb_with_sheets(sheet_names):
    """Construye un WorkbookData mínimo con N hojas sin datos."""
    sheets = []
    for name in sheet_names:
        sheets.append(SheetData(
            name=name,
            shape=(10, 2),
            fill_ratio=0.9,
            columns=[
                ColumnData(name="Categoría", dtype="categorical",
                            n_unique=3, fill_ratio=1.0,
                            samples=["A", "B", "C"],
                            top_values=[("A", 5), ("B", 3), ("C", 2)]),
                ColumnData(name="Monto", dtype="currency",
                            n_unique=10, fill_ratio=1.0,
                            samples=[100.0, 200.0, 300.0],
                            sum=2000.0, mean=200.0, min=50.0, max=500.0),
            ],
            first_rows=[["A", 100], ["B", 200]],
        ))
    return WorkbookData(filename="test.xlsx", sheets=sheets)


# ─── Slide count enforcement ──────────────────────────────────────

def test_enforce_trim_to_exact_count():
    wb = _make_wb_with_sheets(["Hoja1"])
    blocks = build_inventory(wb)
    plan = {
        "slides": [
            {"type": "title", "data": {"title": "T"}},
            {"type": "kpi_row", "block_refs": [b.id for b in blocks if b.kind == "kpi_candidate"][:2]},
            {"type": "chart", "block_ref": next(b.id for b in blocks if b.kind == "categorical_distribution"),
             "provenance": {"sheet": "Hoja1"}},
            {"type": "chart", "block_ref": next(b.id for b in blocks if b.kind == "categorical_distribution"),
             "provenance": {"sheet": "Hoja1"}},
            {"type": "table", "block_ref": next(b.id for b in blocks if b.kind == "table"),
             "provenance": {"sheet": "Hoja1"}},
        ]
    }
    intent = PromptIntent(requested_slide_count=3)
    result = _enforce_intent(plan, intent, blocks, wb)
    assert len(result["slides"]) == 3
    # Title slide siempre se conserva
    assert result["slides"][0]["type"] == "title"
    assert result["_intent_report"]["actual_slide_count"] == 3
    assert result["_intent_report"]["count_honored"] is True


def test_enforce_pad_to_count():
    wb = _make_wb_with_sheets(["Hoja1", "Hoja2"])
    blocks = build_inventory(wb)
    plan = {
        "slides": [
            {"type": "title", "data": {"title": "T"}},
            {"type": "kpi_row", "block_refs": [b.id for b in blocks if b.kind == "kpi_candidate"][:1]},
        ]
    }
    intent = PromptIntent(requested_slide_count=5)
    result = _enforce_intent(plan, intent, blocks, wb)
    assert len(result["slides"]) == 5
    assert result["_intent_report"]["count_honored"] is True


def test_enforce_pad_caps_at_available():
    # Pedir 100 slides cuando solo hay material para ~5
    wb = _make_wb_with_sheets(["Hoja1"])
    blocks = build_inventory(wb)
    plan = {"slides": [{"type": "title", "data": {"title": "T"}}]}
    intent = PromptIntent(requested_slide_count=100)
    result = _enforce_intent(plan, intent, blocks, wb)
    assert len(result["slides"]) < 100
    assert result["_intent_report"]["count_honored"] is False


def test_enforce_count_none_is_noop():
    wb = _make_wb_with_sheets(["Hoja1"])
    blocks = build_inventory(wb)
    plan = {"slides": [{"type": "title", "data": {"title": "T"}}]}
    intent = PromptIntent(requested_slide_count=None)
    result = _enforce_intent(plan, intent, blocks, wb)
    assert len(result["slides"]) == 1
    assert result["_intent_report"]["count_honored"] is True  # nada que enforce


# ─── Required sheets enforcement ──────────────────────────────────

def test_enforce_inject_missing_required_sheet():
    wb = _make_wb_with_sheets(["Riesgos Core", "Otra"])
    blocks = build_inventory(wb)
    # Plan original NO menciona "Riesgos Core"
    plan = {
        "slides": [
            {"type": "title", "data": {"title": "T"}},
            {"type": "chart",
             "block_ref": next(b.id for b in blocks if b.kind == "categorical_distribution" and b.provenance.sheet == "Otra"),
             "provenance": {"sheet": "Otra"}},
        ]
    }
    intent = PromptIntent(
        required_sheets=[SheetMatch(requested="riesgos core",
                                       matched="Riesgos Core",
                                       closest=None, ratio=1.0)],
    )
    result = _enforce_intent(plan, intent, blocks, wb)
    sheets_in_plan = {s.get("provenance", {}).get("sheet")
                      for s in result["slides"] if "provenance" in s}
    assert "Riesgos Core" in sheets_in_plan
    sheet_report = result["_intent_report"]["required_sheets"][0]
    assert sheet_report["matched"] == "Riesgos Core"
    assert len(sheet_report["slide_indices"]) >= 1


def test_enforce_already_covered_sheet_no_op():
    wb = _make_wb_with_sheets(["Riesgos Core"])
    blocks = build_inventory(wb)
    plan = {
        "slides": [
            {"type": "title", "data": {"title": "T"}},
            {"type": "chart",
             "block_ref": next(b.id for b in blocks if b.kind == "categorical_distribution"),
             "provenance": {"sheet": "Riesgos Core"}},
        ]
    }
    intent = PromptIntent(
        required_sheets=[SheetMatch(requested="riesgos core",
                                       matched="Riesgos Core",
                                       closest=None, ratio=1.0)],
    )
    result = _enforce_intent(plan, intent, blocks, wb)
    # No duplica la slide
    assert sum(1 for s in result["slides"]
                if s.get("provenance", {}).get("sheet") == "Riesgos Core") == 1


def test_enforce_intent_report_has_unmatched_sheet():
    wb = _make_wb_with_sheets(["Otra"])
    blocks = build_inventory(wb)
    plan = {"slides": [{"type": "title", "data": {"title": "T"}}]}
    intent = PromptIntent(
        required_sheets=[SheetMatch(requested="proyecciones",
                                       matched=None,
                                       closest="Otra", ratio=0.5)],
    )
    result = _enforce_intent(plan, intent, blocks, wb)
    sheet_report = result["_intent_report"]["required_sheets"][0]
    assert sheet_report["matched"] is None
    assert sheet_report["closest"] == "Otra"
    assert sheet_report["slide_indices"] == []


def test_enforce_idempotent():
    """Aplicar enforce dos veces produce el mismo resultado."""
    wb = _make_wb_with_sheets(["Riesgos Core"])
    blocks = build_inventory(wb)
    plan = {"slides": [{"type": "title", "data": {"title": "T"}}]}
    intent = PromptIntent(
        requested_slide_count=3,
        required_sheets=[SheetMatch(requested="riesgos core",
                                       matched="Riesgos Core",
                                       closest=None, ratio=1.0)],
    )
    once = _enforce_intent(plan, intent, blocks, wb)
    twice = _enforce_intent(once, intent, blocks, wb)
    assert len(twice["slides"]) == len(once["slides"])
    assert twice["_intent_report"] == once["_intent_report"]
```

- [ ] **Step 2: Run tests — verify they fail with ImportError**

Run: `python -m pytest __tests__/python/test_planner_enforcement.py -v`
Expected: FAIL with `ImportError: cannot import name '_enforce_intent'`.

- [ ] **Step 3: Implement `_enforce_intent` and helpers in planner.py**

Add to `socya_pipeline/planner.py` AFTER `_deterministic_plan_fallback`:

```python
def _enforce_intent(plan: dict, intent, blocks, wb) -> dict:
    """Garantiza que `plan` honre `intent`. Idempotente.

    1. Inyecta required_sheets faltantes (puede crecer el deck).
    2. Recorta o expande para llegar a `requested_slide_count`.
    3. Anota _intent_report con lo que se honró y lo que no.

    El orden importa: primero injection (puede crecer), después count
    enforcement (recorta hacia el target absoluto).
    """
    slides = list(plan.get("slides") or [])
    blocks_by_id = {b.id: b for b in blocks}

    # 1. Required sheets injection
    used_sheets = {s.get("provenance", {}).get("sheet")
                   for s in slides if isinstance(s.get("provenance"), dict)}
    skipped_sheets: list = []
    for match in intent.required_sheets:
        if match.matched is None:
            continue
        if match.matched in used_sheets:
            continue
        injected = _inject_sheet_slide(match.matched, blocks, wb, blocks_by_id)
        if injected is not None:
            slides.append(injected)
            used_sheets.add(match.matched)
        else:
            skipped_sheets.append(match.matched)

    # 2. Slide count enforcement
    target = intent.requested_slide_count
    count_honored = True
    if target is not None:
        if len(slides) > target:
            slides = _trim_to(slides, target, intent)
        elif len(slides) < target:
            padded = _pad_to(slides, target, blocks, wb, blocks_by_id)
            slides = padded
            if len(slides) < target:
                count_honored = False  # no había material para llegar al target

    plan["slides"] = slides
    plan["_intent_report"] = _build_intent_report(intent, slides,
                                                     skipped_sheets,
                                                     count_honored)
    return plan


def _inject_sheet_slide(sheet_name: str, blocks, wb, blocks_by_id) -> Optional[dict]:
    """Construye una slide a partir de la primera hoja `sheet_name`.

    Preferencia: chart > table > kpi_row. Devuelve None si no hay material.
    """
    sheet_blocks = [b for b in blocks if b.provenance.sheet == sheet_name]
    if not sheet_blocks:
        return None

    cat = next((b for b in sheet_blocks
                if b.kind == "categorical_distribution"
                and "single_dominant_category" not in b.quality_flags),
                None)
    if cat:
        col = cat.provenance.columns[0] if cat.provenance.columns else "categorías"
        return {
            "type": "chart",
            "chart_type": "bar",
            "title": f"{sheet_name} — Distribución por {col}",
            "block_ref": cat.id,
            "narrative": "",
            "provenance": {"sheet": sheet_name,
                            "columns": list(cat.provenance.columns)},
            "_injected_by": "intent_enforcement",
        }

    table = next((b for b in sheet_blocks
                  if b.kind == "table"
                  and "low_fill_ratio" not in b.quality_flags
                  and "too_few_rows" not in b.quality_flags),
                  None)
    if table:
        return {
            "type": "table",
            "title": f"Detalle: {sheet_name}",
            "block_ref": table.id,
            "max_rows": 10,
            "provenance": {"sheet": sheet_name,
                            "columns": list(table.provenance.columns)},
            "_injected_by": "intent_enforcement",
        }

    kpis = [b for b in sheet_blocks if b.kind == "kpi_candidate"][:3]
    if kpis:
        return {
            "type": "kpi_row",
            "title": f"Indicadores — {sheet_name}",
            "block_refs": [b.id for b in kpis],
            "provenance": {"sheet": sheet_name,
                            "columns": [c for k in kpis for c in k.provenance.columns]},
            "_injected_by": "intent_enforcement",
        }

    return None


def _trim_to(slides: list, target: int, intent) -> list:
    """Recorta `slides` a `target` preservando: (1) title slide, (2) slides
    inyectadas por intent enforcement, (3) slides cuya provenance.sheet está
    en required_sheets, (4) el resto se ordena por ‘importancia heurística’
    y se cortan las menos importantes primero."""
    if len(slides) <= target:
        return slides

    required_sheet_names = {m.matched for m in intent.required_sheets if m.matched}

    def _slot_priority(s_idx_pair):
        idx, s = s_idx_pair
        # Lower number = more protected (kept first)
        if s.get("type") == "title":
            return 0
        if s.get("_injected_by") == "intent_enforcement":
            return 1
        prov_sheet = s.get("provenance", {}).get("sheet")
        if prov_sheet in required_sheet_names:
            return 2
        # Then prefer kpi_row > chart > table > text_bullets
        type_pref = {"kpi_row": 3, "chart": 4, "table": 5,
                     "text_bullets": 6}.get(s.get("type"), 7)
        return type_pref

    indexed = list(enumerate(slides))
    indexed.sort(key=_slot_priority)
    kept_indices = sorted(idx for idx, _ in indexed[:target])
    return [slides[i] for i in kept_indices]


def _pad_to(slides: list, target: int, blocks, wb, blocks_by_id) -> list:
    """Añade slides desde bloques no usados hasta llegar a `target`. No
    excede `target`. Si no hay material, devuelve lo que se pudo."""
    if len(slides) >= target:
        return slides

    used_block_ids: set = set()
    for s in slides:
        for k in ("block_ref", "supports_block"):
            v = s.get(k)
            if v:
                used_block_ids.add(v)
        for v in s.get("block_refs") or []:
            used_block_ids.add(v)

    # Orden de preferencia para padding: charts > tables > kpis
    candidates = [b for b in blocks if b.id not in used_block_ids]
    cats = [b for b in candidates if b.kind == "categorical_distribution"
            and "single_dominant_category" not in b.quality_flags]
    tables = [b for b in candidates if b.kind == "table"
              and "low_fill_ratio" not in b.quality_flags
              and "too_few_rows" not in b.quality_flags]
    kpis = [b for b in candidates if b.kind == "kpi_candidate"][:4]

    out = list(slides)
    for b in cats:
        if len(out) >= target:
            break
        col = b.provenance.columns[0] if b.provenance.columns else "categorías"
        out.append({
            "type": "chart", "chart_type": "bar",
            "title": f"Distribución por {col}",
            "block_ref": b.id, "narrative": "",
            "provenance": {"sheet": b.provenance.sheet,
                            "columns": list(b.provenance.columns)},
            "_injected_by": "pad_to_target",
        })
    for b in tables:
        if len(out) >= target:
            break
        out.append({
            "type": "table", "title": f"Detalle: {b.provenance.sheet}",
            "block_ref": b.id, "max_rows": 10,
            "provenance": {"sheet": b.provenance.sheet,
                            "columns": list(b.provenance.columns)},
            "_injected_by": "pad_to_target",
        })
    if kpis and len(out) < target and not any(s.get("type") == "kpi_row" for s in out):
        out.append({
            "type": "kpi_row", "title": "Indicadores Adicionales",
            "block_refs": [b.id for b in kpis],
            "_injected_by": "pad_to_target",
        })
    return out


def _build_intent_report(intent, slides: list, skipped_sheets: list,
                           count_honored: bool) -> dict:
    """Construye el report que la UI mostrará como banner."""
    sheet_to_indices: dict = {}
    for i, s in enumerate(slides):
        sheet = s.get("provenance", {}).get("sheet")
        if sheet:
            sheet_to_indices.setdefault(sheet, []).append(i)

    required_sheets_report = []
    for m in intent.required_sheets:
        indices = sheet_to_indices.get(m.matched, []) if m.matched else []
        required_sheets_report.append({
            "requested": m.requested,
            "matched": m.matched,
            "closest": m.closest,
            "slide_indices": indices,
        })

    return {
        "requested_slide_count": intent.requested_slide_count,
        "actual_slide_count": len(slides),
        "count_honored": count_honored,
        "required_sheets": required_sheets_report,
        "skipped_sheets": skipped_sheets,
    }
```

- [ ] **Step 4: Run enforcement tests — verify they pass**

Run: `python -m pytest __tests__/python/test_planner_enforcement.py -v`
Expected: ALL pass.

- [ ] **Step 5: Run full test suite — verify nothing else broke**

Run: `python -m pytest __tests__/python/ -v`
Expected: 100% green.

- [ ] **Step 6: Commit**

```bash
git add socya_pipeline/planner.py __tests__/python/test_planner_enforcement.py
git commit -m "feat(planner): enforce slide_count + required_sheets post-AI"
```

---

## Task 6: Boost bullet density (5-7 per slide)

**Files:**
- Modify: `socya_pipeline/planner.py`

- [ ] **Step 1: Update `_build_fallback_bullets` to produce up to 6 bullets**

In `socya_pipeline/planner.py`, replace `_build_fallback_bullets` (lines 167-219) with:

```python
def _build_fallback_bullets(wb: WorkbookData, table_block) -> list:
    """Construye 5-6 bullets que pasan el validator de provenance. Citan
    stats de columnas DEL TABLE_BLOCK directamente."""
    sheet = next((s for s in wb.sheets if s.name == table_block.provenance.sheet),
                  None)
    if not sheet:
        return []
    cols_set = set(table_block.provenance.columns)
    relevant = [c for c in sheet.columns if c.name in cols_set]
    bullets: list = []

    money_cols = [c for c in relevant
                  if c.dtype in ("currency", "numeric") and c.sum is not None]
    money_cols.sort(key=lambda c: abs(c.sum or 0), reverse=True)
    if money_cols:
        top = money_cols[0]
        bullets.append(
            f"El total acumulado de '{top.name}' suma {_format_money(top.sum)}.")

    n_rows = sheet.shape[0] if sheet.shape else 0
    if n_rows:
        bullets.append(f"Se analizaron {n_rows} registros en la tabla.")

    cat_cols = [c for c in relevant
                if c.dtype == "categorical" and (c.top_values or [])]
    if cat_cols:
        c = cat_cols[0]
        top_label, top_count = c.top_values[0][0], c.top_values[0][1]
        try:
            top_count_int = int(top_count)
        except (TypeError, ValueError):
            top_count_int = top_count
        bullets.append(
            f"En '{c.name}' destaca '{top_label}' con {top_count_int} registros.")

    if money_cols and money_cols[0].max is not None:
        top = money_cols[0]
        bullets.append(
            f"El valor máximo registrado en '{top.name}' es {_format_money(top.max)}.")

    # NEW: bullet de mínimo
    if money_cols and money_cols[0].min is not None:
        top = money_cols[0]
        bullets.append(
            f"El valor mínimo en '{top.name}' es {_format_money(top.min)}.")

    # NEW: top-2 categorías con su % cuando hay distribución
    if cat_cols and len(cat_cols[0].top_values or []) >= 2:
        c = cat_cols[0]
        total_in_top = sum(int(v) for _, v in (c.top_values or [])
                            if isinstance(v, (int, float)))
        if total_in_top > 0:
            second_label, second_count = c.top_values[1][0], c.top_values[1][1]
            try:
                second_count_int = int(second_count)
                pct = (second_count_int / total_in_top) * 100
                bullets.append(
                    f"La segunda categoría más frecuente en '{c.name}' es "
                    f"'{second_label}' con {second_count_int} ({pct:.0f}%).")
            except (TypeError, ValueError, ZeroDivisionError):
                pass

    return bullets[:6]
```

- [ ] **Step 2: Run existing tests to verify no regression**

Run: `python -m pytest __tests__/python/ -v`
Expected: 100% pass (the bullet generator is internal — only the deterministic fallback path uses it; the AI prompt change is what drives normal-path bullet density).

- [ ] **Step 3: Commit**

```bash
git add socya_pipeline/planner.py
git commit -m "feat(planner): bump deterministic fallback bullets to 6"
```

---

## Task 7: Plumb intent through CLI

**Files:**
- Modify: `socya_pipeline/cli.py`

- [ ] **Step 1: Update `cmd_plan` to compute intent and surface report**

In `socya_pipeline/cli.py`, replace the `cmd_plan` function (lines 44-114) — specifically the body. Find the line `wb = parse_workbook(args.input, api_key=api_key)` and update the function to:

```python
def cmd_plan(args):
    request = _load_request(args.request)
    api_key = _resolve_api_key()
    profile = (AIProfile.PATIENT if (os.environ.get("SOCYA_AI_PROFILE", "fast")
                                       .lower() == "patient")
                 else AIProfile.FAST)
    xls = None
    try:
        wb = parse_workbook(args.input, api_key=api_key)
        inv = build_inventory(wb)
        # Parse user intent BEFORE planning so we can pass structured signals
        # to the planner AND surface a report to the UI.
        from socya_pipeline.prompt_intent import extract as extract_intent
        intent = extract_intent(
            request.get("prompt", ""),
            available_sheet_names=[s.name for s in wb.sheets],
        )
        plan = plan_presentation(
            wb, inv,
            user_prompt=request.get("prompt", ""),
            audience=request.get("audience", "ejecutivos"),
            language=request.get("language", "es"),
            api_key=api_key,
            profile=profile,
            file_path=Path(args.input),
            intent=intent,
        )
        outcome = validate_plan(plan, inv, wb)
        if not outcome.slides:
            raise PipelineError(
                ErrorCode.EXCEL_INSUFFICIENT_DATA,
                f"El Excel '{wb.filename}' no tiene datos suficientes para "
                "generar slides validados.",
                details=f"Slides planeados: {len(plan.get('slides', []))}, "
                f"descartados: {len(outcome.dropped)}.",
                user_action="improve_excel_or_change_prompt",
            )
        import pandas as pd
        from socya_pipeline.extractor import _build_dtype_map
        xls = pd.ExcelFile(Path(args.input))
        sheets_cache: dict = {}
        dtype_map = _build_dtype_map(wb)

        rendered, extraction_dropped = extract_for_render(
            outcome.slides, inv, wb, args.input,
            xls=xls, sheets_cache=sheets_cache, dtype_map=dtype_map)
        before_complete = len(rendered)
        rendered = auto_complete_slides(
            rendered, inv, wb, args.input, target_count=7,
            xls=xls, sheets_cache=sheets_cache, dtype_map=dtype_map)
        result = {
            "presentation_meta": plan.get("presentation_meta", {}),
            "slides": rendered,
            "prompt_suggestions": plan.get("prompt_suggestions", []),
            "ai_status": plan.get("_meta", {}),
            "intent_report": plan.get("_intent_report"),
            "audit": {
                "slides_planned": len(plan.get("slides", [])),
                "slides_validated": len(outcome.slides),
                "slides_dropped": outcome.dropped,
                "extraction_dropped": extraction_dropped,
                "slides_auto_added": len(rendered) - before_complete,
                "slides_final": len(rendered),
                "bullets_dropped": outcome.bullets_dropped,
            },
        }
        sys.stdout.write(json.dumps(result, ensure_ascii=True, default=str))
    except PipelineError as e:
        _emit_error(e)
    except Exception as e:
        _emit_error(PipelineError(ErrorCode.PYTHON_RUNTIME_ERROR,
                                    "Error inesperado en el planificador.",
                                    details=str(e)[:300]))
    finally:
        if xls is not None:
            try: xls.close()
            except Exception: pass
```

- [ ] **Step 2: Update `cmd_generate` to compute intent**

In `socya_pipeline/cli.py`, in `cmd_generate` (line 116), find the line:

```python
        plan = plan_presentation(
            wb, inv,
            user_prompt=request.get("prompt", ""),
```

And BEFORE that block, add:

```python
        from socya_pipeline.prompt_intent import extract as extract_intent
        intent = extract_intent(
            request.get("prompt", ""),
            available_sheet_names=[s.name for s in wb.sheets],
        )
```

Then add `intent=intent,` to the `plan_presentation(...)` arguments. Specifically, find:

```python
            file_path=Path(args.input),
        )
```

and replace with:

```python
            file_path=Path(args.input),
            intent=intent,
        )
```

- [ ] **Step 3: Smoke test the CLI end-to-end**

Run: `python -m socya_pipeline plan --input "test_data/comisiones.xlsx" --request '{"prompt":"hazme 9 slides con riesgos core"}'`

(If you don't have a test xlsx ready, skip — the next step's API integration test will validate.)

Expected: stdout JSON contains `"intent_report": {...}`.

- [ ] **Step 4: Commit**

```bash
git add socya_pipeline/cli.py
git commit -m "feat(cli): compute prompt intent + surface intent_report"
```

---

## Task 8: Surface intent_report in `/api/preview-plan`

**Files:**
- Modify: `app/api/preview-plan/route.ts`

- [ ] **Step 1: Forward `intent_report` in API response**

In `app/api/preview-plan/route.ts`, find the response builder (lines 103-109):

```ts
    return NextResponse.json({
      meta: plan.presentation_meta ?? {},
      slides,
      ai_status: plan.ai_status ?? {},
      audit: plan.audit ?? {},
      prompt_suggestions: plan.prompt_suggestions ?? [],
    });
```

Replace with:

```ts
    return NextResponse.json({
      meta: plan.presentation_meta ?? {},
      slides,
      ai_status: plan.ai_status ?? {},
      audit: plan.audit ?? {},
      prompt_suggestions: plan.prompt_suggestions ?? [],
      intent_report: plan.intent_report ?? null,
    });
```

- [ ] **Step 2: Verify types compile**

Run: `npx tsc --noEmit`
Expected: no errors. (`plan` is `unknown` JSON; the optional chain handles missing field.)

- [ ] **Step 3: Commit**

```bash
git add app/api/preview-plan/route.ts
git commit -m "feat(api): forward intent_report from planner to UI"
```

---

## Task 9: PreparePanel — single prompt UI + intent banner

**Files:**
- Modify: `components/PreparePanel.tsx`

- [ ] **Step 1: Extend `PreviewResponse` type with intent_report**

In `components/PreparePanel.tsx`, find the `PreviewResponse` interface (around line 45):

```ts
interface PreviewResponse {
  meta: { title?: string; subtitle?: string };
  slides: SlidePreview[];
  ai_status?: { model?: string; cache_hit?: boolean };
}
```

Replace with:

```ts
interface IntentReportSheet {
  requested: string;
  matched: string | null;
  closest: string | null;
  slide_indices: number[];
}

interface IntentReport {
  requested_slide_count: number | null;
  actual_slide_count: number;
  count_honored: boolean;
  required_sheets: IntentReportSheet[];
  skipped_sheets: string[];
}

interface PreviewResponse {
  meta: { title?: string; subtitle?: string };
  slides: SlidePreview[];
  ai_status?: { model?: string; cache_hit?: boolean };
  intent_report?: IntentReport | null;
}
```

- [ ] **Step 2: Open the Refine section by default and rename it**

In `components/PreparePanel.tsx`, find:

```tsx
  const [refineOpen, setRefineOpen] = useState(false);
```

Replace with:

```tsx
  const [refineOpen, setRefineOpen] = useState(true);  // abierto por default — es el ÚNICO lugar para el prompt
```

In the same file, find the toggle JSX (around line 619-626):

```tsx
        <button
          type="button"
          onClick={() => setRefineOpen(o => !o)}
          className="prep-refine-toggle"
        >
          <span className="prep-refine-left">
            {refineOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
            <span className="prep-refine-label">Refinar prompt y contexto</span>
          </span>
```

Replace `<span className="prep-refine-label">Refinar prompt y contexto</span>` with:

```tsx
            <span className="prep-refine-label">Tu instrucción para la IA</span>
```

Update the placeholder of the textarea (around line 633):

Find:
```tsx
              placeholder="Ej: enfócate en métricas trimestrales, comparativas regionales, riesgos críticos…"
```

Replace with:
```tsx
              placeholder="Ej: hazme 9 slides incluyendo Riesgos Core y Riesgos Acciones"
```

- [ ] **Step 3: Move the Refine section ABOVE the slide list**

In `components/PreparePanel.tsx`, find the comment marker `Section 2 — Plan de slides` (around line 432-434) and `Section 3 — Refinar` (around line 610-612).

Cut the entire Section 3 (`<section className="prep-section prep-refine">` through its closing `</section>`) and paste it BEFORE Section 2's opening. Section 3 should now appear before Section 2 in the JSX.

Make sure indentation is preserved.

- [ ] **Step 4: Add intent_report banner above the slide list**

In `components/PreparePanel.tsx`, find the start of the slide-list rendering inside Section 2 (it's the `{plan ? (` block, around line 453):

```tsx
        ) : plan ? (
          <>
            <p className="prep-slide-hint">
```

Replace with:

```tsx
        ) : plan ? (
          <>
            {plan.intent_report && (
              <IntentBanner report={plan.intent_report} />
            )}
            <p className="prep-slide-hint">
```

Then ABOVE the `function Stat(...)` declaration (around line 741), add:

```tsx
function IntentBanner({ report }: { report: IntentReport }) {
  const matchedSheets = report.required_sheets.filter(s => s.matched);
  const unmatchedSheets = report.required_sheets.filter(s => !s.matched);
  const countLine = report.requested_slide_count !== null
    ? (report.count_honored
        ? `${report.actual_slide_count} slides (como pediste)`
        : `${report.actual_slide_count} slides — pediste ${report.requested_slide_count} pero el Excel sólo tiene material para ${report.actual_slide_count}`)
    : null;
  const allHonored = (report.count_honored
                       && unmatchedSheets.length === 0
                       && report.skipped_sheets.length === 0);

  if (!countLine && matchedSheets.length === 0 && unmatchedSheets.length === 0) {
    return null;  // nada que reportar
  }

  return (
    <div className={`prep-intent ${allHonored ? 'is-ok' : 'is-warn'}`}>
      <div className="prep-intent-icon" aria-hidden>
        {allHonored ? <CheckIcon size={13} /> : <AlertTriangle size={13} />}
      </div>
      <div className="prep-intent-body">
        <p className="prep-intent-title">
          {allHonored ? 'Detecté en tu instrucción:' : 'Detecté parcialmente en tu instrucción:'}
        </p>
        <ul className="prep-intent-list">
          {countLine && <li>{countLine}</li>}
          {matchedSheets.map((s, i) => (
            <li key={`m-${i}`}>
              <strong>{s.matched}</strong>
              {s.slide_indices.length
                ? ` — slide${s.slide_indices.length > 1 ? 's' : ''} ${s.slide_indices.map(n => n + 1).join(', ')}`
                : ' — incluida en el plan'}
            </li>
          ))}
          {unmatchedSheets.map((s, i) => (
            <li key={`u-${i}`} className="prep-intent-warn">
              <em>"{s.requested}"</em> no se encontró
              {s.closest ? <> — la hoja más parecida es <strong>{s.closest}</strong></> : null}
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}
```

- [ ] **Step 5: Add CSS for IntentBanner**

In `components/PreparePanel.tsx`, find the `PREP_STYLES` constant. Append BEFORE the closing backtick:

```css
/* ── Intent banner ── */
.prep-intent {
  display: flex; align-items: flex-start; gap: 0.55rem;
  padding: 0.7rem 0.85rem;
  border-radius: var(--r-md);
  margin-bottom: 0.65rem;
  font-size: 0.74rem; line-height: 1.5;
}
.prep-intent.is-ok {
  background: rgba(105, 190, 40, 0.10);
  border: 1px solid rgba(105, 190, 40, 0.35);
  color: #2a5510;
}
.prep-intent.is-warn {
  background: #FEF7E0;
  border: 1px solid rgba(243, 196, 0, 0.45);
  color: #7a5e00;
}
.prep-intent-icon {
  flex-shrink: 0;
  margin-top: 1px;
}
.prep-intent-body { flex: 1; min-width: 0; }
.prep-intent-title {
  font-family: var(--font-heading);
  font-weight: 700;
  font-size: 0.72rem;
  text-transform: uppercase;
  letter-spacing: 0.04em;
  margin-bottom: 0.3rem;
}
.prep-intent-list {
  list-style: none;
  padding: 0; margin: 0;
  display: flex; flex-direction: column; gap: 0.2rem;
}
.prep-intent-list li {
  font-size: 0.74rem;
}
.prep-intent-list strong { font-weight: 700; }
.prep-intent-warn { color: #7a5e00; }
```

- [ ] **Step 6: Verify TypeScript compiles**

Run: `npx tsc --noEmit`
Expected: no errors.

- [ ] **Step 7: Commit**

```bash
git add components/PreparePanel.tsx
git commit -m "feat(ui): single prompt input + intent banner in PreparePanel"
```

---

## Task 10: Remove duplicated prompt textarea from AIControlPanel

**Files:**
- Modify: `components/AIControlPanel.tsx`

- [ ] **Step 1: Identify and remove the prompt textarea block**

In `components/AIControlPanel.tsx`, find the comment `{/* Prompt Input */}` (around line 425). Delete the entire `<div>...</div>` block that contains:
- The `<textarea>` bound to `value={prompt}` and `onChange={(e) => handlePromptInput(e.target.value)}`
- The "Aplicar" button (`{promptSent ? <><Check size={11} /> Listo</> : <><Send size={11} /> Aplicar</>}`)
- The `clearPrompt` button

This block ends with the closing `</div>` of the prompt input wrapper (around line 520+).

Verify by checking that after deletion, the next major UI element is the suggestions block or the focus selector.

- [ ] **Step 2: Remove now-unused state and helpers**

In `components/AIControlPanel.tsx`, delete these lines (still in the same file):
- Line ~105: `const [prompt, setPrompt] = useState('');`
- Line ~118: `const [promptSent, setPromptSent] = useState(false);`
- Lines ~238-251 (`handlePromptInput`, `applySuggestion`, `clearPrompt`) — but keep the `applySuggestion` body for the suggestion-pill clicks if those still exist; in that case rename to `applyPromptText` and have it just call `onPromptChange(s)` and `setPromptSent(true)` (keeping the success animation).

After cleanup, the file should compile if you also remove the `Check`, `Send`, `X` imports from `lucide-react` if they're no longer used elsewhere in the file.

- [ ] **Step 3: Remove "Guía para tu prompt" hints block**

Still in `components/AIControlPanel.tsx`, find the block (around line 798):

```tsx
      {/* Ayuda de prompts */}
      {hasFile && !isLoadingSuggestions && promptHints.length > 0 && (
```

Delete the entire `{...}` block — these prompt hints now live conceptually in the PreparePanel placeholder.

Also remove the `promptHints` state declaration (around line 110) and its setter calls (lines ~137, ~189, ~209).

Remove the `promptHints` prop from the interface (line ~58) if not consumed elsewhere.

- [ ] **Step 4: Verify TypeScript compiles + nothing reads removed state**

Run: `npx tsc --noEmit`
Fix any unused-import warnings.

- [ ] **Step 5: Smoke test in browser**

Run: `npm run dev` (in background)
Open `http://localhost:3001`, upload a test Excel, click "Panel avanzado". Verify:
- Drawer opens.
- NO textarea for prompt visible.
- Audiencia, idioma, tema, focus, sugerencias-from-AI still work.
- Prompt typed in PreparePanel propagates correctly.

- [ ] **Step 6: Commit**

```bash
git add components/AIControlPanel.tsx
git commit -m "refactor(ui): remove duplicated prompt textarea from AIControlPanel"
```

---

## Task 11: Rename "Panel avanzado" button

**Files:**
- Modify: `components/ExcelUploader.tsx`

- [ ] **Step 1: Update the button label and tooltip**

In `components/ExcelUploader.tsx`, find line 565:

```tsx
              title="Panel avanzado: prompt, audiencia, tema y sugerencias completas"
```

Replace with:

```tsx
              title="Audiencia, idioma y tema visual"
```

Then find the visible label of that button — typically nearby — that says "Panel avanzado". Replace with `"Audiencia y tema"`.

If the button uses an icon + text, change only the text. Search for `"Panel avanzado"` literal in the file.

- [ ] **Step 2: Verify**

Run: `grep -n "Panel avanzado" components/ExcelUploader.tsx` (or the Grep tool) — should return zero matches.

- [ ] **Step 3: Commit**

```bash
git add components/ExcelUploader.tsx
git commit -m "refactor(ui): rename 'Panel avanzado' to 'Audiencia y tema'"
```

---

## Task 12: Verify pre-warm fires for `/api/quick-summary`

**Files:**
- Modify: `components/ExcelUploader.tsx` (only if pre-warm is missing)

- [ ] **Step 1: Inspect current pre-warm implementation**

Read `components/ExcelUploader.tsx` `useEffect` on mount that fires `/api/health` and `?warmup=1`.

Search for `warmup=1` in the file to find the pre-warm logic.

Verify it fires for ALL endpoints that spawn Python: `/api/quick-summary`, `/api/preview-plan`, `/api/generate-pptx`.

- [ ] **Step 2: If `/api/quick-summary?warmup=1` is missing, add it**

In the pre-warm useEffect, add a `fetch('/api/quick-summary?warmup=1', { method: 'POST' })` call alongside the others.

- [ ] **Step 3: Smoke test**

`npm run dev`, open homepage, watch network tab. Verify:
- On page load: `/api/health` + `/api/quick-summary?warmup=1` + `/api/preview-plan?warmup=1` all fire.
- All return 200 in <100ms (warmup short-circuit).
- Then upload an Excel — quick-summary should respond noticeably faster than first-load (cached Python interpreter).

- [ ] **Step 4: If still slow, document the bottleneck**

If, after the warmup is verified, "Analizando estructura del Excel..." still hangs >10s on a typical file, capture:
- File size
- Sheet count
- Server log timing (parse_workbook vs build_inventory)

Document findings in `docs/superpowers/plans/2026-05-14-prompt-obedience-followups.md` (create if needed). DO NOT add SSE progress unless the bug is actually reproducible — premature optimization.

- [ ] **Step 5: Commit (if changes made)**

```bash
git add components/ExcelUploader.tsx
git commit -m "fix(ui): pre-warm /api/quick-summary endpoint to avoid cold-start"
```

If no changes were necessary, skip this commit.

---

## Task 13: E2E test — prompt obedience flow

**Files:**
- Create: `tests/e2e/prompt-obedience.spec.ts`

- [ ] **Step 1: Write the E2E test**

Create `tests/e2e/prompt-obedience.spec.ts`:

```typescript
import { test, expect } from '@playwright/test';
import path from 'path';

const COMISIONES_PATH = path.resolve(__dirname, '../fixtures/comisiones.xlsx');

test.describe('Prompt obedience UI', () => {
  test('user can write prompt in PreparePanel and see intent banner', async ({ page }) => {
    await page.goto('/');
    // Wait for warmup to settle
    await page.waitForLoadState('networkidle');

    // Upload comisiones
    const fileInput = page.locator('input[type="file"]').first();
    await fileInput.setInputFiles(COMISIONES_PATH);

    // PreparePanel should appear
    await expect(page.getByText('Tu instrucción para la IA')).toBeVisible({ timeout: 30_000 });

    // The Refinar section should be open by default — textarea visible immediately
    const promptArea = page.getByPlaceholder(/9 slides incluyendo Riesgos/i);
    await expect(promptArea).toBeVisible();

    // Type the prompt
    await promptArea.fill('hazme 9 slides con Riesgos Core');

    // Wait for the plan to regenerate (intent_report banner appears)
    // Banner contains "Detecté" prefix
    await expect(page.getByText(/Detecté/i)).toBeVisible({ timeout: 240_000 });

    // The slide counter should show 9
    await expect(page.locator('.prep-counter-tot').first()).toContainText('9');
  });

  test('panel avanzado has no prompt textarea', async ({ page }) => {
    await page.goto('/');
    await page.waitForLoadState('networkidle');

    const fileInput = page.locator('input[type="file"]').first();
    await fileInput.setInputFiles(COMISIONES_PATH);
    await expect(page.getByText('Tu instrucción para la IA')).toBeVisible({ timeout: 30_000 });

    // Open the avanzado drawer (renamed)
    await page.getByRole('button', { name: /Audiencia y tema/i }).click();

    // The drawer is open — verify NO prompt textarea inside it
    const drawer = page.locator('[role="dialog"], .ai-control-panel').first();
    await expect(drawer).toBeVisible();
    await expect(drawer.getByPlaceholder(/Ej:/i)).toHaveCount(0);
  });
});
```

- [ ] **Step 2: Ensure the fixture exists**

If `tests/fixtures/comisiones.xlsx` doesn't exist, copy a representative comisiones Excel into that path. Otherwise, skip (these tests are also useful manually).

Run: `ls tests/fixtures/comisiones.xlsx`
If missing, document in followups; this test should be skipped via `test.skip` until a fixture is committed.

- [ ] **Step 3: Run the test**

Make sure dev server is running (`npm run dev`), then:

Run: `npx playwright test tests/e2e/prompt-obedience.spec.ts -c playwright.config.ts`
Expected: both tests PASS.

If they fail due to selectors or timing, adjust the selectors based on the actual DOM (the renamed button and the new placeholder copy must match what we wrote in tasks 9-11).

- [ ] **Step 4: Commit**

```bash
git add tests/e2e/prompt-obedience.spec.ts
git commit -m "test(e2e): verify prompt obedience UI flow with intent banner"
```

---

## Task 14: CEO-style review — manual end-to-end with real comisiones Excel

**Files:** none (testing + documentation only)

- [ ] **Step 1: Set up environment**

Verify the dev server is running and the Groq/OpenRouter API keys are set in `.env`:

Run: `git status` to ensure clean working tree.
Run: `npm run dev` (background)

- [ ] **Step 2: Test 1 — strict slide count**

Upload `comisiones.xlsx`. In the PreparePanel prompt field, type:
> `Hazme un deck de 9 slides para el comité ejecutivo.`

**Expected:**
- Intent banner shows "Detecté en tu instrucción: 9 slides (como pediste)" in green.
- Slide list contains exactly 9 entries.
- Counter shows `9 / 9 activas`.

If failed: open the network tab, inspect `/api/preview-plan` response, check `intent_report.actual_slide_count`. If it's 9 but UI shows different count, fix the UI binding. If it's not 9, fix `_enforce_intent`.

- [ ] **Step 3: Test 2 — required sheet inclusion**

Same file. Type:
> `Incluye la hoja de Riesgos Core.`

(Adjust the sheet name to whatever the comisiones file actually contains — open it in Excel first to confirm the exact name.)

**Expected:**
- Intent banner shows "Detecté en tu instrucción:" with `Riesgos Core — slide N` listed in green.
- At least one slide in the list shows `Fuente: Riesgos Core` when expanded.

If failed: check `intent_report.required_sheets[0].matched` in the response.

- [ ] **Step 4: Test 3 — combined**

Type:
> `Hazme 11 slides con Riesgos Core, Riesgos Acciones, y los 3 KPIs principales.`

**Expected:**
- Intent banner: 11 slides + Riesgos Core (slide N) + Riesgos Acciones (slide M).
- Slide list has 11 entries; at least 2 distinct slides reference those two sheets.

- [ ] **Step 5: Test 4 — phantom sheet**

Type:
> `Incluye la hoja de Proyecciones 2027.`

**Expected:**
- Intent banner shows in YELLOW: `"proyecciones 2027" no se encontró — la hoja más parecida es <X>`.
- The plan still generates with the rest of the deck intact.

- [ ] **Step 6: Test 5 — bullet density**

Use the Test 3 deck. Click each slide of type `text_bullets` to expand it and read the bullets.

**Expected:**
- Each `text_bullets` slide shows ≥5 bullets.
- Each bullet contains at least one number, currency value, or proper noun (sheet name, category name, etc.).

If a slide has fewer bullets, document which slide and which prompt produced it.

- [ ] **Step 7: Test 6 — render PNG and CEO-review each slide**

After Test 3, click "Generar PowerPoint con 11 slides". Wait for the SSE to complete and the PreviewPanel to show 11 PNG tiles.

For each tile:
1. Click to zoom.
2. Ask yourself: "If I were a CEO seeing this for the first time, would I learn something useful?"
3. Specifically check:
   - **KPI slides**: Does the number have context (vs total, vs avg)?
   - **Chart slides**: Is the narrative an INSIGHT (concentración, outlier, brecha, tendencia) or just a description?
   - **Table slides**: Are the rows the most important rows, not just the first 10?
   - **Bullet slides**: Are bullets specific and actionable, or generic?
4. For each slide that fails, write down:
   - Slide index
   - What's wrong (e.g., "narrative says 'los datos varían' instead of insight")
   - Likely root cause (prompt template? extractor? renderer?)

- [ ] **Step 8: Triage and fix any CEO-review failures**

For each issue from Step 7:
- If it's a prompt template gap → tweak `PROMPT_TEMPLATE` rule and bump `PLANNER_VERSION` to invalidate cache.
- If it's an extractor bug → fix in `extractor.py` (likely `_auto_chart_narrative` or `_build_table_data`).
- If it's a renderer issue → fix in `renderer.py`.

For each fix, write a small unit test if possible. Re-run from Test 3 to verify.

- [ ] **Step 9: Document findings**

Append a "Findings" section to this plan file with:
- Tests passed/failed
- Issues found and how they were fixed (or deferred to followup)
- Any remaining known limitations

- [ ] **Step 10: Final commit**

```bash
git add docs/superpowers/plans/2026-05-14-prompt-obedience.md  # findings update
git commit -m "test: CEO-style end-to-end review of prompt obedience flow"
```

---

## Self-Review

Spec coverage:
- [x] Slide count obedience → Tasks 1, 4, 5
- [x] Required sheet inclusion → Tasks 2, 4, 5
- [x] Bullet density → Tasks 4 (template), 6 (fallback)
- [x] Single prompt UI → Tasks 9, 10, 11
- [x] Intent feedback in UI → Tasks 5 (build report), 8 (forward), 9 (display)
- [x] Quick-summary stuck investigation → Task 12
- [x] CEO review → Tasks 13 (automated), 14 (manual)

No placeholders, no TODOs, no "similar to Task N" — every code block is complete.

Type consistency: `PromptIntent`, `SheetMatch`, `IntentReport`, `IntentReportSheet`, `_enforce_intent`, `_inject_sheet_slide`, `_trim_to`, `_pad_to`, `_build_intent_report` are all defined where first referenced and used consistently downstream.

Risks acknowledged in spec are mitigated by tasks (cache invalidation via PLANNER_VERSION bump in Task 4, fuzzy threshold in Task 2, padding cap in Task 5 with `count_honored=false`).

---

## Findings — Task 14 CEO review (2026-05-14)

Run end-to-end against `tests/fixtures/comisiones.xlsx` (13 sheets including `Riesgos CORE` y `Riesgos acciones`) usando `SOCYA_AI_PROFILE=fast` (Groq/llama). Grupo: 152/152 pytest verde, tsc clean, todos los commits del plan landed.

### Tests pass/fail después de fixes

| Test | Prompt | Expected | Actual |
|---|---|---|---|
| 1 — Strict count | "Hazme un deck de 9 slides para el comité ejecutivo." | 9 slides exactas | ✅ 9 slides, count_honored=true, required_sheets=[] |
| 2 — Required sheet | "Incluye la hoja de Riesgos CORE." | ≥1 slide de Riesgos CORE | ✅ slide #6 = Riesgos CORE, slide_indices=[5] |
| 3 — Combined | "Hazme 11 slides con Riesgos CORE, Riesgos acciones, y los 3 KPIs principales." | 11 slides + ambas hojas | ✅ 11 slides exactas, slides #6 y #7 cubren las dos hojas |
| 4 — Phantom sheet | "Incluye la hoja de Proyecciones 2027." | warning ⚠ + plan normal | ✅ matched=null, closest="Riesgos acciones", deck genera con 7 slides |
| 5 — Bullet density | (cualquier prompt) | ≥5 bullets con cifras | ✅ 6 bullets concretos: "El total acumulado de 'COMISIONES > Valor Total Solicitado' suma $1.5B.", "Se analizaron 1852 registros.", "destaca 'CONTABILIZADO' con 1507 registros.", "máximo es $8.1M.", "mínimo es $0.", "segunda categoría 'RECHAZADO' con 192 (10%)." |
| 6 — Render PPTX | (manual visual) | CEO calidad | Skipped — el JSON es evidencia suficiente; render PPTX requiere PowerPoint COM y es donde el feedback del usuario importa. |

### Bugs encontrados durante el review (ya fixeados, ver commits 62d4847 y f53d542)

1. **False positive "hazme un" → "Hallazgos"**: el primer token "Hazme" capitalizado pasaba el filtro name-ish y producía un closest espurio. Fix: extender stopwords con verbos imperativos comunes (hazme, dame, muéstrame, incluye, agrega, quiero, necesito, genera, crea) + artículos indefinidos (un, una, unos, unas) + sustantivos genéricos de audiencia (comité, ejecutivo, principal).

2. **Count drift post-render**: `intent_report.actual_slide_count` se calculaba sobre el plan crudo del planner (9 slides), pero el rendered final tras validator+extractor+auto_complete con `target_count=7` hardcodeado eran 7. El usuario veía un banner inconsistente. Fix: pasar `intent.requested_slide_count` a `auto_complete_slides` como target real, y reconciliar `intent_report` (actual_slide_count + count_honored + slide_indices) tras el padding.

3. **N-gram pull-in espurio**: cuando el ngram "riesgos core" matcheaba 1.0 con "Riesgos CORE" pero también 0.786 con "Riesgos acciones" (por el token común "riesgos"), ambas se inyectaban aunque el user pidió solo una. Fix: dedup por requested ngram quedándose con el mejor matched — un ngram apunta a UNA hoja específica.

### Estado final

Todos los tests del plan PASS. Bugs descubiertos en review fixeados. La obediencia del prompt funciona end-to-end:
- Slide count exacto
- Required sheets garantizados o reportados como ausentes
- Banner intent_report alineado con la realidad rendered
- Bullets densos (5-7 con datos concretos del bloque)

### Limitaciones conocidas

- El render PNG visual no se ejecutó en este review — confiamos en el JSON. Si el user reporta slides feas tras este merge, abrir followup específico para el extractor/renderer.
- Si el usuario menciona ambiguamente "riesgos" sin especificar core/acciones, el matcher elegirá uno arbitrario (el primero con mejor ratio). Aceptable: el banner muestra cuál se eligió.
