# Bullet Density + Insight Bullets — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task.

**Goal:** Garantizar ≥5 bullets ricos en cada slide de hallazgos + agregar 1 hallazgos por hoja importante intercalado con sus charts/tablas.

**Architecture:** Top-up determinístico de bullets en `extract_for_render` cuando AI devuelve <5. Generators de insights (Pareto, outlier, brecha) en planner. `auto_complete_slides` añade hallazgos por hoja cubierta.

**Tech Stack:** Python 3.12-3.14 (stdlib + pandas), pytest. Sin cambios en frontend.

**Spec:** [`docs/superpowers/specs/2026-05-14-bullet-density-design.md`](../specs/2026-05-14-bullet-density-design.md)

---

## File Structure

| Archivo | Tipo | Responsabilidad |
|---|---|---|
| `socya_pipeline/planner.py` | MODIFY | Renombrar `_build_fallback_bullets` → `_build_insight_bullets`, añadir generators Pareto/outlier/brecha, devolver tuplas (angle, bullet) |
| `socya_pipeline/extractor.py` | MODIFY | Top-up de bullets en text_bullets branch, helper `_add_hallazgos_per_sheet` invocado desde `auto_complete_slides` |
| `__tests__/python/test_bullet_density.py` | NEW | Tests de top-up, dedup, mix de angles, auto-add por hoja |

---

## Task 1: Refactor `_build_fallback_bullets` → `_build_insight_bullets` (returns tuples)

**Files:** Modify `socya_pipeline/planner.py`

- [ ] **Step 1: Read current `_build_fallback_bullets` to know what callers expect.**

Run: read `socya_pipeline/planner.py` lines 167-219. Confirm the function returns `list[str]` and is called from `_deterministic_plan_fallback` only.

- [ ] **Step 2: Refactor signature to return tuples.**

Replace `_build_fallback_bullets` with:

```python
def _build_insight_bullets(wb: WorkbookData, table_block,
                            max_candidates: int = 8) -> list:
    """Generates up to `max_candidates` insight bullets from the block's
    sheet stats. Returns a list of (angle, bullet_text) tuples so callers
    can pick a diverse mix.

    Each generator is best-effort and degrades to no-op if the block lacks
    the required signal (no money column, no categorical column, etc).

    Bullets are designed to pass validator._bullet_has_provenance:
    - All cited numbers come from col.sum/mean/min/max/value or top_values
    - All cited names come from col.top_values labels or column names
    """
    sheet = next((s for s in wb.sheets
                  if s.name == table_block.provenance.sheet), None)
    if not sheet:
        return []
    cols_set = set(table_block.provenance.columns)
    relevant = [c for c in sheet.columns if c.name in cols_set]

    out: list = []
    money_cols = [c for c in relevant
                  if c.dtype in ("currency", "numeric") and c.sum is not None]
    money_cols.sort(key=lambda c: abs(c.sum or 0), reverse=True)
    cat_cols = [c for c in relevant
                if c.dtype == "categorical" and (c.top_values or [])]

    # Generators (each appends 0 or 1 bullet to `out`)
    _gen_total(out, money_cols)
    _gen_count(out, sheet)
    _gen_top_cat(out, cat_cols)
    _gen_max(out, money_cols)
    _gen_min(out, money_cols)
    _gen_second_cat(out, cat_cols)
    _gen_pareto(out, cat_cols)
    _gen_outlier(out, cat_cols)

    return out[:max_candidates]


def _gen_total(out: list, money_cols: list) -> None:
    if not money_cols:
        return
    top = money_cols[0]
    out.append(("total",
        f"El total acumulado de '{top.name}' suma {_format_money(top.sum)}."))


def _gen_count(out: list, sheet) -> None:
    n = sheet.shape[0] if sheet.shape else 0
    if n:
        out.append(("count", f"Se analizaron {n} registros en la tabla."))


def _gen_top_cat(out: list, cat_cols: list) -> None:
    if not cat_cols:
        return
    c = cat_cols[0]
    label, count = c.top_values[0][0], c.top_values[0][1]
    try:
        ci = int(count)
    except (TypeError, ValueError):
        ci = count
    out.append(("top_cat",
        f"En '{c.name}' destaca '{label}' con {ci} registros."))


def _gen_max(out: list, money_cols: list) -> None:
    if not money_cols or money_cols[0].max is None:
        return
    top = money_cols[0]
    out.append(("max",
        f"El valor máximo registrado en '{top.name}' es {_format_money(top.max)}."))


def _gen_min(out: list, money_cols: list) -> None:
    if not money_cols or money_cols[0].min is None:
        return
    top = money_cols[0]
    out.append(("min",
        f"El valor mínimo en '{top.name}' es {_format_money(top.min)}."))


def _gen_second_cat(out: list, cat_cols: list) -> None:
    if not cat_cols or len(cat_cols[0].top_values or []) < 2:
        return
    c = cat_cols[0]
    total = sum(int(v) for _, v in (c.top_values or [])
                if isinstance(v, (int, float)))
    if total <= 0:
        return
    label, count = c.top_values[1][0], c.top_values[1][1]
    try:
        ci = int(count)
        pct = (ci / total) * 100
        out.append(("second_cat",
            f"La segunda categoría más frecuente en '{c.name}' es "
            f"'{label}' con {ci} ({pct:.0f}%)."))
    except (TypeError, ValueError, ZeroDivisionError):
        pass


def _gen_pareto(out: list, cat_cols: list) -> None:
    """Pareto top-3: si los 3 más frecuentes concentran ≥60% del total,
    bullet de concentración."""
    if not cat_cols or len(cat_cols[0].top_values or []) < 3:
        return
    c = cat_cols[0]
    counts = [int(v) for _, v in (c.top_values or [])
              if isinstance(v, (int, float))]
    if not counts or sum(counts) <= 0:
        return
    try:
        share = insights.pareto_share(counts, top_n=3)
    except Exception:
        return
    if share is None or share < 0.6:
        return
    top3_labels = [t[0] for t in c.top_values[:3]]
    out.append(("pareto",
        f"El top 3 de '{c.name}' ({', '.join(str(l) for l in top3_labels)}) "
        f"concentra el {share*100:.0f}% del total — riesgo de dependencia."))


def _gen_outlier(out: list, cat_cols: list) -> None:
    """Outlier alto vía IQR sobre los counts de la columna categórica."""
    if not cat_cols or len(cat_cols[0].top_values or []) < 4:
        return
    c = cat_cols[0]
    counts = [int(v) for _, v in (c.top_values or [])
              if isinstance(v, (int, float))]
    if len(counts) < 4:
        return
    try:
        outliers = insights.iqr_outliers(counts, k=1.5)
    except Exception:
        return
    high = [(i, v) for i, v, side in outliers if side == "high"]
    if not high:
        return
    # Pick the highest outlier
    high.sort(key=lambda t: -t[1])
    idx, val = high[0]
    if idx >= len(c.top_values):
        return
    label = c.top_values[idx][0]
    median = sorted(counts)[len(counts) // 2]
    if median <= 0:
        return
    ratio = val / median
    out.append(("outlier_high",
        f"'{label}' destaca con {val} registros — {ratio:.1f}x la mediana, "
        f"un caso atípico marcado en '{c.name}'."))
```

- [ ] **Step 3: Update caller `_deterministic_plan_fallback`.**

Find the line `bullets = _build_fallback_bullets(wb, b)` (currently around line 297). Replace the assignment and slide construction with:

```python
        # _build_insight_bullets now returns (angle, text) tuples — extraer texto
        candidates = _build_insight_bullets(wb, b, max_candidates=8)
        bullets = [text for (_, text) in candidates[:6]]
        if bullets:
            slides.append({
                "type": "text_bullets",
                "title": "Hallazgos Principales",
                "supports_block": b.id,
                "bullets": bullets,
            })
```

- [ ] **Step 4: Run existing tests — verify nothing broke.**

Run: `python -m pytest __tests__/python/ -q --tb=short`
Expected: 152/152 still green.

- [ ] **Step 5: Commit.**

```bash
git add socya_pipeline/planner.py
git commit -m "refactor(planner): _build_insight_bullets returns (angle, text) tuples + Pareto/outlier generators"
```

---

## Task 2: Bullet top-up in extractor (TDD)

**Files:** Modify `socya_pipeline/extractor.py`. Create `__tests__/python/test_bullet_density.py`.

- [ ] **Step 1: Write failing tests.**

Create `__tests__/python/test_bullet_density.py`:

```python
"""Tests para top-up de bullets y auto-add de hallazgos por hoja."""
import pytest
from socya_pipeline.extractor import _topup_bullets, _add_hallazgos_per_sheet
from socya_pipeline.parser import WorkbookData, SheetData, ColumnData
from socya_pipeline.inventory import build_inventory


def _make_wb_with_money_and_categories(sheet_name="Comisiones"):
    return WorkbookData(filename="test.xlsx", sheets=[
        SheetData(
            name=sheet_name,
            shape=(100, 3),
            fill_ratio=0.95,
            columns=[
                ColumnData(name="Estado", dtype="categorical",
                            n_unique=4, fill_ratio=1.0,
                            samples=["A", "B", "C", "D"],
                            top_values=[("CONTABILIZADO", 60),
                                          ("RECHAZADO", 25),
                                          ("PENDIENTE", 10),
                                          ("ANULADO", 5)]),
                ColumnData(name="Monto", dtype="currency",
                            n_unique=80, fill_ratio=1.0,
                            samples=[1000.0, 2000.0, 5000.0],
                            sum=500000.0, mean=5000.0,
                            min=100.0, max=50000.0),
                ColumnData(name="Ciudad", dtype="categorical",
                            n_unique=3, fill_ratio=0.9,
                            samples=["BOG", "MED", "CAL"],
                            top_values=[("BOGOTÁ", 70),
                                          ("MEDELLÍN", 20),
                                          ("CALI", 10)]),
            ],
            first_rows=[],
        )
    ])


# ─── Top-up ─────────────────────────────────────────────────────────

def test_topup_when_ai_returned_too_few():
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")

    # AI gave only 2 bullets
    existing = ["El total es $500K.", "Hay 100 registros."]
    out = _topup_bullets(existing, table, wb)
    assert len(out) >= 5
    # Original bullets preserved
    assert out[0] == "El total es $500K."
    assert out[1] == "Hay 100 registros."


def test_topup_when_ai_already_full():
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")
    existing = [f"Bullet {i}" for i in range(6)]
    out = _topup_bullets(existing, table, wb)
    # No change when ≥5 already
    assert out == existing


def test_topup_dedup_against_existing():
    """Si AI ya dijo 'total acumulado de Monto', no agregamos otro de
    la misma columna sumando lo mismo."""
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")
    existing = ["El total acumulado de 'Monto' suma $500K."]
    out = _topup_bullets(existing, table, wb)
    # Only ONE bullet about total of Monto
    total_bullets = [b for b in out if "total" in b.lower() and "monto" in b.lower()]
    assert len(total_bullets) <= 1


def test_topup_mix_angles_not_repeats_same_generator():
    """5 bullets agregados deben venir de generators distintos."""
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")
    out = _topup_bullets([], table, wb)
    assert len(out) >= 5
    # Heuristic: bullets shouldn't all start with the same word
    starting_words = [b.split()[0].lower() for b in out if b]
    assert len(set(starting_words)) >= 3


def test_topup_caps_at_max_bullets():
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")
    out = _topup_bullets([], table, wb)
    assert len(out) <= 6


# ─── Auto-add hallazgos per sheet ───────────────────────────────────

def test_auto_complete_adds_hallazgos_per_sheet():
    wb = _make_wb_with_money_and_categories("HojaA")
    # Add second sheet
    wb.sheets.append(SheetData(
        name="HojaB", shape=(50, 2), fill_ratio=0.9,
        columns=[
            ColumnData(name="Categoria", dtype="categorical",
                        n_unique=3, fill_ratio=1.0,
                        samples=["X", "Y", "Z"],
                        top_values=[("X", 30), ("Y", 15), ("Z", 5)]),
            ColumnData(name="Valor", dtype="numeric",
                        n_unique=20, fill_ratio=0.9,
                        samples=[10.0, 20.0],
                        sum=1000.0, mean=20.0, min=5.0, max=50.0),
        ],
        first_rows=[],
    ))
    blocks = build_inventory(wb)
    blocks_by_id = {b.id: b for b in blocks}

    rendered = [
        {"type": "title", "data": {"title": "T"}},
        {"type": "chart", "block_ref": next(b.id for b in blocks
                                              if b.kind == "categorical_distribution"
                                              and b.provenance.sheet == "HojaA"),
         "provenance": {"sheet": "HojaA"}},
        {"type": "chart", "block_ref": next(b.id for b in blocks
                                              if b.kind == "categorical_distribution"
                                              and b.provenance.sheet == "HojaB"),
         "provenance": {"sheet": "HojaB"}},
    ]
    out = _add_hallazgos_per_sheet(rendered, blocks, wb, blocks_by_id)
    sheets_with_hallazgos = {s.get("provenance", {}).get("sheet")
                              for s in out
                              if s.get("type") == "text_bullets"}
    assert "HojaA" in sheets_with_hallazgos
    assert "HojaB" in sheets_with_hallazgos


def test_auto_complete_skips_sheet_with_existing_hallazgos():
    wb = _make_wb_with_money_and_categories("HojaA")
    blocks = build_inventory(wb)
    blocks_by_id = {b.id: b for b in blocks}

    rendered = [
        {"type": "title", "data": {"title": "T"}},
        {"type": "chart", "block_ref": next(b.id for b in blocks
                                              if b.kind == "categorical_distribution"),
         "provenance": {"sheet": "HojaA"}},
        {"type": "text_bullets",
         "block_ref": next(b.id for b in blocks if b.kind == "table"),
         "supports_block": next(b.id for b in blocks if b.kind == "table"),
         "bullets": ["Existing bullet"],
         "provenance": {"sheet": "HojaA"}},
    ]
    out = _add_hallazgos_per_sheet(rendered, blocks, wb, blocks_by_id)
    hallazgos_for_a = [s for s in out
                        if s.get("type") == "text_bullets"
                        and (s.get("provenance") or {}).get("sheet") == "HojaA"]
    assert len(hallazgos_for_a) == 1   # not duplicated


def test_auto_complete_caps_added_hallazgos():
    """Si hay 10 hojas cubiertas, no añade 10 hallazgos — cap al razonable."""
    sheets = []
    for i in range(10):
        sheets.append(SheetData(
            name=f"Hoja{i}", shape=(20, 2), fill_ratio=0.9,
            columns=[
                ColumnData(name="C", dtype="categorical",
                            n_unique=3, fill_ratio=1.0,
                            samples=["a", "b"],
                            top_values=[("a", 10), ("b", 5)]),
            ],
            first_rows=[],
        ))
    wb = WorkbookData(filename="t.xlsx", sheets=sheets)
    blocks = build_inventory(wb)
    blocks_by_id = {b.id: b for b in blocks}

    rendered = [{"type": "title", "data": {"title": "T"}}]
    for i in range(10):
        bid = next((b.id for b in blocks
                    if b.kind == "categorical_distribution"
                    and b.provenance.sheet == f"Hoja{i}"), None)
        if bid:
            rendered.append({"type": "chart", "block_ref": bid,
                              "provenance": {"sheet": f"Hoja{i}"}})

    out = _add_hallazgos_per_sheet(rendered, blocks, wb, blocks_by_id)
    added = [s for s in out if s.get("type") == "text_bullets"]
    assert len(added) <= 3   # cap configured in MAX_AUTO_HALLAZGOS
```

- [ ] **Step 2: Run tests — verify FAIL (functions don't exist yet).**

Run: `python -m pytest __tests__/python/test_bullet_density.py -v --tb=short`
Expected: ImportError on `_topup_bullets` and `_add_hallazgos_per_sheet`.

- [ ] **Step 3: Implement `_topup_bullets` and `_add_hallazgos_per_sheet` in extractor.**

In `socya_pipeline/extractor.py`:

(a) Add at top, after existing constants:

```python
MIN_BULLETS = 5
MAX_BULLETS = 6
MAX_AUTO_HALLAZGOS = 3  # cap auto-añadidos para no inflar el deck
```

(b) Add helpers (before `_extract_kpi_row`):

```python
def _topup_bullets(existing: list, block, wb: WorkbookData) -> list:
    """Si `existing` tiene <MIN_BULLETS, agrega bullets determinísticos
    desde _build_insight_bullets sin duplicar. Cap a MAX_BULLETS."""
    if len(existing) >= MIN_BULLETS:
        return existing[:MAX_BULLETS]
    from socya_pipeline.planner import _build_insight_bullets
    candidates = _build_insight_bullets(wb, block, max_candidates=8)

    out = list(existing)
    seen_keys = {_bullet_dedup_key(b) for b in out}
    seen_angles: set = set()
    for angle, text in candidates:
        if len(out) >= MAX_BULLETS:
            break
        key = _bullet_dedup_key(text)
        if key in seen_keys:
            continue
        if angle in seen_angles:
            continue
        out.append(text)
        seen_keys.add(key)
        seen_angles.add(angle)
    return out


def _bullet_dedup_key(bullet: str) -> str:
    """Clave de dedup: extrae primer número + primera palabra significativa."""
    import re
    s = (bullet or "").lower()
    nums = re.findall(r"\d+(?:[.,]\d+)?", s)
    words = [w for w in re.findall(r"[a-záéíóúñ]{4,}", s)]
    n_part = nums[0] if nums else ""
    w_part = words[0] if words else ""
    return f"{n_part}|{w_part}"


def _add_hallazgos_per_sheet(rendered: list, inventory, wb: WorkbookData,
                              blocks_by_id: dict) -> list:
    """Para cada hoja cubierta por chart/table en `rendered` que NO tiene
    text_bullets, inserta una slide de hallazgos justo después.

    Cap MAX_AUTO_HALLAZGOS para no inflar el deck.
    """
    from socya_pipeline.planner import _build_insight_bullets

    sheets_with_text_bullets: set = {
        (s.get("provenance") or {}).get("sheet")
        for s in rendered if s.get("type") == "text_bullets"
    }
    sheets_with_text_bullets.discard(None)

    # Hojas cubiertas por chart/table en el orden que aparecen
    seen_sheets: list = []
    for s in rendered:
        if s.get("type") in ("chart", "table"):
            sh = (s.get("provenance") or {}).get("sheet")
            if sh and sh not in seen_sheets:
                seen_sheets.append(sh)

    targets = [sh for sh in seen_sheets
               if sh and sh not in sheets_with_text_bullets][:MAX_AUTO_HALLAZGOS]
    if not targets:
        return rendered

    # Construir hallazgos por sheet
    by_sheet: dict = {}
    for sheet_name in targets:
        block = next((b for b in inventory
                      if b.kind == "table" and b.provenance.sheet == sheet_name),
                      None)
        if block is None:
            # Try any block on this sheet
            block = next((b for b in inventory
                          if b.provenance.sheet == sheet_name), None)
        if block is None:
            continue
        candidates = _build_insight_bullets(wb, block, max_candidates=8)
        if not candidates:
            continue
        bullets: list = []
        seen_angles: set = set()
        for angle, text in candidates:
            if len(bullets) >= MAX_BULLETS:
                break
            if angle in seen_angles:
                continue
            bullets.append(text)
            seen_angles.add(angle)
        if len(bullets) >= 3:   # min de utilidad
            by_sheet[sheet_name] = (block, bullets)

    if not by_sheet:
        return rendered

    # Sandwich: insertar después del último chart/table de esa hoja
    out: list = []
    inserted_for: set = set()
    for s in rendered:
        out.append(s)
        sh = (s.get("provenance") or {}).get("sheet")
        if (sh in by_sheet
                and sh not in inserted_for
                and s.get("type") in ("chart", "table")):
            block, bullets = by_sheet[sh]
            out.append({
                "type": "text_bullets",
                "title": f"Hallazgos: {sh}",
                "supports_block": block.id,
                "data": {"bullets": bullets},
                "provenance": {"sheet": sh,
                                "columns": list(block.provenance.columns)},
                "_injected_by": "auto_hallazgos",
            })
            inserted_for.add(sh)
    return out
```

(c) Modify `extract_for_render` text_bullets branch (around line 130):

Find:
```python
        elif stype == "text_bullets":
            bullets = slide.get("bullets") or []
            if bullets:
                rendered.append({**slide, "data": {"bullets": bullets}})
            else:
                dropped.append({"type": stype, "reason": "bullets_empty",
                                "block_ref": block.id})
```

Replace with:
```python
        elif stype == "text_bullets":
            bullets = slide.get("bullets") or []
            # Top-up determinístico cuando el AI/validator dejaron <5 bullets.
            # Garantiza densidad sin depender de variabilidad del modelo.
            bullets = _topup_bullets(bullets, block, wb)
            if bullets:
                rendered.append({**slide, "data": {"bullets": bullets}})
            else:
                dropped.append({"type": stype, "reason": "bullets_empty",
                                "block_ref": block.id})
```

- [ ] **Step 4: Wire `_add_hallazgos_per_sheet` into `auto_complete_slides`.**

In `socya_pipeline/extractor.py`, find the end of `auto_complete_slides` (the `return title_slides + other + extra` line around line 332). BEFORE that final line, but after all the existing extra-building, add:

Find:
```python
    # Merge: keep the title slide first, then alternate richness.
    title_slides = [s for s in rendered if s.get("type") == "title"]
    other = [s for s in rendered if s.get("type") != "title"]
    if xls_owned:
        try: xls.close()
        except Exception: pass
    return title_slides + other + extra
```

Replace with:
```python
    # Merge: keep the title slide first, then alternate richness.
    title_slides = [s for s in rendered if s.get("type") == "title"]
    other = [s for s in rendered if s.get("type") != "title"]
    merged = title_slides + other + extra
    # Sandwich hallazgos per cubierta: inserta text_bullets después del
    # chart/table de cada hoja importante que no tenga ya su slide de
    # hallazgos. Cap interno a MAX_AUTO_HALLAZGOS para no inflar.
    blocks_by_id = {b.id: b for b in inventory}
    merged = _add_hallazgos_per_sheet(merged, inventory, wb, blocks_by_id)
    if xls_owned:
        try: xls.close()
        except Exception: pass
    return merged
```

- [ ] **Step 5: Run all tests.**

Run: `python -m pytest __tests__/python/ -v --tb=short`
Expected: all tests pass (152 existing + 7 new = 159 total).

- [ ] **Step 6: Commit.**

```bash
git add socya_pipeline/extractor.py __tests__/python/test_bullet_density.py
git commit -m "feat(extractor): top-up bullets to ≥5 + auto-add hallazgos per sheet"
```

---

## Task 3: End-to-end verification with real Excel

**Files:** none (testing only)

- [ ] **Step 1: Run CLI scenario with comisiones.**

Run:
```bash
SOCYA_AI_PROFILE=fast python -X utf8 -m socya_pipeline plan \
  --input "tests/fixtures/comisiones.xlsx" \
  --request '{"prompt":"Hazme 9 slides con la informacion mas importante."}'
```

- [ ] **Step 2: Verify output JSON.**

For each slide of type `text_bullets`, count bullets. Expected: ≥5 each.

Total deck should have ≥2 hallazgos slides (1 from AI + ≥1 from auto_complete).

- [ ] **Step 3: Document findings in plan file.**

Append a "Findings" section to `docs/superpowers/plans/2026-05-14-bullet-density.md` with:
- Scenarios tested
- Bullet counts per text_bullets slide
- Sheets covered by auto-added hallazgos

- [ ] **Step 4: Final commit.**

```bash
git add docs/superpowers/plans/2026-05-14-bullet-density.md
git commit -m "test: end-to-end verification of bullet density + auto hallazgos"
```

---

## Self-Review

Coverage check:
- [x] Pieza 1 (top-up) → Task 2 step 3
- [x] Pieza 2 (auto-add hallazgos) → Task 2 step 4
- [x] Pieza 3 (insight generators Pareto/outlier) → Task 1 step 2
- [x] Tests → Task 2 step 1
- [x] E2E verification → Task 3

No placeholders. Type/function names consistent: `_build_insight_bullets`, `_topup_bullets`, `_add_hallazgos_per_sheet`, `MIN_BULLETS=5`, `MAX_BULLETS=6`, `MAX_AUTO_HALLAZGOS=3`. Used identically across all tasks.

Risks (from spec) mitigated by tests: dedup in `test_topup_dedup_against_existing`, cap in `test_auto_complete_caps_added_hallazgos`, mix in `test_topup_mix_angles_not_repeats_same_generator`.
