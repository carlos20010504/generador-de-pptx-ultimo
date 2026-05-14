"""Pulls real cell data per validated slide. Strips ugly literals.

Resilient: falls back to first columns if AI's column_subset doesn't match,
auto-picks 6 best columns for tables, returns (rendered_slides, dropped) so
the caller can surface drop reasons in the audit.

Also exposes `auto_complete_slides` — a non-AI completer that adds slides
from inventory blocks the planner didn't use, when the deck would otherwise
be too short.
"""
import math
from pathlib import Path
from typing import List, Tuple, Optional
import pandas as pd
from socya_pipeline.parser import WorkbookData, promote_real_headers, strip_total_rows
from socya_pipeline import insights

UGLY_LITERALS_LOWER = insights.UGLY_LITERALS  # back-compat alias
MAX_TABLE_COLS = 6
MIN_BULLETS = 5
MAX_BULLETS = 6
MAX_AUTO_HALLAZGOS = 3  # cap auto-añadidos para no inflar el deck


def _build_dtype_map(wb: WorkbookData) -> dict:
    """(sheet_name, col_name) → dtype. Built once per pipeline run."""
    return {(s.name, c.name): c.dtype for s in wb.sheets for c in s.columns}


def extract_for_render(validated_slides, inventory, wb: WorkbookData,
                        file_path,
                        xls: Optional[pd.ExcelFile] = None,
                        sheets_cache: Optional[dict] = None,
                        dtype_map: Optional[dict] = None,
                        ) -> Tuple[List[dict], List[dict]]:
    """Returns (rendered_slides, dropped_slides). Dropped entries are
    {type, reason, block_ref?, title?} for transparency in audit.json.

    `xls`, `sheets_cache`, and `dtype_map` can be supplied by the caller to
    avoid re-opening the Excel and re-deriving the dtype map between
    `extract_for_render` and `auto_complete_slides` (the two are typically
    invoked back-to-back from cli.py). When omitted they are built locally."""
    blocks_by_id = {b.id: b for b in inventory}
    # Si el caller no nos pasó un ExcelFile, lo abrimos nosotros — y SOMOS
    # responsables de cerrarlo al final. En Windows un ExcelFile sin cerrar
    # mantiene un lock sobre el archivo que impide re-uploads.
    xls_owned = xls is None
    if xls is None:
        xls = pd.ExcelFile(Path(file_path))
    if sheets_cache is None:
        sheets_cache = {}
    if dtype_map is None:
        dtype_map = _build_dtype_map(wb)

    rendered: List[dict] = []
    dropped: List[dict] = []

    for slide in validated_slides:
        stype = slide.get("type")

        # ── title ─────────────────────────────────────────
        if stype == "title":
            existing = slide.get("data") or {}
            rendered.append({**slide, "data": {
                "title": existing.get("title") or slide.get("title", ""),
                "subtitle": existing.get("subtitle") or slide.get("subtitle", ""),
            }})
            continue

        # Resolve primary block (for chart/table/text_bullets)
        primary_id = (slide.get("block_ref")
                      or (slide.get("block_refs") or [None])[0]
                      or slide.get("supports_block"))
        block = blocks_by_id.get(primary_id) if primary_id else None
        if block is None and stype != "kpi_row":
            dropped.append({"type": stype, "reason": "block_not_found",
                            "block_ref": primary_id})
            continue

        if stype == "kpi_row":
            data = _extract_kpi_row(slide, blocks_by_id)
            if data:
                rendered.append({**slide, "data": data})
            else:
                dropped.append({"type": stype, "reason": "no_valid_kpis"})
            continue

        # Load source sheet (cached). Apply the same cleaning the parser used
        # at inventory time so totals/headers stay consistent across stages.
        sheet_name = block.provenance.sheet
        if sheet_name not in sheets_cache:
            try:
                raw = xls.parse(sheet_name)
            except (ValueError, KeyError) as e:
                # Sheet referenced by the plan no longer exists in the file
                # (cache mismatch / renamed sheet). Drop the slide cleanly
                # rather than crashing the whole pipeline.
                dropped.append({"type": stype, "reason": "sheet_not_found",
                                "block_ref": block.id, "sheet": sheet_name,
                                "details": str(e)[:120]})
                continue
            cleaned = promote_real_headers(xls, sheet_name, raw)
            cleaned = strip_total_rows(cleaned)
            sheets_cache[sheet_name] = cleaned
        df = sheets_cache[sheet_name]

        if stype == "chart":
            chart_data = _build_chart_data(block, df, slide.get("chart_type", "bar"))
            if chart_data:
                # Anti-hallucination: if the AI's narrative contains numbers
                # that don't exist in the chart's real data, replace with an
                # auto-narrative built from chart_data itself.
                ai_narr = (slide.get("narrative") or "").strip()
                col_name = (block.provenance.columns[0]
                              if block.provenance.columns else "")
                honest_narr = _validate_or_replace_narrative(
                    ai_narr, chart_data, col_name)
                rendered.append({**slide, "narrative": honest_narr,
                                  "data": chart_data})
            else:
                dropped.append({"type": stype, "reason": "chart_data_empty",
                                "block_ref": block.id})

        elif stype == "table":
            table_data = _build_table_data(slide, block, df,
                                             dtype_map=dtype_map)
            if table_data:
                rendered.append({**slide, "data": table_data})
            else:
                dropped.append({"type": stype, "reason": "table_data_empty",
                                "block_ref": block.id})

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

    if xls_owned:
        try: xls.close()
        except Exception: pass
    return rendered, dropped


def auto_complete_slides(rendered: List[dict], inventory, wb: WorkbookData,
                          file_path, target_count: int = 7,
                          xls: Optional[pd.ExcelFile] = None,
                          sheets_cache: Optional[dict] = None,
                          dtype_map: Optional[dict] = None,
                          ) -> List[dict]:
    """Add high-quality slides from unused inventory blocks when `rendered`
    is short. `xls` / `sheets_cache` / `dtype_map` should be the same objects
    passed to `extract_for_render` to avoid re-opening the Excel.

    Nota: ya no early-returneamos cuando rendered == target — los pasos
    internos respetan el cap individualmente. Eliminamos el early return
    para que `_add_hallazgos_per_sheet` tenga oportunidad de correr y
    decida según su propio cap_external."""

    if dtype_map is None:
        dtype_map = _build_dtype_map(wb)

    used_block_ids = set()
    used_chart_columns = set()  # column names already charted (avoid duplicates by semantics)
    used_table_sheets = set()
    for s in rendered:
        for k in ("block_ref", "supports_block"):
            v = s.get(k)
            if v:
                used_block_ids.add(v)
        for v in s.get("block_refs", []) or []:
            used_block_ids.add(v)
        # Track semantic dedup
        prov = s.get("provenance") or {}
        if s.get("type") == "chart":
            for col in (prov.get("columns") or []):
                used_chart_columns.add(_col_norm(col))
        elif s.get("type") == "table":
            sheet = prov.get("sheet")
            if sheet:
                used_table_sheets.add(sheet)

    blocks_by_id = {b.id: b for b in inventory}
    xls_owned = xls is None
    if xls is None:
        xls = pd.ExcelFile(Path(file_path))
    if sheets_cache is None:
        sheets_cache = {}

    def get_df(sheet_name):
        if sheet_name not in sheets_cache:
            try:
                raw = xls.parse(sheet_name)
            except (ValueError, KeyError):
                # Cache None so we don't keep retrying the missing sheet
                sheets_cache[sheet_name] = None
                return None
            cleaned = promote_real_headers(xls, sheet_name, raw)
            cleaned = strip_total_rows(cleaned)
            sheets_cache[sheet_name] = cleaned
        return sheets_cache[sheet_name]

    extra: List[dict] = []
    needed = target_count - len(rendered)

    # 0. PRIORITY: hallazgos por hoja cubierta. Si rendered tiene un chart de
    # 'Riesgos' sin slide de hallazgos propia, sandwicheamos uno aquí ANTES
    # de meter chart filler. Hallazgos > charts genéricos cuando hay budget
    # apretado (target_count del user). El cap interno también limita.
    blocks_by_id_now = {b.id: b for b in inventory}
    rendered = _add_hallazgos_per_sheet(rendered, inventory, wb,
                                          blocks_by_id_now,
                                          target_count=target_count)
    if len(rendered) >= target_count:
        if xls_owned:
            try: xls.close()
            except Exception: pass
        return rendered

    # 1. Add a kpi_row of the best unused KPI candidates
    unused_kpis = [b for b in inventory
                   if b.kind == "kpi_candidate"
                   and b.id not in used_block_ids
                   and b.extra.get("value") is not None]
    # Prefer: derived ratios first (always interesting),
    # then non-subsumed currency totals (the headline numbers),
    # then everything else by magnitude.
    def _kpi_priority(b):
        is_derived = "derived" in b.quality_flags
        is_subsumed = "subsumed_by_total" in b.quality_flags
        is_sum = b.extra.get("agg") == "sum"
        return (
            0 if is_derived else 1,
            1 if is_subsumed else 0,
            0 if is_sum else 1,
            -(b.extra.get("value") or 0),
        )
    unused_kpis.sort(key=_kpi_priority)
    if unused_kpis and not any(s.get("type") == "kpi_row" for s in rendered):
        top = unused_kpis[:4]
        kpi_data = _extract_kpi_row(
            {"type": "kpi_row", "block_refs": [b.id for b in top]},
            blocks_by_id)
        if kpi_data:
            extra.append({
                "type": "kpi_row",
                "title": "Indicadores Adicionales",
                "data": kpi_data,
            })

    # 2. Add charts from unused categorical distributions (pick the best ones).
    # Skip free-text columns (high cardinality, no useful chart) and observations columns.
    BAD_CHART_NAMES = ("observ", "comentario", "descrip", "nota", "detalle",
                        "objeto", "agenda")
    def _good_for_chart(b):
        n = b.extra.get("n_unique") or 0
        if n < 2 or n > 12:
            return False
        col_name = (b.provenance.columns[0] or "").lower()
        for bad in BAD_CHART_NAMES:
            if bad in col_name:
                return False
        # Skip "Unnamed: N" / single-symbol columns — bad slide titles
        if (col_name.startswith("unnamed:")
                or sum(1 for c in col_name if c.isalnum()) <= 1):
            return False
        # Skip when every category has count=1 (no aggregation, just a list)
        top = b.extra.get("top_values") or []
        if top and all(int(c) <= 1 for _, c in top):
            return False
        return True

    unused_cats = [b for b in inventory
                   if b.kind == "categorical_distribution"
                   and b.id not in used_block_ids
                   and "single_dominant_category" not in b.quality_flags
                   and _good_for_chart(b)]
    # Sort by number of distinct categories (3-8 is ideal)
    def _cat_score(b):
        n = b.extra.get("n_unique") or 0
        if 3 <= n <= 8: return -10
        if 2 <= n <= 12: return -5
        return n
    unused_cats.sort(key=_cat_score)

    chart_types = ["bar", "pie", "bar"]
    chart_idx = 0
    for cat in unused_cats:
        if len(rendered) + len(extra) >= target_count:
            break
        col_name = cat.provenance.columns[0]
        if _col_norm(col_name) in used_chart_columns:
            continue  # avoid duplicate chart for the same semantic column
        df = get_df(cat.provenance.sheet)
        if df is None:
            continue  # source sheet missing — skip silently
        chart_data = _build_chart_data(cat, df, chart_types[chart_idx % len(chart_types)])
        if chart_data:
            chart_idx += 1
            used_chart_columns.add(_col_norm(col_name))
            extra.append({
                "type": "chart",
                "title": f"Distribución por {col_name.strip()}",
                "narrative": _auto_chart_narrative(chart_data, col_name),
                "data": chart_data,
                "provenance": {"sheet": cat.provenance.sheet,
                                "columns": [col_name]},
            })
        if chart_idx >= 3:  # cap auto charts at 3
            break

    # 3. Add a detail table from the largest unused table block. Skip
    # redundant sheets unless we have no alternative.
    candidates = [b for b in inventory
                   if b.kind == "table"
                   and b.id not in used_block_ids
                   and "low_fill_ratio" not in b.quality_flags
                   and "too_few_rows" not in b.quality_flags]
    non_redundant = [b for b in candidates
                      if "redundant_sheet" not in b.quality_flags]
    unused_tables = non_redundant if non_redundant else candidates
    unused_tables.sort(key=lambda b: -(b.extra.get("shape", [0])[0]))
    for tbl in unused_tables:
        if len(rendered) + len(extra) >= target_count:
            break
        if tbl.provenance.sheet in used_table_sheets:
            continue  # avoid duplicate table from the same sheet
        df = get_df(tbl.provenance.sheet)
        if df is None:
            continue  # source sheet missing
        table_data = _build_table_data({"max_rows": 10}, tbl, df,
                                         dtype_map=dtype_map)
        if table_data and len(table_data.get("headers", [])) >= 2:
            used_table_sheets.add(tbl.provenance.sheet)
            extra.append({
                "type": "table",
                "title": f"Detalle: {tbl.provenance.sheet.strip()}",
                "subtitle": f"Vista de {len(table_data['rows'])} filas representativas",
                "data": table_data,
                "provenance": {"sheet": tbl.provenance.sheet,
                                "columns": table_data.get("headers", [])},
            })

    # Merge: keep the title slide first, then alternate richness. Hallazgos
    # ya se sandwichearon en el paso 0 antes del padding — no re-correr.
    title_slides = [s for s in rendered if s.get("type") == "title"]
    other = [s for s in rendered if s.get("type") != "title"]
    if xls_owned:
        try: xls.close()
        except Exception: pass
    return title_slides + other + extra


def _col_norm(s: str) -> str:
    """Normalize a column name for semantic-equality comparison."""
    return str(s or "").strip().lower().replace(" ", "").replace("_", "")


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
                              blocks_by_id: dict,
                              target_count: Optional[int] = None) -> list:
    """Para cada hoja cubierta por chart/table en `rendered` que NO tiene
    text_bullets, inserta una slide de hallazgos justo después.

    Cap interno MAX_AUTO_HALLAZGOS para no inflar el deck. Si `target_count`
    está especificado, también respeta ese límite — para honrar el
    `requested_slide_count` del usuario.
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

    # Cap por MAX_AUTO_HALLAZGOS (interno) Y por target_count (externo del user).
    # Si el user pidió 9 slides y rendered ya tiene 7, podemos añadir hasta 2.
    cap_internal = MAX_AUTO_HALLAZGOS
    cap_external = (max(0, target_count - len(rendered))
                    if target_count is not None else cap_internal)
    cap = min(cap_internal, cap_external)
    targets = [sh for sh in seen_sheets
               if sh and sh not in sheets_with_text_bullets][:cap]
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


def _extract_kpi_row(slide: dict, blocks_by_id: dict) -> Optional[dict]:
    kpis = []
    used_blocks = []
    for ref in slide.get("block_refs", []):
        b = blocks_by_id.get(ref)
        if not b or b.kind != "kpi_candidate":
            continue
        value = b.extra.get("value")
        if value is None:
            continue
        description = ""
        agg = b.extra.get("agg")
        mean = b.extra.get("mean")
        unit = b.extra.get("display_unit")
        # Percent display takes precedence: a percent KPI must always render
        # with the % suffix, regardless of how it was computed (ratio derived
        # vs. mean of a percent column).
        if unit == "%":
            v = float(value)
            display_pct = v * 100 if -1 <= v <= 1 else v
            display_value = f"{display_pct:.1f}%"
            if agg == "ratio":
                description = "Indicador derivado (ratio)"
            elif b.extra.get("min") is not None and b.extra.get("max") is not None:
                lo = b.extra["min"]
                hi = b.extra["max"]
                lo_d = lo * 100 if -1 <= lo <= 1 else lo
                hi_d = hi * 100 if -1 <= hi <= 1 else hi
                # Skip the description if it's a trivial "0% – 100%" — that's
                # the theoretical full range of any percent column and gives
                # the reader nothing useful.
                trivial = (lo_d <= 0.5 and hi_d >= 99.5)
                if not trivial:
                    description = f"Rango: {lo_d:.1f}% – {hi_d:.1f}%"
        elif agg == "sum" and mean is not None:
            display_value = _format_kpi_value(value)
            description = f"Promedio por registro: {_format_kpi_value(mean)}"
        elif agg == "mean" and b.extra.get("min") is not None:
            display_value = _format_kpi_value(value)
            description = (f"Rango: {_format_kpi_value(b.extra.get('min'))} – "
                            f"{_format_kpi_value(b.extra.get('max'))}")
        else:
            display_value = _format_kpi_value(value)

        # Si el parser detectó una unidad explícita en el header (kg, USD,
        # Ton, hab, etc.) y el value formateado no la incluye ya, la
        # sufijamos. Así "1.5K" pasa a ser "1.5K kg" sin que el AI deba
        # adivinar la unidad.
        unit_hint = b.extra.get("unit_hint")
        if unit_hint and not _value_already_carries_unit(display_value, unit_hint):
            # Skip si la unidad es una currency code que ya se mostraría
            # con su prefijo $ (display_unit currency case ya lo cubre).
            if unit_hint not in ("USD", "EUR", "COP", "MXN", "CLP", "ARS", "BRL", "GBP", "JPY", "CNY"):
                display_value = f"{display_value} {unit_hint}"
        kpis.append({"label": b.label,
                      "value": display_value,
                      "description": description})
        used_blocks.append(b)
    if not kpis:
        return None
    # Disambiguate duplicate labels by appending the source sheet — two cards
    # called 'Total Solicitado' from sheets 'Master' and 'Hoja1' would
    # otherwise look identical to the reader.
    _disambiguate_kpi_labels(kpis, used_blocks)
    return {"kpis": kpis}


def _value_already_carries_unit(display_value: str, unit: str) -> bool:
    """True si el formateado ya menciona la unidad (case-insensitive).
    Evita 'kg kg' cuando el AI puso la unidad en el label original."""
    if not display_value or not unit:
        return False
    return unit.lower() in str(display_value).lower()


def _disambiguate_kpi_labels(kpis: List[dict], blocks: List) -> None:
    """In-place: when two KPIs have the same label, suffix the source sheet
    name in parentheses. No-op when all labels are already unique."""
    label_counts: dict = {}
    for k in kpis:
        label_counts[k["label"]] = label_counts.get(k["label"], 0) + 1
    if all(v == 1 for v in label_counts.values()):
        return
    for k, b in zip(kpis, blocks):
        if label_counts[k["label"]] > 1:
            sheet = (b.provenance.sheet or "").strip()
            if sheet and sheet not in k["label"]:
                k["label"] = f"{k['label']} ({sheet})"


def _build_table_data(slide: dict, block, df: pd.DataFrame,
                       dtype_map: Optional[dict] = None) -> Optional[dict]:
    """Build {headers, rows} for a table slide. Auto-picks the 6 best columns
    when AI didn't specify columns_subset; samples diverse rows (not just
    head); formats cells using the parser-inferred dtype when available
    ($1,234,567 for currency, 12.3% for percent, 2024 for year, etc.).
    Excludes ID-like columns unless explicitly requested."""
    dtype_map = dtype_map or {}
    sheet_name = block.provenance.sheet
    requested = slide.get("columns_subset") or []
    # Resolve requested columns with case-insensitive / whitespace tolerance —
    # AI sometimes emits "Ciudad_Destino" when the actual header is
    # "Ciudad_destino" or "ciudad_destino ". Without this, the slide gets
    # auto-picked columns instead of the ones the planner intended.
    cols = []
    for c in requested:
        resolved = _resolve_column(str(c), df) if c is not None else None
        if resolved is not None and resolved not in cols:
            cols.append(resolved)
    if not cols:
        cols = _auto_pick_columns(block, df, max_cols=MAX_TABLE_COLS)
    # Always filter ID-like columns (even if AI requested them — they're noise
    # for the reader). Exception: if filtering would leave us with nothing.
    filtered = [c for c in cols if not _looks_id_column(c, df[c])]
    cols = filtered if filtered else cols
    cols = cols[:MAX_TABLE_COLS]
    if not cols:
        return None

    sub = df[cols].copy()
    cleaned = _clean_dataframe(sub)
    if cleaned.empty or len(cleaned.columns) < 1:
        cleaned = sub.copy().map(_simple_clean)
    if cleaned.empty or len(cleaned.columns) < 1:
        return None

    max_rows = int(slide.get("max_rows") or 10)
    # Smart sort: si hay columna numérica sumable y el AI no especificó
    # un orden, ordenar por la primera numérica desc. La tabla "habla" del
    # ranking inmediatamente — el lector no tiene que buscar el máximo.
    cleaned = _smart_sort_table(cleaned, df, dtype_map, sheet_name)
    sampled = _diverse_sample(cleaned, df, max_rows)

    # Detect PII columns once per table — cheaper than per-cell checks
    pii_kinds = _detect_pii_columns(df, list(sampled.columns))

    # Per-column formatting based on source dtype + PII masking
    formatted_rows = []
    numeric_col_indices: list = []   # para totales row
    for col_idx, col_name in enumerate(sampled.columns):
        col_dtype = dtype_map.get((sheet_name, col_name))
        if col_dtype in ("currency", "numeric") and col_name not in pii_kinds:
            numeric_col_indices.append(col_idx)

    for _, row in sampled.iterrows():
        out_row = []
        for col_name in sampled.columns:
            raw_val = row[col_name]
            source_col = df[col_name] if col_name in df.columns else None
            kind = pii_kinds.get(col_name)
            if kind:
                out_row.append(insights.mask_pii(raw_val, kind) if raw_val else "")
            else:
                col_dtype = dtype_map.get((sheet_name, col_name))
                out_row.append(_format_table_cell(raw_val, source_col,
                                                   col_name, col_dtype))
        formatted_rows.append(out_row)

    # Total row: cuando hay >=1 columna numérica sumable y la tabla tiene
    # >=3 filas (con menos no aporta). Sólo currency/numeric, no percent
    # (% no se suma) ni year (años no se suman). Marcamos la fila como
    # is_total para que el renderer pueda darle estilo bold/separator.
    totals_row = None
    if len(formatted_rows) >= 3 and numeric_col_indices:
        totals_row = _build_totals_row(cleaned, sampled, numeric_col_indices,
                                          dtype_map, sheet_name)

    out = {
        "headers": [_humanize_header(h) or f"Col {i+1}" for i, h in enumerate(sampled.columns)],
        "rows": formatted_rows,
    }
    if totals_row:
        out["totals_row"] = totals_row
    # Pasamos qué columnas son numéricas para que el renderer pueda
    # aplicar heatmap-style cell coloring sin tener que re-detectar.
    if numeric_col_indices:
        out["numeric_col_indices"] = numeric_col_indices
        # Min/max por columna numérica para el gradient
        col_ranges = {}
        for ci in numeric_col_indices:
            col_name = sampled.columns[ci]
            try:
                series = pd.to_numeric(cleaned[col_name], errors="coerce").dropna()
                if not series.empty:
                    col_ranges[ci] = [float(series.min()), float(series.max())]
            except Exception:
                continue
        if col_ranges:
            out["numeric_col_ranges"] = col_ranges
    return out


def _smart_sort_table(cleaned, df, dtype_map: dict, sheet_name: str):
    """Si hay columna numérica/currency sumable, ordena descendente por la
    primera (mayor magnitud típica). El lector ve el ranking de un vistazo.
    Mantiene orden Excel cuando NO hay numérica clara."""
    try:
        for col_name in cleaned.columns:
            col_dtype = dtype_map.get((sheet_name, col_name))
            if col_dtype in ("currency", "numeric"):
                series = pd.to_numeric(cleaned[col_name], errors="coerce")
                # Skip si está casi todo vacío
                if series.notna().sum() < max(3, len(series) * 0.3):
                    continue
                return cleaned.assign(__sort_key=series).sort_values(
                    "__sort_key", ascending=False, na_position="last",
                ).drop(columns="__sort_key").reset_index(drop=True)
    except Exception:
        pass
    return cleaned


def _build_totals_row(cleaned, sampled, numeric_col_indices: list,
                       dtype_map: dict, sheet_name: str):
    """Devuelve una fila TOTAL alineada con las columnas de `sampled`.
    Suma sobre TODA la tabla cleaned (no sólo las filas sampled), porque
    la fila TOTAL debe representar el universo, no la muestra."""
    try:
        cells = []
        for col_idx, col_name in enumerate(sampled.columns):
            if col_idx == 0:
                cells.append("TOTAL")
                continue
            if col_idx in numeric_col_indices:
                series = pd.to_numeric(cleaned[col_name], errors="coerce")
                series = series.replace([float("inf"), float("-inf")], pd.NA).dropna()
                if series.empty:
                    cells.append("")
                    continue
                total = float(series.sum())
                col_dtype = dtype_map.get((sheet_name, col_name))
                # Reusar el formateador de celdas — paso valor sintético
                source_col = cleaned[col_name] if col_name in cleaned.columns else None
                cells.append(_format_table_cell(total, source_col,
                                                  col_name, col_dtype))
            else:
                cells.append("")
        return cells
    except Exception:
        return None


def _humanize_header(value) -> str:
    """Display-form column name (delegates to insights.humanize_header)."""
    return insights.humanize_header(value)


def _detect_pii_columns(df: pd.DataFrame, cols: list) -> dict:
    """Return {col_name: pii_kind} for columns whose sampled values are
    dominantly PII. Conservative: doc_id requires the column name to also
    hint at identification (cédula/documento/dni) — never enmascarar a
    column whose name screams 'valor' or 'monto' even if digits look ID-like."""
    out = {}
    for c in cols:
        if c not in df.columns:
            continue
        try:
            samples = df[c].dropna().head(30).tolist()
        except Exception:
            continue
        kind = insights.is_pii_column(samples, col_name=str(c))
        if kind:
            out[c] = kind
    return out


def _looks_id_column(col_name: str, series: pd.Series) -> bool:
    """An 'ID column' is one whose name matches an identifier pattern AND whose
    cardinality ≈ row count (each row a unique value). Currency is never ID."""
    name = str(col_name).strip().lower()
    if pd.api.types.is_numeric_dtype(series):
        # currency-like names should not be considered IDs
        for token in ("total", "valor", "monto", "precio", "costo", "ingreso",
                       "salario", "importe"):
            if token in name:
                return False
    for token in ("id ", " id", "id_", "_id", "código", "codigo", "número",
                   "numero", "n°", "no.", "folio", "consecutivo", "uuid", "key"):
        if token in name or name == token.strip():
            try:
                ratio = float(series.nunique(dropna=True)) / max(1, len(series.dropna()))
                if ratio >= 0.85:
                    return True
            except Exception:
                return True
    return False


def _diverse_sample(cleaned: pd.DataFrame, source_df: pd.DataFrame,
                     max_rows: int) -> pd.DataFrame:
    """Pick `max_rows` rows that are diverse on the most-populated string
    column (so a table about Comisiones doesn't show 9 rows of the same
    Solicitante). Falls back to head() when no string column has good cardinality.
    """
    if len(cleaned) <= max_rows:
        return cleaned

    # Pick a "diversity column" — the first string column with reasonable cardinality
    div_col = None
    for c in cleaned.columns:
        if c not in source_df.columns:
            continue
        s = source_df[c]
        if pd.api.types.is_numeric_dtype(s):
            continue
        n_unique = s.nunique(dropna=True)
        if 3 <= n_unique <= max(50, len(s) * 0.5):
            div_col = c
            break

    if div_col is None:
        return cleaned.head(max_rows)

    # 1 row per unique value of div_col, up to max_rows
    seen = set()
    picked_idx = []
    for idx, val in cleaned[div_col].items():
        key = str(val).strip().lower()
        if key in seen or not key:
            continue
        seen.add(key)
        picked_idx.append(idx)
        if len(picked_idx) >= max_rows:
            break
    if len(picked_idx) < max_rows:
        # Top-up with sequential rows we haven't picked
        for idx in cleaned.index:
            if idx not in picked_idx:
                picked_idx.append(idx)
                if len(picked_idx) >= max_rows:
                    break
    return cleaned.loc[picked_idx[:max_rows]]


def _format_table_cell(value, source_series, col_name: str,
                         col_dtype: Optional[str] = None) -> str:
    """Format a cell based on the source column's inferred dtype.
        currency → $1,234,567
        percent  → 12.3%
        year     → 2024 (no thousand separator)
        score    → 4.5
        numeric  → 1,234,567 (or 1.5K/1.5M for big magnitudes)
        date     → ISO
        text     → trimmed and de-shouted

    When `col_dtype` is provided (from parser), uses it as the authoritative
    signal. Falls back to name-based heuristics only when dtype is missing."""
    if value is None or value == "":
        return ""

    dt_hint = (col_dtype or "").lower()

    # Date columns get formatted before any other coercion. Pandas Timestamps
    # stringify as "2024-06-12 09:21:25.987000" by default which is unreadable
    # in a deck table — we strip to ISO date or short datetime.
    if dt_hint == "date":
        try:
            ts = pd.to_datetime(value, errors="raise")
            # If time portion is meaningless (midnight), show date only
            if ts.hour == 0 and ts.minute == 0 and ts.second == 0:
                return ts.strftime("%Y-%m-%d")
            return ts.strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass

    # Try numeric coercion first
    try:
        num = float(str(value).replace(",", "").replace("$", "")
                      .replace("%", "").strip())
        if math.isnan(num):
            return ""

        dt = dt_hint
        if dt == "currency":
            return _format_currency_compact(num)
        if dt == "percent":
            # Decimal form (e.g. 0.22, -0.22) → scale to %; otherwise assume
            # already on 0..100 scale.
            display = num * 100 if -1 <= num <= 1 else num
            return f"{display:.1f}%"
        if dt == "year":
            return str(int(num)) if num.is_integer() else f"{num:.0f}"
        if dt == "score":
            return f"{num:.1f}" if not num.is_integer() else str(int(num))

        # No explicit dtype: fallback to name-based heuristics (i18n vocabulary)
        if not dt:
            if insights.looks_money_by_name(col_name):
                return _format_currency_compact(num)

        if num.is_integer() and abs(num) < 100_000:
            return f"{int(num):,}"
        if abs(num) >= 1_000:
            return f"{num:,.0f}" if num.is_integer() else f"{num:,.1f}"
        return str(int(num)) if num.is_integer() else f"{num:.2f}"
    except (ValueError, AttributeError):
        pass

    # Non-numeric cell
    s = str(value).strip()
    if len(s) > 38:
        s = s[:36] + "…"
    if s.isupper() and len(s) > 6:
        s = s.title()
    return s


def _format_currency_compact(num: float) -> str:
    return insights.format_compact(num, prefix="$")


def _auto_pick_columns(block, df: pd.DataFrame, max_cols: int) -> List[str]:
    """Pick the top columns from a block by fill ratio, biasing toward
    a useful mix of identifying-name + categorical + numeric columns."""
    # Tolerate case/whitespace mismatches between block.provenance.columns
    # (recorded at inventory time) and the live df.columns (could have been
    # re-headered by promote_real_headers).
    candidate_cols = []
    for c in block.provenance.columns:
        resolved = _resolve_column(str(c), df) if c is not None else None
        if resolved is not None and resolved not in candidate_cols:
            candidate_cols.append(resolved)
    if not candidate_cols:
        candidate_cols = list(df.columns)

    # Score each column
    def score(col_name: str) -> float:
        s = df[col_name]
        fill = float(s.notna().sum()) / max(1, len(s))
        if fill < 0.2:
            return -1.0
        # Reward "name-ish" columns first (Solicitante, Nombre, Cliente, etc.)
        name_lower = str(col_name).lower()
        bonus = 0.0
        for token in ("nombre", "solicitante", "cliente", "ciudad", "estado",
                       "fecha", "mes ", "categoría", "categoria", "tipo",
                       "valor total", "total", "monto", "código", "id "):
            if token in name_lower:
                bonus += 0.3
                break
        # Penalize "Unnamed:" columns
        if name_lower.startswith("unnamed:"):
            bonus -= 0.5
        return fill + bonus

    scored = [(score(c), c) for c in candidate_cols]
    scored = [(s, c) for s, c in scored if s > 0]
    scored.sort(key=lambda x: -x[0])
    return [c for _, c in scored[:max_cols]]


def _clean_header(value) -> str:
    s = str(value or "").strip()
    if s.startswith("Unnamed:") or s.lower() in ("nan", "none", "null"):
        return ""
    return s


def _simple_clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    s = str(v).strip()
    if s.lower() in UGLY_LITERALS_LOWER:
        return ""
    return s


def _format_kpi_value(value):
    return insights.format_compact(value)


def _build_chart_data(block, df, chart_type):
    # Histogram: requiere columna numérica continua. Funciona bien con
    # block.kind == "kpi_candidate" (que apunta a una columna numeric/currency).
    if chart_type == "histogram" and block.provenance.columns:
        col = _resolve_column(block.provenance.columns[0], df)
        if not col:
            return None
        nums = pd.to_numeric(df[col], errors="coerce")
        nums = nums.replace([float("inf"), float("-inf")], pd.NA).dropna()
        if len(nums) >= 8:
            return _build_histogram_data(col, nums)
        # Insuficientes puntos para un histograma honesto. Devolvemos None
        # → la slide se dropea limpiamente como chart_data_empty. Mejor que
        # forzar a "bar" sobre un kpi_candidate (que no es categorical y
        # nunca rellena datos válidos).
        return None

    if block.kind == "categorical_distribution":
        col = _resolve_column(block.provenance.columns[0], df)
        if not col:
            return None
        vc = df[col].dropna().astype(str).str.strip()
        vc = vc[~vc.str.lower().isin(UGLY_LITERALS_LOWER)]
        vc = vc.value_counts().head(15)  # take a few extra so Top-N has room
        if len(vc) < 2:
            return None
        labels = vc.index.tolist()
        values = [int(v) for v in vc.values.tolist()]
        labels, values, is_chrono = _maybe_sort_chronologically(labels, values)
        # Don't roll a chronological series into "Otros" — the order matters.
        if not is_chrono:
            labels, values = _consolidate_long_tail(labels, values, max_categories=10)
        else:
            labels = labels[:12]
            values = values[:12]
        labels, values = _drop_non_finite_pairs(labels, values)
        if len(values) < 2:
            return None
        return {
            "chart_type": chart_type,
            "name": col,
            "labels": labels,
            "values": values,
            "chronological": is_chrono,
        }
    if block.kind == "time_series_candidate":
        x_col = _resolve_column(block.extra.get("x"), df)
        y_col = _resolve_column(block.extra.get("y"), df)
        if not x_col or not y_col:
            return None
        sub = df[[x_col, y_col]].dropna().copy()
        if len(sub) < 2:
            return None
        # Coerce to datetime + numeric. Aggregate by month so we get a clean
        # ~12-point trend instead of N raw rows with timestamps.
        sub[x_col] = pd.to_datetime(sub[x_col], errors="coerce")
        sub[y_col] = pd.to_numeric(sub[y_col], errors="coerce")
        # Drop ±inf along with NaN before aggregations — otherwise sums
        # inherit the inf and the chart blows up.
        sub[y_col] = sub[y_col].replace([float("inf"), float("-inf")], pd.NA)
        sub = sub.dropna()
        if len(sub) < 2:
            return None
        # Decide aggregation granularity by date span
        span_days = (sub[x_col].max() - sub[x_col].min()).days
        if span_days > 365:
            sub["__period"] = sub[x_col].dt.to_period("Q").astype(str)
        elif span_days > 60:
            sub["__period"] = sub[x_col].dt.to_period("M").astype(str)
        elif span_days > 14:
            sub["__period"] = sub[x_col].dt.to_period("W").astype(str)
        else:
            sub["__period"] = sub[x_col].dt.strftime("%Y-%m-%d")
        agg = sub.groupby("__period")[y_col].sum().reset_index()
        agg = agg.sort_values("__period").head(24)
        if len(agg) < 2:
            return None
        labels = [str(x) for x in agg["__period"]]
        values = [float(v) for v in agg[y_col]]
        labels, values = _drop_non_finite_pairs(labels, values)
        if len(values) < 2:
            return None
        return {
            "chart_type": "line",
            "name": y_col,
            "labels": labels,
            "values": values,
        }
    return None


def _build_histogram_data(col_name: str, nums) -> dict:
    """Bin a numeric series into ~10 bins for histogram rendering. Choses
    bin count by Freedman-Diaconis when feasible, else sqrt rule. Returns
    labels (bin range strings) + counts."""
    try:
        import numpy as _np
        arr = _np.array(nums.tolist(), dtype=float)
        if arr.size == 0:
            return None
        # Freedman-Diaconis bin width
        q75, q25 = _np.percentile(arr, [75, 25])
        iqr = float(q75 - q25)
        n = arr.size
        if iqr > 0:
            h = 2 * iqr * (n ** (-1 / 3))
            data_range = float(arr.max() - arr.min())
            bins = max(5, min(15, int(_np.ceil(data_range / max(h, 1e-9)))))
        else:
            bins = max(5, min(12, int(_np.ceil(_np.sqrt(n)))))
        counts, edges = _np.histogram(arr, bins=bins)
        labels = []
        for lo, hi in zip(edges[:-1], edges[1:]):
            labels.append(f"{insights.format_compact(lo)}–{insights.format_compact(hi)}")
        return {
            "chart_type": "histogram",
            "name": col_name,
            "labels": labels,
            "values": [int(c) for c in counts.tolist()],
            "_stats": {
                "mean": float(arr.mean()),
                "median": float(_np.median(arr)),
                "n": int(n),
            },
        }
    except Exception:
        return None


def _drop_non_finite_pairs(labels, values):
    """Remove (label, value) pairs where value isn't a finite number. Some
    sheets sneak NaN/Inf past upstream cleanups via aggregations; if any
    survives into matplotlib it crashes silently and the slide renders blank.
    """
    out_l, out_v = [], []
    for l, v in zip(labels, values):
        try:
            fv = float(v)
        except (TypeError, ValueError):
            continue
        if not math.isfinite(fv):
            continue
        out_l.append(l)
        out_v.append(fv)
    return out_l, out_v


def _consolidate_long_tail(labels, values, max_categories: int = 10):
    """Group the long tail of small categories into one 'Otros' slice when a
    chart has more than `max_categories` categories. Bar/donut/pie charts
    with >12 slices become unreadable; collapsing the long tail keeps the
    leaders legible without losing total magnitude.
    Assumes labels/values are already sorted high-to-low (vc.value_counts()).
    """
    if len(labels) <= max_categories:
        return labels, values
    head_l = list(labels[: max_categories - 1])
    head_v = list(values[: max_categories - 1])
    tail_v = sum(values[max_categories - 1 :])
    head_l.append("Otros")
    head_v.append(tail_v)
    return head_l, head_v


def _maybe_sort_chronologically(labels, values):
    """If labels look like month names (any supported language) or weekdays,
    sort chronologically. Returns (labels, values, is_chronological).
    Delegates language detection to insights.month_index/weekday_index."""
    if not labels:
        return labels, values, False
    months = [insights.month_index(l) for l in labels]
    if sum(1 for m in months if m is not None) / len(labels) >= 0.7:
        indexed = sorted(enumerate(months),
                          key=lambda t: t[1] if t[1] is not None else 999)
        return ([labels[i] for i, _ in indexed],
                [values[i] for i, _ in indexed],
                True)
    weekdays = [insights.weekday_index(l) for l in labels]
    if sum(1 for d in weekdays if d is not None) / len(labels) >= 0.7:
        indexed = sorted(enumerate(weekdays),
                          key=lambda t: t[1] if t[1] is not None else 999)
        return ([labels[i] for i, _ in indexed],
                [values[i] for i, _ in indexed],
                True)
    return labels, values, False


def _resolve_column(name: str, df: pd.DataFrame) -> Optional[str]:
    """Find `name` in df.columns with tolerance: exact match first, then
    stripped, then case-insensitive, then substring."""
    if name in df.columns:
        return name
    for c in df.columns:
        if str(c).strip() == str(name).strip():
            return c
    target = str(name).strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == target:
            return c
    for c in df.columns:
        if target in str(c).strip().lower():
            return c
    return None


def _auto_chart_narrative(chart_data: dict, col_name: str) -> str:
    """Build an honest narrative ENTIRELY from chart_data — every number
    cited here is guaranteed to be in the chart itself.

    Adds insight layers (degrade gracefully if `insights` isn't useful):
      - For chronological / line data: peak, valley, growth, *seasonality*.
      - For ranked data: leader, runners-up, *Pareto concentration* of top-3,
        and outlier callout when one value dominates the average.
    """
    labels = chart_data.get("labels", [])
    values = chart_data.get("values", [])
    if len(labels) < 2:
        return ""
    chart_type = chart_data.get("chart_type", "bar")
    is_chrono = bool(chart_data.get("chronological"))
    total = sum(float(v) for v in values) or 1

    if chart_type == "line" or is_chrono:
        first_v = values[0]
        last_v = values[-1]
        max_idx = max(range(len(values)), key=lambda i: values[i])
        min_idx = min(range(len(values)), key=lambda i: values[i])
        peak = (labels[max_idx], values[max_idx])
        valley = (labels[min_idx], values[min_idx])
        delta = last_v - first_v
        delta_pct = (delta / first_v * 100) if first_v else 0
        direction = ("creció" if delta > 0
                     else ("decreció" if delta < 0 else "se mantuvo"))
        body = (f"El máximo se registró en {peak[0]} con {_fmt_n(peak[1])}, "
                f"el mínimo en {valley[0]} con {_fmt_n(valley[1])}. "
                f"Entre {labels[0]} y {labels[-1]} el indicador {direction} "
                f"un {abs(delta_pct):.0f}%.")
        # Optional seasonality hint
        try:
            hint = insights.seasonality_hint(labels, values)
            if hint:
                body += " " + hint
        except Exception:
            pass
        # Trend acceleration / deceleration — only when there's a clear shift
        # in slope between the first and second halves of the series.
        try:
            accel = insights.trend_acceleration(values)
            if accel == "accelerating":
                body += " La tendencia se acelera en el último tramo."
            elif accel == "decelerating":
                body += " El crecimiento pierde fuerza en el último tramo."
        except Exception:
            pass
        # Volatility callout — useful for "esto va en zigzag, no tendencia clara"
        try:
            vol = insights.volatility_score(values)
            if vol is not None and vol >= 0.55 and len(values) >= 6:
                body += f" Serie volátil (CV ≈ {vol*100:.0f}%); promedios suavizan más que la línea."
        except Exception:
            pass
        return body

    # Ranking-style — first label IS the leader by sort
    lead_pct = values[0] / total * 100
    parts = [f"{labels[0]} lidera con {_fmt_n(values[0])} ({lead_pct:.0f}% del total)"]
    if len(labels) >= 3:
        parts.append(f"seguido por {labels[1]} ({_fmt_n(values[1])}) "
                      f"y {labels[2]} ({_fmt_n(values[2])})")
    elif len(labels) >= 2:
        parts.append(f"seguido por {labels[1]} con {_fmt_n(values[1])}")
    body = ", ".join(parts) + "."

    # Pareto-style insight: top-3 share of total
    try:
        share = insights.pareto_share(values, top_n=3)
        if share is not None and len(values) >= 5 and share >= 0.6:
            body += f" Top 3 concentra el {share*100:.0f}% del total."
    except Exception:
        pass

    if len(labels) > 4:
        body += (f" Las {len(labels) - 3} categorías restantes acumulan "
                  f"{_fmt_n(total - sum(values[:3]))}.")

    # IQR outliers — catches BOTH high and low outliers, anywhere in the
    # ranking (not just the leader), and is robust to skewed distributions.
    # We keep the leader-vs-median heuristic as a backup when IQR finds none.
    try:
        outliers = insights.iqr_outliers(values, k=1.5)
        # Filter to "meaningful" — skip outliers that are within the natural
        # leader gap (which the leader callout already covers).
        high = [(i, v) for i, v, side in outliers if side == "high"]
        low = [(i, v) for i, v, side in outliers if side == "low"]
        callouts = []
        if high:
            top = sorted(high, key=lambda t: -t[1])[:2]
            names = " y ".join(f"{labels[i]} ({_fmt_n(v)})" for i, v in top)
            callouts.append(f"sobresalen {names}")
        if low and len(values) >= 6:
            bot = sorted(low, key=lambda t: t[1])[:2]
            names = " y ".join(f"{labels[i]} ({_fmt_n(v)})" for i, v in bot)
            callouts.append(f"quedan rezagados {names}")
        if callouts:
            body += " Como casos atípicos, " + " y ".join(callouts) + "."
    except Exception:
        # Fallback to the simple ratio callout
        try:
            rest = sorted([float(v) for v in values[1:]], reverse=True)
            if rest:
                median_rest = rest[len(rest) // 2]
                if median_rest > 0 and values[0] / median_rest >= 5:
                    ratio = values[0] / median_rest
                    body += (f" {labels[0]} es {ratio:.0f}x la mediana del resto, "
                              f"un outlier marcado.")
        except Exception:
            pass

    return body


def _fmt_n(v):
    """Compact number formatter for narratives — matches what charts show."""
    return insights.format_compact(v)


_NUM_RE_NARR = __import__("re").compile(r"\d{1,3}(?:[\.,]\d{3})*(?:[\.,]\d+)?|\d+")


def _validate_or_replace_narrative(ai_narr: str, chart_data: dict,
                                     col_name: str) -> str:
    """If `ai_narr` mentions any integer that doesn't exist in chart_data
    (or a close fuzzy match), replace it with an auto-generated narrative
    built only from real chart values."""
    if not ai_narr:
        return _auto_chart_narrative(chart_data, col_name)

    chart_values = [float(v) for v in chart_data.get("values", []) if v is not None]
    chart_total = sum(chart_values)
    valid_nums = set(chart_values) | {chart_total}
    # Allow rounded versions: int, /1000, /1000000
    expanded = set(valid_nums)
    for v in valid_nums:
        expanded.add(round(v))
        expanded.add(round(v / 1_000, 1))
        expanded.add(round(v / 1_000_000, 1))
    # Allow %s of values vs total
    if chart_total > 0:
        for v in chart_values:
            expanded.add(round(v / chart_total * 100))

    found = _NUM_RE_NARR.findall(ai_narr)
    for token in found:
        try:
            n = float(token.replace(".", "").replace(",", "."))
        except ValueError:
            try:
                n = float(token.replace(",", ""))
            except ValueError:
                continue
        # Match if any chart value is within 2% relative tolerance
        ok = False
        for h in expanded:
            if h == 0:
                if n == 0:
                    ok = True; break
                continue
            if abs(n - h) <= max(1.0, abs(h) * 0.02):
                ok = True; break
        if not ok:
            return _auto_chart_narrative(chart_data, col_name)
    return ai_narr


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Strip ugly placeholders cell-by-cell, then drop sparse rows/cols.
    Vectorized — no Python `apply(axis=1)` per-row scan."""
    cleaned = df.map(_simple_clean)
    if cleaned.empty:
        return cleaned
    n_cols = cleaned.shape[1] or 1
    # Vectorized fill ratio per row: count non-empty cells with `.ne("")`
    row_fill = cleaned.ne("").sum(axis=1) / n_cols
    cleaned = cleaned[row_fill >= 0.5]
    if cleaned.empty:
        return cleaned
    n_rows = cleaned.shape[0] or 1
    col_fill = cleaned.ne("").sum(axis=0) / n_rows
    cleaned = cleaned.loc[:, col_fill >= 0.3]
    return cleaned
