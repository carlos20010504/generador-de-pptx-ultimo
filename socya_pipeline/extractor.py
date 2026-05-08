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
from socya_pipeline.parser import WorkbookData, promote_real_headers

UGLY_LITERALS_LOWER = {"nan", "none", "null", "nat", "???", "—", "s/d", "n/a", "na"}
MAX_TABLE_COLS = 6


def extract_for_render(validated_slides, inventory, wb: WorkbookData,
                        file_path) -> Tuple[List[dict], List[dict]]:
    """Returns (rendered_slides, dropped_slides). Dropped entries are
    {type, reason, block_ref?, title?} for transparency in audit.json."""
    blocks_by_id = {b.id: b for b in inventory}
    xls = pd.ExcelFile(Path(file_path))
    sheets_cache = {}

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

        # Load source sheet (cached)
        sheet_name = block.provenance.sheet
        if sheet_name not in sheets_cache:
            raw = xls.parse(sheet_name)
            sheets_cache[sheet_name] = promote_real_headers(xls, sheet_name, raw)
        df = sheets_cache[sheet_name]

        if stype == "chart":
            chart_data = _build_chart_data(block, df, slide.get("chart_type", "bar"))
            if chart_data:
                rendered.append({**slide, "data": chart_data})
            else:
                dropped.append({"type": stype, "reason": "chart_data_empty",
                                "block_ref": block.id})

        elif stype == "table":
            table_data = _build_table_data(slide, block, df)
            if table_data:
                rendered.append({**slide, "data": table_data})
            else:
                dropped.append({"type": stype, "reason": "table_data_empty",
                                "block_ref": block.id})

        elif stype == "text_bullets":
            bullets = slide.get("bullets") or []
            if bullets:
                rendered.append({**slide, "data": {"bullets": bullets}})
            else:
                dropped.append({"type": stype, "reason": "bullets_empty",
                                "block_ref": block.id})

    return rendered, dropped


def auto_complete_slides(rendered: List[dict], inventory, wb: WorkbookData,
                          file_path, target_count: int = 7) -> List[dict]:
    """If `rendered` has fewer than target_count slides, add high-quality
    slides from inventory blocks the planner didn't use. Editorial-grade
    fallback so the user always gets a substantive deck."""
    if len(rendered) >= target_count:
        return rendered

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
    xls = pd.ExcelFile(Path(file_path))
    sheets_cache = {}

    def get_df(sheet_name):
        if sheet_name not in sheets_cache:
            raw = xls.parse(sheet_name)
            sheets_cache[sheet_name] = promote_real_headers(xls, sheet_name, raw)
        return sheets_cache[sheet_name]

    extra: List[dict] = []
    needed = target_count - len(rendered)

    # 1. Add a kpi_row of the best unused KPI candidates
    unused_kpis = [b for b in inventory
                   if b.kind == "kpi_candidate"
                   and b.id not in used_block_ids
                   and b.extra.get("value") is not None]
    # Prefer currency KPIs first
    unused_kpis.sort(key=lambda b: (
        0 if b.extra.get("agg") == "sum" else 1,
        -(b.extra.get("value") or 0)
    ))
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

    # 3. Add a detail table from the largest unused table block
    unused_tables = [b for b in inventory
                     if b.kind == "table"
                     and b.id not in used_block_ids
                     and "low_fill_ratio" not in b.quality_flags
                     and "too_few_rows" not in b.quality_flags]
    unused_tables.sort(key=lambda b: -(b.extra.get("shape", [0])[0]))
    for tbl in unused_tables:
        if len(rendered) + len(extra) >= target_count:
            break
        if tbl.provenance.sheet in used_table_sheets:
            continue  # avoid duplicate table from the same sheet
        df = get_df(tbl.provenance.sheet)
        table_data = _build_table_data({"max_rows": 10}, tbl, df)
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

    # Merge: keep the title slide first, then alternate richness.
    title_slides = [s for s in rendered if s.get("type") == "title"]
    other = [s for s in rendered if s.get("type") != "title"]
    return title_slides + other + extra


def _col_norm(s: str) -> str:
    """Normalize a column name for semantic-equality comparison."""
    return str(s or "").strip().lower().replace(" ", "").replace("_", "")


def _extract_kpi_row(slide: dict, blocks_by_id: dict) -> Optional[dict]:
    kpis = []
    for ref in slide.get("block_refs", []):
        b = blocks_by_id.get(ref)
        if not b or b.kind != "kpi_candidate":
            continue
        value = b.extra.get("value")
        if value is None:
            continue
        description = ""
        agg = b.extra.get("agg")
        if agg == "sum":
            description = "Acumulado total"
        elif agg == "mean" and b.extra.get("min") is not None:
            description = (f"Rango: {_format_kpi_value(b.extra.get('min'))} – "
                            f"{_format_kpi_value(b.extra.get('max'))}")
        kpis.append({"label": b.label,
                      "value": _format_kpi_value(value),
                      "description": description})
    if not kpis:
        return None
    return {"kpis": kpis}


def _build_table_data(slide: dict, block, df: pd.DataFrame) -> Optional[dict]:
    """Build {headers, rows} for a table slide. Auto-picks the 6 best columns
    when AI didn't specify columns_subset (or specified ones that don't match)."""
    requested = slide.get("columns_subset") or []
    cols = [c for c in requested if c in df.columns]
    if not cols:
        # Auto-pick: top 6 columns by fill ratio, prefer mix of categorical+numeric
        cols = _auto_pick_columns(block, df, max_cols=MAX_TABLE_COLS)
    cols = cols[:MAX_TABLE_COLS]
    if not cols:
        return None

    sub = df[cols].copy()
    cleaned = _clean_dataframe(sub)
    if cleaned.empty or len(cleaned.columns) < 1:
        # Fall back to lightly cleaned (no fill-ratio drops, just ugly literal stripping)
        cleaned = sub.copy().map(_simple_clean)
    if cleaned.empty or len(cleaned.columns) < 1:
        return None

    max_rows = int(slide.get("max_rows") or 12)
    cleaned = cleaned.head(max_rows)
    return {
        "headers": [_clean_header(h) or f"Col {i+1}" for i, h in enumerate(cleaned.columns)],
        "rows": cleaned.values.tolist(),
    }


def _auto_pick_columns(block, df: pd.DataFrame, max_cols: int) -> List[str]:
    """Pick the top columns from a block by fill ratio, biasing toward
    a useful mix of identifying-name + categorical + numeric columns."""
    candidate_cols = [c for c in block.provenance.columns if c in df.columns]
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
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(f):
        return "—"
    if abs(f) >= 1_000_000:
        return f"{f/1_000_000:.1f}M"
    if abs(f) >= 1_000:
        return f"{f/1_000:.1f}K"
    if f.is_integer():
        return str(int(f))
    return f"{f:.2f}"


def _build_chart_data(block, df, chart_type):
    if block.kind == "categorical_distribution":
        col = _resolve_column(block.provenance.columns[0], df)
        if not col:
            return None
        vc = df[col].dropna().astype(str).str.strip()
        # Drop ugly literals from string values
        vc = vc[~vc.str.lower().isin(UGLY_LITERALS_LOWER)]
        vc = vc.value_counts().head(8)
        if len(vc) < 2:
            return None
        return {
            "chart_type": chart_type,
            "name": col,
            "labels": vc.index.tolist(),
            "values": [int(v) for v in vc.values.tolist()],
        }
    if block.kind == "time_series_candidate":
        x_col = _resolve_column(block.extra.get("x"), df)
        y_col = _resolve_column(block.extra.get("y"), df)
        if not x_col or not y_col:
            return None
        sub = df[[x_col, y_col]].dropna()
        if len(sub) < 2:
            return None
        sub = sub.sort_values(x_col).head(20)
        return {
            "chart_type": "line",
            "name": y_col,
            "labels": [str(x) for x in sub[x_col]],
            "values": [float(v) for v in pd.to_numeric(sub[y_col], errors="coerce").fillna(0)],
        }
    return None


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
    labels = chart_data.get("labels", [])
    values = chart_data.get("values", [])
    if len(labels) < 2:
        return ""
    parts = [f"{labels[0]} concentra {values[0]} registros"]
    if len(labels) >= 2:
        parts.append(f"seguido por {labels[1]} con {values[1]}")
    if len(labels) >= 3:
        parts.append(f"y {labels[2]} con {values[2]}")
    total = sum(values)
    pct = (values[0] / total * 100) if total else 0
    return ", ".join(parts) + f". El líder representa el {pct:.0f}% del total observado."


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.map(_simple_clean)
    # Drop rows with <50% filled
    row_fill = cleaned.apply(lambda r: sum(1 for v in r if v != "") / max(1, len(r)),
                               axis=1)
    cleaned = cleaned[row_fill >= 0.5]
    # Drop columns with <30% filled
    col_fill = cleaned.apply(lambda c: sum(1 for v in c if v != "") / max(1, len(c)),
                               axis=0)
    cleaned = cleaned.loc[:, col_fill >= 0.3]
    return cleaned
