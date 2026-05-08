"""Pulls real cell data per validated slide. Strips ugly literals."""
import math
from pathlib import Path
import pandas as pd
from socya_pipeline.parser import WorkbookData

UGLY_LITERALS_LOWER = {"nan", "none", "null", "nat", "???", "—", "s/d", "n/a", "na"}


def extract_for_render(validated_slides, inventory, wb: WorkbookData,
                        file_path) -> list:
    blocks_by_id = {b.id: b for b in inventory}
    xls = pd.ExcelFile(Path(file_path))
    sheets_cache = {}

    rendered = []
    for slide in validated_slides:
        stype = slide.get("type")
        if stype == "title":
            rendered.append({**slide, "data": {
                "title": slide.get("title", ""),
                "subtitle": slide.get("subtitle", ""),
            }})
            continue

        primary_id = slide.get("block_ref") or (slide.get("block_refs") or [None])[0] \
                      or slide.get("supports_block")
        block = blocks_by_id.get(primary_id) if primary_id else None
        if block is None:
            continue

        sheet_name = block.provenance.sheet
        if sheet_name not in sheets_cache:
            sheets_cache[sheet_name] = xls.parse(sheet_name)
        df = sheets_cache[sheet_name]

        if stype == "kpi_row":
            kpis = []
            for ref in slide.get("block_refs", []):
                b = blocks_by_id.get(ref)
                if not b or b.kind != "kpi_candidate":
                    continue
                value = b.extra.get("value")
                if value is None:
                    continue
                kpis.append({"label": b.label, "value": _format_kpi_value(value)})
            if kpis:
                rendered.append({**slide, "data": {"kpis": kpis}})

        elif stype == "chart":
            chart_data = _build_chart_data(block, df, slide.get("chart_type", "bar"))
            if chart_data:
                rendered.append({**slide, "data": chart_data})

        elif stype == "table":
            cols = slide.get("columns_subset") or block.provenance.columns
            cols = [c for c in cols if c in df.columns]
            sub = df[cols].copy()
            sub = _clean_dataframe(sub)
            max_rows = int(slide.get("max_rows") or 12)
            if sub.empty or len(sub.columns) < 2:
                continue
            sub = sub.head(max_rows)
            rendered.append({**slide, "data": {
                "headers": list(sub.columns),
                "rows": sub.values.tolist(),
            }})

        elif stype == "text_bullets":
            bullets = slide.get("bullets") or []
            if bullets:
                rendered.append({**slide, "data": {"bullets": bullets}})

    return rendered


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
        col = block.provenance.columns[0]
        if col not in df.columns:
            return None
        vc = df[col].dropna().astype(str).value_counts().head(6)
        if len(vc) < 2:
            return None
        return {
            "chart_type": chart_type,
            "name": col,
            "labels": vc.index.tolist(),
            "values": vc.values.tolist(),
        }
    if block.kind == "time_series_candidate":
        x_col = block.extra.get("x")
        y_col = block.extra.get("y")
        if x_col not in df.columns or y_col not in df.columns:
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


def _clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    def clean_cell(v):
        if v is None:
            return ""
        if isinstance(v, float) and math.isnan(v):
            return ""
        s = str(v).strip()
        if s.lower() in UGLY_LITERALS_LOWER:
            return ""
        return s

    # pandas 3.0+ removed applymap; use map (introduced as rename in 2.1)
    cleaned = df.map(clean_cell)
    # Drop rows with <50% filled
    row_fill = cleaned.apply(lambda r: sum(1 for v in r if v != "") / max(1, len(r)),
                               axis=1)
    cleaned = cleaned[row_fill >= 0.5]
    # Drop columns with <30% filled
    col_fill = cleaned.apply(lambda c: sum(1 for v in c if v != "") / max(1, len(c)),
                               axis=0)
    cleaned = cleaned.loc[:, col_fill >= 0.3]
    return cleaned
