"""Conditional AI inspector — invoked ONLY when the deterministic parser
detects confusion (no typed columns, mostly Unnamed headers, very low fill,
single-column with data, etc.).

Design goals:
- ZERO AI calls for normal Excels (~95% of uploads). The trigger is strict.
- Cached aggressively (file hash + version) — repeated uploads of the same
  problematic Excel only cost the AI once.
- Best-effort: any failure here MUST NOT break the main pipeline. If the
  inspector throws, raises, or returns garbage, the parser falls back to
  whatever it had figured out on its own.
- Tiny token budget: ≤30 sample rows × ≤8 cols per problematic sheet.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional, Tuple

from socya_pipeline import insights
from socya_pipeline.ai_cache import PlanCache, compute_cache_key
from socya_pipeline.ai_chain import AIChain, AIProfile
from socya_pipeline.parser import WorkbookData

INSPECTOR_VERSION = "i1"
MAX_SAMPLE_ROWS = 30
MAX_SAMPLE_COLS = 8


# ─────────────────────────────────────────────────────────────────────
# Trigger detection — the gate that keeps token cost near zero
# ─────────────────────────────────────────────────────────────────────

def needs_inspection(wb: WorkbookData) -> List[Tuple[str, List[str]]]:
    """Return [(sheet_name, [reasons...])] for sheets where the parser likely
    failed to extract structure. Empty list = no AI call needed.

    Reasons are deliberately strict: a normal Excel triggers ZERO of them.
    """
    out: List[Tuple[str, List[str]]] = []
    for sheet in wb.sheets:
        reasons: List[str] = []

        # All columns came out as text/categorical with NO numeric typing —
        # parser found nothing useful for KPIs / charts.
        if len(sheet.columns) >= 2:
            typed_useful = sum(1 for c in sheet.columns
                                 if c.dtype in ("currency", "percent", "year",
                                                  "score", "numeric", "date"))
            if typed_useful == 0:
                reasons.append("no_typed_columns")

        # >50% of column headers are still 'Unnamed: N' after our promote
        # logic — header detection failed.
        if sheet.columns:
            unnamed = sum(1 for c in sheet.columns
                            if str(c.name).startswith("Unnamed:"))
            if unnamed / len(sheet.columns) > 0.5:
                reasons.append("many_unnamed_columns")

        # Very low fill ratio with non-trivial size → probable sub-tables
        # or weird layout where the data is sparse on purpose.
        if sheet.shape[0] >= 5 and sheet.fill_ratio < 0.30:
            reasons.append("low_fill_ratio")

        # Single column with multiple rows → a layout the parser can't help
        # with (probably needs a transpose or sub-table interpretation).
        if len(sheet.columns) == 1 and sheet.shape[0] > 5:
            reasons.append("single_column_with_data")

        if reasons:
            out.append((sheet.name, reasons))
    return out


# ─────────────────────────────────────────────────────────────────────
# Prompt + payload
# ─────────────────────────────────────────────────────────────────────

INSPECTOR_PROMPT = """Eres un analista de datos que inspecciona hojas de Excel raras.

Mi parser determinístico no pudo extraer una estructura clara de las hojas listadas abajo. Tu trabajo: mirar las primeras filas de cada hoja problemática y devolver un JSON con pistas para reinterpretarla.

Hojas problemáticas:
{payload_json}

Devuelve EXCLUSIVAMENTE un JSON con esta forma exacta (sin markdown):

{{
  "sheets": {{
    "<sheet_name>": {{
      "header_row_index": 0,           // índice 0-based de la fila con los headers reales
      "dtype_overrides": {{            // opcional: mapeo columna→dtype para corregir
        "<column_name>": "currency"   // valores válidos: currency, percent, date, year, numeric, categorical, text, bool
      }},
      "comment": "breve explicación humana"
    }}
  }}
}}

Reglas:
- Si no estás seguro, devuelve `{{"sheets": {{}}}}` — vale más no dar pistas que dar pistas malas.
- NUNCA inventes nombres de columna que no estén en los datos.
- Solo incluye sheets que aparecen en la lista de problemáticas.
"""


def _build_payload(wb: WorkbookData, issues: List[Tuple[str, List[str]]],
                     file_path) -> dict:
    """Pack a compact JSON describing the problematic sheets — first
    MAX_SAMPLE_ROWS rows, first MAX_SAMPLE_COLS cols only."""
    import pandas as pd
    xls = pd.ExcelFile(Path(file_path))
    payload: dict = {"problematic_sheets": []}
    for sheet_name, reasons in issues:
        try:
            raw = xls.parse(sheet_name, header=None, nrows=MAX_SAMPLE_ROWS)
        except Exception:
            continue
        rows = raw.iloc[:, :MAX_SAMPLE_COLS].fillna("").astype(str).values.tolist()
        # Trim each cell to keep tokens low
        rows = [[v[:40] for v in row] for row in rows]
        sheet = next((s for s in wb.sheets if s.name == sheet_name), None)
        payload["problematic_sheets"].append({
            "name": sheet_name,
            "issues_detected": reasons,
            "current_columns_detected_by_parser": (
                [c.name for c in sheet.columns] if sheet else []
            ),
            "first_rows_raw": rows,
        })
    return payload


# ─────────────────────────────────────────────────────────────────────
# Main entry — orchestrates trigger + cache + AI call + apply
# ─────────────────────────────────────────────────────────────────────

def inspect_workbook(wb: WorkbookData, file_path, api_key: str,
                       cache_dir: Optional[Path] = None) -> dict:
    """Returns a hints dict like:
        {"sheets": {"Hoja1": {"header_row_index": 4,
                                 "dtype_overrides": {"Cant": "currency"},
                                 "comment": "..."}}}
    Returns {} when no inspection is needed, when API fails, or when the
    response is malformed. NEVER raises."""
    if not api_key:
        return {}
    issues = needs_inspection(wb)
    if not issues:
        return {}

    # Cache: same file content + version → same hints, no extra AI call
    try:
        file_bytes = Path(file_path).read_bytes()
        ckey = compute_cache_key(file_bytes, "inspector", "", "",
                                   INSPECTOR_VERSION)
        cache = PlanCache(cache_dir=cache_dir)
        cached = cache.get(f"insp_{ckey}")
        if cached:
            return cached
    except OSError:
        cache = None
        ckey = None

    # Build the AI prompt (tiny budget)
    try:
        payload = _build_payload(wb, issues, file_path)
        payload_json = json.dumps(payload, ensure_ascii=False, default=str)
        prompt = INSPECTOR_PROMPT.format(payload_json=payload_json)
        chain = AIChain(api_key=api_key, profile=AIProfile.FAST)
        result = chain.call(prompt)
        hints = insights.parse_loose_json(result.content) or {}
        if cache and ckey and hints:
            try:
                cache.set(f"insp_{ckey}", hints)
            except Exception:
                pass
        return hints
    except Exception:
        # Best-effort: any failure → no hints, parser keeps its defaults
        return {}


def apply_inspector_hints(wb: WorkbookData, hints: dict) -> None:
    """Mutate `wb` in-place to apply the AI's hints. Never raises."""
    if not hints or not isinstance(hints, dict):
        return
    sheets_hints = hints.get("sheets") or {}
    if not isinstance(sheets_hints, dict):
        return
    valid_dtypes = {"currency", "percent", "date", "year", "score", "numeric",
                      "categorical", "text", "bool"}
    for sheet_name, sh in sheets_hints.items():
        if not isinstance(sh, dict):
            continue
        sheet = next((s for s in wb.sheets if s.name == sheet_name), None)
        if not sheet:
            continue
        # Apply dtype overrides where they're sensible
        overrides = sh.get("dtype_overrides") or {}
        if isinstance(overrides, dict):
            for col_name, dtype in overrides.items():
                if dtype not in valid_dtypes:
                    continue
                col = next((c for c in sheet.columns if c.name == col_name), None)
                if col:
                    col.dtype = dtype
        # NOTE: header_row_index would require re-reading the sheet from the
        # source file. For v1 we only apply dtype hints — header re-promotion
        # is a future extension once we validate the AI gives reliable indices.


# ─────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────

def _parse_response(text: str) -> dict:
    """Back-compat alias used by tests; new code should use insights.parse_loose_json."""
    return insights.parse_loose_json(text) or {}
