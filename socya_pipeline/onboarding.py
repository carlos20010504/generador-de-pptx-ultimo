"""Smart-onboarding helpers — turn parser+inventory output into a friendly
summary the UI can show BEFORE the AI plan runs.

Zero AI calls here — all detection is deterministic so this can be invoked
on every upload without burning quota. The returned shape is JSON-friendly
and stable so the frontend can render it without surprises.

Returned structure:
{
  "filename": str,
  "totals": {
      "sheets": int,
      "rows": int,
      "columns": int,
      "kpi_candidates": int,
      "categorical_distributions": int,
      "time_series": int,
  },
  "sheets": [
      {"name": str, "rows": int, "columns": int, "fill_ratio": float,
        "is_redundant": bool, "first_cols": [str, ...]}
  ],
  "kpis_preview": [
      {"label": str, "value": str, "kind": str, "sheet": str}, ...
  ],
  "warnings": [
      {"severity": "info" | "warn" | "error",
        "icon": "low_fill" | "duplicate_sheet" | "no_kpis" | "missing_dates" | ...,
        "title": str, "detail": str}
  ],
  "suggestions": [
      {"id": str, "label": str, "prompt": str, "why": str}, ...
  ],
  "deck_estimate": {
      "min_slides": int,
      "max_slides": int,
      "expected_sections": [str, ...]
  }
}
"""
from __future__ import annotations

from typing import List, Dict, Any

from socya_pipeline import insights
from socya_pipeline.parser import WorkbookData


# ─────────────────────────────────────────────────────────────────────
# Public entry-point
# ─────────────────────────────────────────────────────────────────────

def quick_summary(wb: WorkbookData, inventory: list) -> dict:
    """Build the onboarding payload. Pure function: given the same inputs,
    always produces the same output. Safe to call on every upload."""
    sheets_summary = _build_sheets_summary(wb)
    totals = _build_totals(wb, inventory)
    kpis_preview = _build_kpis_preview(inventory, max_n=4)
    pii_findings = _detect_pii_in_workbook(wb)
    warnings = _build_warnings(wb, inventory, pii_findings)
    suggestions = _build_suggestions(wb, inventory)
    deck_estimate = _estimate_deck_size(inventory)

    return {
        "filename": wb.filename,
        "totals": totals,
        "sheets": sheets_summary,
        "kpis_preview": kpis_preview,
        "warnings": warnings,
        "suggestions": suggestions,
        "deck_estimate": deck_estimate,
        # Surface raw PII detections so the UI can build a dedicated alert
        # block (icon, count, list of (sheet, column, kind)) — separate from
        # the generic warnings list to allow distinct UX treatment.
        "pii_findings": pii_findings,
    }


# ─────────────────────────────────────────────────────────────────────
# Composition helpers
# ─────────────────────────────────────────────────────────────────────

def _build_totals(wb: WorkbookData, inv: list) -> dict:
    return {
        "sheets": len(wb.sheets),
        "rows": sum(s.shape[0] for s in wb.sheets),
        "columns": sum(s.shape[1] for s in wb.sheets),
        "kpi_candidates": sum(1 for b in inv if b.kind == "kpi_candidate"),
        "categorical_distributions": sum(1 for b in inv
                                              if b.kind == "categorical_distribution"),
        "time_series": sum(1 for b in inv if b.kind == "time_series_candidate"),
    }


def _build_sheets_summary(wb: WorkbookData) -> List[dict]:
    try:
        redundant = set(insights.redundant_sheet_ids(wb.sheets))
    except Exception:
        redundant = set()
    out = []
    for s in wb.sheets:
        first_cols = [insights.humanize_header(c.name) or c.name
                       for c in s.columns[:6]]
        out.append({
            "name": s.name,
            "rows": s.shape[0],
            "columns": s.shape[1],
            "fill_ratio": round(s.fill_ratio, 2),
            "is_redundant": s.name in redundant,
            "first_cols": first_cols,
        })
    return out


def _build_kpis_preview(inv: list, max_n: int = 4) -> List[dict]:
    """Show the strongest KPIs the parser detected — gives the user immediate
    signal that the system 'gets' their data."""
    kpis = [b for b in inv if b.kind == "kpi_candidate"
              and b.extra.get("value") is not None]
    # Same priority as the auto-completer: derived ratios > non-subsumed sums
    def _score(b):
        is_derived = "derived" in b.quality_flags
        is_subsumed = "subsumed_by_total" in b.quality_flags
        is_sum = b.extra.get("agg") == "sum"
        return (0 if is_derived else 1,
                1 if is_subsumed else 0,
                0 if is_sum else 1,
                -(b.extra.get("value") or 0))
    kpis.sort(key=_score)
    out = []
    for b in kpis[:max_n]:
        v = b.extra.get("value")
        unit = b.extra.get("display_unit")
        if unit == "%":
            scaled = v * 100 if -1 <= float(v) <= 1 else v
            display = f"{float(scaled):.1f}%"
        else:
            display = insights.format_compact(
                v, prefix="$" if b.extra.get("agg") == "sum" else "")
        out.append({
            "label": b.label,
            "value": display,
            "kind": "derived" if "derived" in b.quality_flags else b.extra.get("agg", "value"),
            "sheet": b.provenance.sheet,
        })
    return out


_PII_LABELS = {
    "email":   "correos",
    "phone":   "teléfonos",
    "doc_id":  "documentos / cédulas",
    "iban":    "cuentas bancarias",
    "credit_card": "tarjetas de crédito",
}


def _detect_pii_in_workbook(wb: WorkbookData) -> List[dict]:
    """Scan each column's samples for PII (emails, doc IDs, phones, etc.).
    Returns a list of {sheet, column, kind, label} so the UI can show an
    explicit alert and the user can decide BEFORE generating whether to
    enmascarar. Conservative: relies on insights.is_pii_column threshold."""
    findings: List[dict] = []
    for sheet in wb.sheets:
        for col in sheet.columns:
            try:
                kind = insights.is_pii_column(col.samples or [], col_name=col.name)
            except Exception:
                kind = None
            if kind:
                findings.append({
                    "sheet": sheet.name,
                    "column": col.name,
                    "kind": kind,
                    "label": _PII_LABELS.get(kind, kind),
                })
    return findings


def _build_warnings(wb: WorkbookData, inv: list,
                     pii_findings: List[dict] = None) -> List[dict]:
    out = []
    pii_findings = pii_findings or []

    # PII alert FIRST — most actionable thing the user needs to see.
    if pii_findings:
        # Group by kind to give a tight summary
        by_kind: dict = {}
        for f in pii_findings:
            by_kind.setdefault(f["kind"], []).append(f)
        parts = []
        for kind, items in by_kind.items():
            label = _PII_LABELS.get(kind, kind)
            sample_cols = ", ".join(f"'{i['column']}'" for i in items[:3])
            extra = f" y {len(items) - 3} más" if len(items) > 3 else ""
            parts.append(f"{label} en {sample_cols}{extra}")
        out.append({
            "severity": "warn", "icon": "pii_detected",
            "title": f"Datos personales detectados ({len(pii_findings)} columna{'s' if len(pii_findings) != 1 else ''})",
            "detail": "Encontramos: " + " · ".join(parts) +
                       ". Esos campos se enmascaran automáticamente al renderizar tablas, "
                       "pero te recomendamos revisar si deben aparecer en la presentación.",
        })

    # No numeric KPIs means a chart-driven deck is unlikely to land
    n_kpis = sum(1 for b in inv if b.kind == "kpi_candidate")
    if n_kpis == 0:
        out.append({
            "severity": "warn", "icon": "no_kpis",
            "title": "No detectamos KPIs numéricos",
            "detail": "El deck saldrá enfocado en tablas y categorías. "
                       "Si esperabas totales o promedios, verifica que las "
                       "columnas numéricas tengan valores limpios.",
        })

    # Time series presence (or absence)
    n_ts = sum(1 for b in inv if b.kind == "time_series_candidate")
    if n_ts == 0:
        # Only warn if there are date columns but no good Y pair — otherwise silent
        has_dates = any(c.dtype == "date" for s in wb.sheets for c in s.columns)
        if not has_dates:
            out.append({
                "severity": "info", "icon": "missing_dates",
                "title": "Sin columnas de fecha",
                "detail": "El deck no incluirá gráficos de evolución. "
                           "Agrega una columna 'Fecha' al Excel si quieres "
                           "tendencias.",
            })

    # Low-fill sheets
    sparse = [s for s in wb.sheets
                if s.fill_ratio < 0.5 and s.shape[0] >= 3]
    if sparse:
        names = ", ".join(f"'{s.name}'" for s in sparse[:3])
        out.append({
            "severity": "warn", "icon": "low_fill",
            "title": f"{len(sparse)} hoja(s) con poca data",
            "detail": f"{names} tiene <50% de celdas llenas. Verás menos "
                       "slides desde estas hojas.",
        })

    # Redundant duplicates
    try:
        redundant = list(insights.redundant_sheet_ids(wb.sheets))
    except Exception:
        redundant = []
    if redundant:
        names = ", ".join(f"'{n}'" for n in redundant[:3])
        out.append({
            "severity": "info", "icon": "duplicate_sheet",
            "title": f"{len(redundant)} hoja(s) duplicada(s) detectada(s)",
            "detail": f"{names} parece(n) copia(s) de otras hojas. Las "
                       "ignoraremos para no repetir slides.",
        })

    # Single-dominant categories will pivot to barh — surface it
    dom = [b for b in inv if b.kind == "categorical_distribution"
             and "single_dominant_category" in b.quality_flags]
    if len(dom) >= 2:
        out.append({
            "severity": "info", "icon": "single_dominant",
            "title": "Distribuciones muy desbalanceadas",
            "detail": f"{len(dom)} categorías tienen una opción que domina "
                       ">85%. Mostraremos esas como barras horizontales.",
        })

    return out


def _build_suggestions(wb: WorkbookData, inv: list) -> List[dict]:
    """Templated prompt suggestions — no AI, just pattern-matching on what
    the inventory contains. The user can click one to pre-fill the prompt."""
    out = []

    has_currency = any(b.extra.get("agg") == "sum"
                        for b in inv if b.kind == "kpi_candidate")
    has_geo = any("ciudad" in (c.name or "").lower()
                    or "destino" in (c.name or "").lower()
                    or "region" in (c.name or "").lower()
                    for s in wb.sheets for c in s.columns)
    has_time = any(b.kind == "time_series_candidate" for b in inv)
    has_status = any("estado" in (c.name or "").lower()
                       or "status" in (c.name or "").lower()
                       for s in wb.sheets for c in s.columns)
    has_ratio = any("derived" in b.quality_flags for b in inv)

    if has_currency and has_geo:
        out.append({
            "id": "geo_breakdown",
            "label": "Desglose financiero por ciudad",
            "prompt": "Análisis ejecutivo enfocado en montos por ciudad/destino, "
                       "identifica concentración geográfica y oportunidades.",
            "why": "Tienes columnas de valor + ubicación geográfica.",
        })

    if has_time and has_currency:
        out.append({
            "id": "trend_analysis",
            "label": "Evolución temporal de indicadores",
            "prompt": "Análisis de tendencias mensuales/trimestrales: "
                       "evolución de montos, picos y valles, estacionalidad.",
            "why": "Tienes fechas + valores numéricos.",
        })

    if has_status:
        out.append({
            "id": "status_funnel",
            "label": "Estado del proceso (semáforo)",
            "prompt": "Análisis del funnel: distribución por estado, "
                       "identificar cuellos de botella y casos rechazados.",
            "why": "Tienes columnas de estado/status — ideal para análisis de proceso.",
        })

    if has_ratio:
        out.append({
            "id": "ratio_focus",
            "label": "Foco en ratios y eficiencia",
            "prompt": "Reporte ejecutivo con énfasis en ratios derivados "
                       "(% aprobación, ejecución, cumplimiento) y brechas.",
            "why": "Detectamos pares Solicitado/Aprobado o similares.",
        })

    # Always-available generic suggestions
    out.append({
        "id": "executive_summary",
        "label": "Resumen ejecutivo conciso",
        "prompt": "Resumen ejecutivo de 6-7 slides: KPIs principales, "
                   "distribuciones clave, hallazgos accionables.",
        "why": "Funciona bien con cualquier dataset.",
    })
    out.append({
        "id": "deep_dive",
        "label": "Análisis exhaustivo (más slides)",
        "prompt": "Análisis profundo con 9-11 slides: incluye todos los "
                   "cortes relevantes y narrativas detalladas.",
        "why": "Para audiencias que quieren detalle.",
    })

    return out[:6]  # keep it focused — too many options = paralysis


def _estimate_deck_size(inv: list) -> dict:
    """Rough estimate based on the inventory. Honest about uncertainty."""
    n_kpis = sum(1 for b in inv if b.kind == "kpi_candidate")
    n_cats = sum(1 for b in inv if b.kind == "categorical_distribution"
                   and "single_dominant_category" not in b.quality_flags)
    n_ts = sum(1 for b in inv if b.kind == "time_series_candidate")
    n_tbls = sum(1 for b in inv if b.kind == "table"
                    and "low_fill_ratio" not in b.quality_flags
                    and "redundant_sheet" not in b.quality_flags)

    sections = ["Portada"]
    if n_kpis >= 1:
        sections.append("Indicadores clave")
    if n_cats >= 1:
        sections.extend(["Distribución principal"]
                          + (["Distribución secundaria"] if n_cats >= 2 else []))
    if n_ts >= 1:
        sections.append("Evolución temporal")
    if n_tbls >= 1:
        sections.append("Detalle operativo")
    sections.append("Hallazgos")

    estimate = max(7, min(11, len(sections) + 1))
    return {
        "min_slides": 7,
        "max_slides": estimate,
        "expected_sections": sections,
    }
