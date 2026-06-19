"""Single-call IA planner. Takes inventory + prompt, returns full slide plan."""
import json
from pathlib import Path
from typing import Optional
from socya_pipeline.ai_cache import PlanCache, compute_cache_key
from socya_pipeline.ai_chain import AIChain, AIProfile
from socya_pipeline.errors import PipelineError, ErrorCode
from socya_pipeline.inventory import Block
from socya_pipeline.parser import WorkbookData
from socya_pipeline import insights

PLANNER_VERSION = "p7"  # bump when prompt template changes — invalidates cache

MAX_PAYLOAD_CHARS = 24_000  # rough 6K-token budget (~4 chars/token)
MAX_SAMPLES_PER_COL = 8
MAX_FIRST_ROWS = 8

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

    # Auto-infer a prompt if the user didn't supply one — helps the model
    # focus on the right narrative angle for this workbook.
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

    # Adaptive truncation if too large
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


def _col_payload(c) -> dict:
    out = {"name": c.name, "dtype": c.dtype,
           "n_unique": c.n_unique, "fill_ratio": round(c.fill_ratio, 3),
           "samples": c.samples[:MAX_SAMPLES_PER_COL]}
    for k in ("min", "max", "mean", "sum"):
        v = getattr(c, k, None)
        if v is not None:
            out[k] = v
    if c.top_values:
        out["top_values"] = c.top_values[:MAX_SAMPLES_PER_COL]
    return out


def _block_payload(b: Block) -> dict:
    return {"id": b.id, "kind": b.kind, "label": b.label,
            "provenance": {"sheet": b.provenance.sheet,
                            "columns": b.provenance.columns,
                            "rows": list(b.provenance.rows)},
            "quality_flags": sorted(b.quality_flags),
            **{k: v for k, v in b.extra.items() if k != "raw"}}


PROMPT_TEMPLATE = """Eres un director de arte y analista que diseña presentaciones ejecutivas tipo McKinsey/BCG a partir de datos reales de Excel. Tu trabajo es producir un deck **rico, completo y narrativamente coherente**, con insights que se sienten como hechos por un analista senior.

REGLAS CRÍTICAS:

0a. **Cantidad de slides — RESPETO LITERAL**: si `user_intent.slide_count` está presente y NO es null, debes producir EXACTAMENTE ese número de slides totales (incluyendo portada). NO menos, NO más. Si está null, aplica la regla 1 (Volumen).

0b. **Hojas obligatorias — RESPETO LITERAL**: si `user_intent.required_sheets` no está vacío, cada hoja listada debe aparecer como `provenance.sheet` en al menos UNA slide. Si una hoja listada no tiene datos suficientes para construir una slide válida, OMÍTELA y al final del JSON añade `"_skipped_required_sheets": ["nombre"]`. Nunca inventes datos para forzar la cobertura.

1. **Volumen** (solo cuando `user_intent.slide_count` es null): produce entre 7 y 11 slides. Una portada + 6-10 slides de contenido. NO menos.

2. **Procedencia obligatoria**: cada slide DEBE referenciar bloques existentes por su `id` (`block_ref`, `block_refs`, o `supports_block`). NUNCA inventes IDs.

3. **CERO alucinaciones**: ningún número, nombre, fecha, porcentaje en el JSON puede ser inventado. Si va en un `narrative` o `bullets`, debe existir en el bloque referenciado (`samples`, `top_values`, `min/max/mean/sum/value`, o `first_rows`).

4. **Bullets específicos y abundantes**: cada slide `text_bullets` debe tener **5-7 bullets**, cada uno con al menos una cifra concreta o un nombre propio del bloque. PROHIBIDO bullets genéricos como "los datos muestran variabilidad" o "es importante revisar".

5. **Narrativas tipo analista (NO descriptivas)**: cada `narrative` de un chart debe explicar QUÉ pasa Y POR QUÉ importa. Incluir mínimo 2 datos concretos y al menos uno de estos ángulos. Los ejemplos siguientes son PATRONES de FORMATO — sustituye SIEMPRE los datos por los del Excel cargado:
   - **Concentración** (patrón: "top 3 <categoría> = X% del total → riesgo de dependencia"),
   - **Outlier** (patrón: "<entidad> concentra N, Yx la mediana → caso atípico"),
   - **Brecha** (patrón: "solo X% aprobado vs Y% pendiente → cuello de botella"),
   - **Tendencia** (patrón: "decreció X% entre <periodo1> y <periodo2> → caída sostenida").

6. **Estructura recomendada** (ordenada para máximo impacto narrativo):
   - 1 slide `title` (portada)
   - 1 slide `kpi_row` con 2-4 KPIs principales. PREFIERE en este orden: (a) bloques con `quality_flags: ["effective_total"]` (suma filtrada por estado — el dinero **realmente ejecutado**, sin rechazados); (b) bloques `["derived"]` (% derivados); (c) currency totales SIN flag `subsumed_by_total`; (d) NUNCA pongas en headline un KPI `["needs_status_disclaimer"]` si existe su contrapartida `["effective_total"]` para la misma columna — esos `needs_status_disclaimer` son sumas brutas que **incluyen registros rechazados** y son engañosos como headline. Si los usas, debe ser solo en una bullet/narrative que **mencione explícitamente el desglose por estado** (cuántos rechazados, cuántos efectivos).
   - **REGLA ANTI-ALUCINACIÓN-FINANCIERA**: cuando un bloque KPI tiene `status_filter` en su `extra`, su `value` es la SUMA CRUDA de toda la columna (incluye estados negativos como RECHAZADO). NO digas que es "lo que la empresa movió", "lo ejecutado", "lo desembolsado". El total real ejecutado está en el bloque hermano con flag `effective_total`. Si necesitas mencionar el bruto, di "bruto solicitado" o "valor total solicitado (incluye rechazados)" y cita el desglose desde `status_filter.positive_sum`, `status_filter.negative_sum`.
   - **FORMATO DE MONEDA OBLIGATORIO — NUNCA uses "B"**: en español "billón" = 10¹² (mil veces más grande que la "B" americana = 10⁹), por lo que "$1.5B" sería leído como "$1,500,000,000,000". Para CUALQUIER monto monetario ≥ 1,000,000,000 (mil millones), exprésalo en millones con separador de miles: "$1,504M", "$258M", "$1,226M". Para montos entre 1M y 999M usa "$X.XM" (ej. "$412.7M"). Para entre 1K y 999K usa "$X.XK". Esta regla aplica a NARRATIVES, BULLETS y CUALQUIER texto generado.
   - **PROHIBIDO ATRIBUIR EXTREMOS A REGISTROS QUE NO EXISTAN JUNTOS**: si dices "el máximo de columna C es $X (atribución)", la atribución (persona/ciudad/estado) debe venir de la MISMA fila donde está $X — NO mezcles "máximo de C" con "top de otra columna D". Si las bullets pre-generadas (`insight_bullets`) ya traen el monto con `(Persona, Ciudad, Estado)` entre paréntesis, ÚSALAS TAL CUAL — no las recombines. Si no traes la atribución completa, omite la atribución en lugar de inventarla.
   - 2-4 slides `chart` con distintos cortes (mezcla `bar`, `pie`, `line` según los `kind` disponibles)
   - 1-2 slides `table` con cortes detallados (T*) — usa `columns_subset` con solo 4-6 columnas relevantes (NO IDs, NO observaciones largas)
   - 1-2 slides `text_bullets` con hallazgos accionables (5-7 bullets con cifras; cada bullet debe incluir un dato concreto del bloque referenciado)

7. **Variedad visual**: si tienes ≥2 distribuciones categóricas, usa una `bar` y una `pie`. Si hay serie temporal `S*`, dale un `line`.

8. **Calidad de KPIs**: prefiere bloques `K*` cuyo nombre **no** contenga "id", "código", "número", "folio", "consecutivo". Esos son identificadores, no KPIs.

8b. **Hojas autoritativas vs muestras** — REGLA DE FUENTE: si un bloque (K, C, T, S) tiene `quality_flags` que incluye `"redundant_sheet"`, su hoja es una MUESTRA o COPIA de otra hoja más grande del mismo libro. PROHIBIDO usar bloques `redundant_sheet` en charts, KPIs o tablas cuando exista otro bloque NO redundante con la misma columna semántica (ej.: "Ciudad Destino" en Hoja1 (31 filas, redundant) vs "Comisiones- Base" (1852 filas, autoritativa) → usa SIEMPRE Comisiones- Base). Decir "Bogotá lidera con 19 (61%)" cuando el universo real es de 1852 registros donde Bogotá tiene 252 es un error de auditoría grave. Solo está permitido usar un bloque `redundant_sheet` si NO hay alternativa con esa misma columna en otra hoja.

9. **Títulos editorial**: títulos atractivos en español ("Concentración geográfica de destinos", no "Gráfica de Datos"). Subtítulos opcionales con contexto.

10. **JSON estricto y único**: devuelve EXCLUSIVAMENTE un JSON válido (sin markdown, sin texto adicional) con esta forma exacta:

{{
  "presentation_meta": {{ "title": "Título atractivo del deck", "subtitle": "Subtítulo con contexto" }},
  "slides": [
    {{ "type": "title", "data": {{ "title": "Título portada", "subtitle": "Bajada" }} }},
    {{ "type": "kpi_row", "title": "Indicadores Clave", "block_refs": ["K1","K2","K3"] }},
    {{ "type": "chart", "chart_type": "bar", "title": "Distribución por <Categoría>",
       "block_ref": "C1", "narrative": "<top-categoría> concentra N registros (X%), seguida por <2da> y <3era>." }},
    {{ "type": "chart", "chart_type": "pie", "title": "Composición por <Categoría>",
       "block_ref": "C2", "narrative": "El X% está en estado <ESTADO-DOMINANTE>, mientras Y% está <ESTADO-2DO>." }},
    {{ "type": "chart", "chart_type": "histogram", "title": "Distribución de <Columna Numérica>",
       "block_ref": "K2", "narrative": "La mayoría se concentra entre $X y $Y; cola larga hasta $Z." }},
    {{ "type": "table", "title": "Detalle Operativo", "block_ref": "T1",
       "columns_subset": ["<col1>","<col2>","<col3>","<col4>"], "max_rows": 10 }},
    {{ "type": "text_bullets", "title": "Hallazgos Accionables", "supports_block": "T1",
       "bullets": [
         "El monto máximo individual asciende a $<X> en <fila real del max>.",
         "Existen <N> registros en estado <ESTADO-NEGATIVO> que requieren revisión.",
         "La concentración se ubica en <top-entidad> con el mayor número de registros.",
         "El total acumulado de <Columna> suma $<sum>.",
         "<top-entidad> representa el X% del volumen, evidenciando concentración.",
         "El monto promedio por registro es $<mean>, con <N> outliers sobre 2x ese valor."
       ] }}
  ],
  "prompt_suggestions": [
    "Detalla el análisis por <dimensión temporal>",
    "Compara métricas por <dimensión categórica>",
    "Resalta los 5 <entidades> con mayor <métrica>"
  ]
}}

**IMPORTANTE — los `<placeholders>` arriba son del template; SIEMPRE reemplázalos con los datos reales del Excel cargado. Nunca dejes `<...>` literal en tu output.**

Datos disponibles:
{payload_json}
"""

def _format_money(v) -> str:
    """Formato de moneda compacto en convención hispana (sin "B"; valores
    grandes se expresan en millones con separador): 1500000 → '$1.5M',
    1504462283 → '$1,504M', 12345 → '$12K'. Delega en `format_compact`
    para mantener una sola fuente de verdad cross-módulo."""
    from socya_pipeline.insights import format_compact
    return format_compact(v, prefix="$", fallback="$0")


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

    # Load the actual dataframe so max/min generators can include the full
    # row attribution (Solicitante / Ciudad / Estado). Without this, the AI
    # tends to invent a city for the "$8.1M max" by picking the top city of
    # an unrelated distribution — a real hallucination caught during audit.
    df_for_extremes = _load_sheet_df_for_attribution(wb, sheet.name)

    # Generators (each appends 0 or 1 bullet to `out`)
    _gen_total(out, money_cols, sheet)
    _gen_count(out, sheet)
    _gen_top_cat(out, cat_cols)
    _gen_effective_breakdown(out, money_cols, sheet)
    _gen_max(out, money_cols, df_for_extremes)
    _gen_min(out, money_cols, df_for_extremes)
    _gen_second_cat(out, cat_cols)
    _gen_pareto(out, cat_cols)
    _gen_outlier(out, cat_cols)

    return out[:max_candidates]


# Columns whose value at the extreme row makes the bullet self-contained
# (the reader sees what record holds the max/min, where, and its status).
_ATTRIBUTION_COLUMN_HINTS = (
    "solicitante", "responsable", "encargado", "beneficiario",
    "ciudad destino", "ciudad", "destino", "estado", "status",
    "centro de costos", "area", "área", "proyecto", "categor",
)


def _load_sheet_df_for_attribution(wb, sheet_name: str):
    """Read just the sheet that the extreme bullet will quote, so we can
    pull the full row at argmax/argmin. Best-effort: returns None on any
    failure (the generator then falls back to the pre-fix behaviour)."""
    try:
        import pandas as pd
        path = getattr(wb, "filename", None)
        if not path:
            return None
        xl = pd.ExcelFile(path)
        if sheet_name not in xl.sheet_names:
            return None
        # Heuristic for header row: first row that contains a non-Unnamed
        # string in column 0. Matches the parser's behaviour without
        # depending on it (the parser does deeper cleaning we don't need
        # for row-level attribution).
        df_h0 = pd.read_excel(path, sheet_name=sheet_name, header=0)
        if all(str(c).startswith("Unnamed") or str(c).strip().lower() in ("nan",)
               for c in df_h0.columns):
            df_h1 = pd.read_excel(path, sheet_name=sheet_name, header=1)
            return df_h1
        return df_h0
    except Exception:
        return None


def _row_attribution(df, idx) -> str:
    """Pick the most informative columns from a row and format them as
    a compact parenthetical: '(Miguel Angel Gomez Torres, Atlantico - Sabanalarga, RECHAZADO)'.
    """
    try:
        if df is None or idx is None or idx not in df.index:
            return ""
    except Exception:
        return ""
    parts = []
    for col in df.columns:
        name_l = str(col).strip().lower()
        if not any(h in name_l for h in _ATTRIBUTION_COLUMN_HINTS):
            continue
        try:
            v = df.at[idx, col]
        except Exception:
            continue
        if v is None:
            continue
        s = str(v).strip()
        if not s or s.lower() in ("nan", "none", "null"):
            continue
        parts.append(s)
        if len(parts) >= 3:
            break
    return f" ({', '.join(parts)})" if parts else ""


def _gen_total(out: list, money_cols: list, sheet=None) -> None:
    if not money_cols:
        return
    top = money_cols[0]
    # When the sheet has a status filter for this column, prefer the
    # honest "efectivo" number as the headline bullet — otherwise the
    # planner happily cites $1.5B "total" while $258M of that was rejected.
    bd = ((sheet.status_breakdowns or {}).get(top.name)
          if sheet is not None else None)
    if bd and bd.get("negative_count", 0) > 0 and (bd.get("positive_sum") or 0) > 0:
        pos = _format_money(bd["positive_sum"])
        gross = _format_money(bd["gross_sum"])
        neg = _format_money(bd["negative_sum"])
        neg_c = bd.get("negative_count") or 0
        out.append(("total",
            f"El total efectivo de '{top.name}' suma {pos} "
            f"(de un bruto solicitado de {gross}; {neg} en {neg_c} registros "
            f"rechazados o anulados)."))
        return
    out.append(("total",
        f"El total acumulado de '{top.name}' suma {_format_money(top.sum)}."))


def _gen_effective_breakdown(out: list, money_cols: list, sheet) -> None:
    """Adds a second bullet enumerating the per-state slices of the most
    relevant currency column. Only fires when the sheet has a status filter
    with at least one negative state — that's exactly when readers benefit
    from seeing the deck not hide the rejected slice."""
    if sheet is None or not money_cols:
        return
    breakdowns = getattr(sheet, "status_breakdowns", None) or {}
    if not breakdowns:
        return
    top = money_cols[0]
    bd = breakdowns.get(top.name)
    if not bd or bd.get("negative_count", 0) == 0:
        return
    parts = []
    for bucket in ("positive", "negative", "pending"):
        cnt = bd.get(f"{bucket}_count") or 0
        if cnt == 0:
            continue
        s = _format_money(bd.get(f"{bucket}_sum") or 0)
        verb = {"positive": "ejecutados",
                 "negative": "rechazados",
                 "pending":  "en trámite"}[bucket]
        parts.append(f"{cnt} {verb} ({s})")
    if len(parts) >= 2:
        out.append(("status_breakdown",
            f"Desglose de '{top.name}' por estado: " + "; ".join(parts) + "."))


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


def _gen_max(out: list, money_cols: list, df=None) -> None:
    if not money_cols or money_cols[0].max is None:
        return
    top = money_cols[0]
    attr = ""
    if df is not None and top.name in df.columns:
        try:
            import pandas as pd
            s = pd.to_numeric(df[top.name], errors="coerce")
            idx = s.idxmax()
            attr = _row_attribution(df, idx)
        except Exception:
            attr = ""
    out.append(("max",
        f"El valor máximo registrado en '{top.name}' es "
        f"{_format_money(top.max)}{attr}."))


def _gen_min(out: list, money_cols: list, df=None) -> None:
    if not money_cols or money_cols[0].min is None:
        return
    top = money_cols[0]
    attr = ""
    if df is not None and top.name in df.columns:
        try:
            import pandas as pd
            s = pd.to_numeric(df[top.name], errors="coerce")
            idx = s.idxmin()
            attr = _row_attribution(df, idx)
        except Exception:
            attr = ""
    out.append(("min",
        f"El valor mínimo en '{top.name}' es "
        f"{_format_money(top.min)}{attr}."))


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


def _deterministic_plan_fallback(wb: WorkbookData, blocks,
                                   audience: str, language: str) -> dict:
    """Plan determinístico derivado del inventory cuando la AI satura o
    devuelve garbage. Antes esto causaba AI_SATURATED que bloqueaba al
    usuario; ahora genera un deck simple pero completamente funcional.

    Incluye bullets construidos con _build_insight_bullets() que citan
    stats de block.extra (sum/min/max/value/top_values) — eso garantiza
    que pasen el validator de provenance sin caer."""
    kpi_blocks = [b for b in blocks if b.kind == "kpi_candidate"][:4]
    chart_cats = [b for b in blocks if b.kind == "categorical_distribution"]
    chart_ts = [b for b in blocks if b.kind == "time_series_candidate"]
    tables = [b for b in blocks if b.kind == "table"]

    base_name = (wb.filename or "Reporte").rsplit(".", 1)[0]
    slides: list = [
        {"type": "title",
         "data": {"title": f"Reporte: {base_name}",
                  "subtitle": "Análisis automático del archivo cargado"}},
    ]

    if kpi_blocks:
        slides.append({
            "type": "kpi_row",
            "title": "Indicadores Clave",
            "block_refs": [b.id for b in kpi_blocks],
        })

    if chart_cats:
        b = chart_cats[0]
        col = b.provenance.columns[0] if b.provenance.columns else "categorías"
        slides.append({
            "type": "chart",
            "chart_type": "bar",
            "title": f"Distribución por {col}",
            "block_ref": b.id,
            # Vacío → el extractor inyecta narrativa honesta auto-generada
            # con _auto_chart_narrative (mejor que cualquier template hardcoded).
            "narrative": "",
        })

    if chart_ts:
        b = chart_ts[0]
        slides.append({
            "type": "chart",
            "chart_type": "line",
            "title": "Evolución temporal",
            "block_ref": b.id,
            "narrative": "",
        })

    if len(chart_cats) > 1:
        b = chart_cats[1]
        col = b.provenance.columns[0] if b.provenance.columns else "categorías"
        slides.append({
            "type": "chart",
            "chart_type": "pie",
            "title": f"Composición de {col}",
            "block_ref": b.id,
            "narrative": "",
        })

    if tables:
        b = tables[0]
        slides.append({
            "type": "table",
            "title": "Detalle Operativo",
            "block_ref": b.id,
            "max_rows": 10,
        })

        # Slide final de hallazgos. Si los bullets fueran inválidos, el validator
        # los dropea por provenance y la slide queda vacía → el extractor a su
        # vez la dropea por bullets_empty. Solo añadimos la slide si tenemos
        # bullets que el validator probablemente acepta.
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

    return {
        "presentation_meta": {
            "title": f"Reporte: {base_name}",
            "subtitle": "Generado automáticamente desde Excel",
        },
        "slides": slides,
        "prompt_suggestions": [
            "Detalla el análisis por categoría principal",
            "Compara métricas por período",
            "Resalta los hallazgos más relevantes",
        ],
        "_meta": {
            "model": "deterministic_fallback",
            "cache_hit": False,
            "fallback_steps": [],
            "planner_version": PLANNER_VERSION,
            "deterministic_fallback": True,
        },
    }


def _enforce_intent(plan: dict, intent, blocks, wb) -> dict:
    """Garantiza que `plan` honre `intent`. Idempotente.

    1. Inyecta required_sheets faltantes (puede crecer el deck).
    2. Recorta o expande para llegar a `requested_slide_count`.
    3. Anota _intent_report con lo que se honró y lo que no.
    """
    slides = list(plan.get("slides") or [])
    blocks_by_id = {b.id: b for b in blocks}

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

    target = intent.requested_slide_count
    count_honored = True
    if target is not None:
        if len(slides) > target:
            slides = _trim_to(slides, target, intent)
        elif len(slides) < target:
            padded = _pad_to(slides, target, blocks, wb, blocks_by_id)
            slides = padded
            if len(slides) < target:
                count_honored = False

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
    """Recorta `slides` a `target` preservando: title, slides inyectadas,
    slides cuya provenance.sheet está en required_sheets, luego por tipo."""
    if len(slides) <= target:
        return slides

    required_sheet_names = {m.matched for m in intent.required_sheets if m.matched}

    def _slot_priority(s_idx_pair):
        idx, s = s_idx_pair
        if s.get("type") == "title":
            return 0
        if s.get("_injected_by") == "intent_enforcement":
            return 1
        prov_sheet = s.get("provenance", {}).get("sheet")
        if prov_sheet in required_sheet_names:
            return 2
        type_pref = {"kpi_row": 3, "chart": 4, "table": 5,
                     "text_bullets": 6}.get(s.get("type"), 7)
        return type_pref

    indexed = list(enumerate(slides))
    indexed.sort(key=_slot_priority)
    kept_indices = sorted(idx for idx, _ in indexed[:target])
    return [slides[i] for i in kept_indices]


def _pad_to(slides: list, target: int, blocks, wb, blocks_by_id) -> list:
    """Añade slides desde bloques no usados hasta llegar a `target`."""
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


def plan_presentation(wb: WorkbookData, blocks, user_prompt: str, audience: str,
                       language: str, api_key: str,
                       profile: AIProfile = AIProfile.PATIENT,
                       cache_dir: Optional[Path] = None,
                       file_path: Optional[Path] = None,
                       intent=None,
                       preferred_model: Optional[str] = None) -> dict:
    cache_key = None
    cache: Optional[PlanCache] = None
    if file_path is not None:
        try:
            file_bytes = Path(file_path).read_bytes()
            intent_signature = ""
            if intent is not None:
                req_sheets = ",".join(sorted(
                    m.matched for m in intent.required_sheets if m.matched
                ))
                intent_signature = (f"|sc={intent.requested_slide_count}"
                                     f"|rs={req_sheets}"
                                     f"|pm={preferred_model or ''}")
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

    # AI con fallback determinístico: si la AI satura o devuelve garbage,
    # devolvemos un plan derivado del inventory en vez de un error que deja
    # al usuario bloqueado. NO cacheamos el fallback — queremos reintentar
    # AI en el próximo request por si vuelve a estar disponible.
    chain = AIChain(api_key=api_key, profile=profile,
                     preferred_model=preferred_model)
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
        # Some models emit `"slides": null` or `"slides": "TBD"` when they
        # think the input is too small / ambiguous. Without this guard we'd
        # crash later in `validate_plan` with a TypeError on iteration.
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

    if intent is not None:
        parsed = _enforce_intent(parsed, intent, blocks, wb)

    if cache and cache_key:
        cache.set(cache_key, parsed)

    return parsed


