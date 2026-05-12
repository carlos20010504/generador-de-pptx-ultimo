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

PLANNER_VERSION = "p3"  # bump when prompt template changes — invalidates cache

MAX_PAYLOAD_CHARS = 24_000  # rough 6K-token budget (~4 chars/token)
MAX_SAMPLES_PER_COL = 8
MAX_FIRST_ROWS = 8

def build_payload(wb: WorkbookData, blocks, user_prompt: str, audience: str,
                   language: str) -> dict:
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

    payload = {
        "user_prompt": effective_prompt,
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

1. **Volumen**: produce entre 7 y 11 slides. Una portada + 6-10 slides de contenido. NO menos.

2. **Procedencia obligatoria**: cada slide DEBE referenciar bloques existentes por su `id` (`block_ref`, `block_refs`, o `supports_block`). NUNCA inventes IDs.

3. **CERO alucinaciones**: ningún número, nombre, fecha, porcentaje en el JSON puede ser inventado. Si va en un `narrative` o `bullets`, debe existir en el bloque referenciado (`samples`, `top_values`, `min/max/mean/sum/value`, o `first_rows`).

4. **Bullets específicos**: cada bullet DEBE contener al menos una cifra concreta o un nombre propio del bloque. PROHIBIDO bullets genéricos como "los datos muestran variabilidad" o "es importante revisar".

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
   - 1-2 slides `text_bullets` con hallazgos accionables (bullets con cifras; cada bullet debe incluir un dato concreto del bloque referenciado)

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
         "La concentración geográfica se ubica en MEDELLÍN con el mayor número de registros."
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

def plan_presentation(wb: WorkbookData, blocks, user_prompt: str, audience: str,
                       language: str, api_key: str,
                       profile: AIProfile = AIProfile.PATIENT,
                       cache_dir: Optional[Path] = None,
                       file_path: Optional[Path] = None) -> dict:
    cache_key = None
    cache: Optional[PlanCache] = None
    if file_path is not None:
        try:
            file_bytes = Path(file_path).read_bytes()
            cache_key = compute_cache_key(file_bytes, user_prompt, audience,
                                            language, PLANNER_VERSION)
            cache = PlanCache(cache_dir=cache_dir)
            cached = cache.get(cache_key)
            if cached:
                cached["_meta"] = {**cached.get("_meta", {}),
                                    "cache_hit": True}
                return cached
        except OSError:
            pass

    payload = build_payload(wb, blocks, user_prompt, audience, language)
    payload_json = json.dumps(payload, ensure_ascii=False, default=str)
    prompt = PROMPT_TEMPLATE.format(payload_json=payload_json)

    chain = AIChain(api_key=api_key, profile=profile)
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

    parsed["_meta"] = {
        "model": result.model,
        "cache_hit": False,
        "fallback_steps": result.fallback_steps,
        "planner_version": PLANNER_VERSION,
    }

    if cache and cache_key:
        cache.set(cache_key, parsed)

    return parsed


