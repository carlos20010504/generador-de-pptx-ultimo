"""Single-call IA planner. Takes inventory + prompt, returns full slide plan."""
import json
import re
from pathlib import Path
from typing import Optional
from socya_pipeline.ai_cache import PlanCache, compute_cache_key
from socya_pipeline.ai_chain import AIChain, AIProfile
from socya_pipeline.errors import PipelineError, ErrorCode
from socya_pipeline.inventory import Block
from socya_pipeline.parser import WorkbookData

PLANNER_VERSION = "p2"  # bump when prompt template changes — invalidates cache

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

    payload = {
        "user_prompt": user_prompt or "",
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


PROMPT_TEMPLATE = """Eres un director de arte que diseña presentaciones ejecutivas tipo McKinsey/BCG a partir de datos reales de Excel. Tu trabajo es producir un deck **rico, completo y narrativamente coherente**.

REGLAS CRÍTICAS:

1. **Volumen**: produce entre 7 y 11 slides en total. Una sola portada + 6-10 slides de contenido. NO pocos slides.

2. **Procedencia obligatoria**: cada slide DEBE referenciar bloques existentes por su `id` (`block_ref`, `block_refs`, o `supports_block`). NUNCA inventes IDs.

3. **CERO alucinaciones**: ningún número, nombre, fecha, porcentaje en el JSON puede ser inventado. Si va en un `narrative` o `bullets`, debe existir en el bloque referenciado (`samples`, `top_values`, `min/max/mean/sum`, o `first_rows`).

4. **Bullets específicos**: cada bullet DEBE contener al menos una cifra concreta o un nombre propio del bloque. PROHIBIDO bullets genéricos como "los datos muestran variabilidad", "se observa una tendencia positiva", "es importante revisar".

5. **Narrativas con cifras**: en cada `chart`, la `narrative` debe nombrar al menos 2 categorías con sus valores exactos. Ej: "Bogotá lidera con 53 ventas, seguido de Medellín con 23".

6. **Estructura recomendada** (ordenada para máximo impacto):
   - 1 slide `title` (portada)
   - 1 slide `kpi_row` con 2-4 KPIs principales (preferir bloques `K*` numéricos no-ID, sobre todo currency)
   - 2-3 slides `chart` con distintos cortes (diferentes `block_refs` C* o S*)
   - 1-2 slides `table` con cortes detallados (T*) — usa `columns_subset` para mostrar solo las 4-6 columnas más relevantes, no todas
   - 1-2 slides `text_bullets` con hallazgos accionables (`supports_block` = T*, bullets con cifras)

7. **Variedad visual**: si tienes ≥2 distribuciones categóricas, usa una `bar` y una `pie`. Si hay serie temporal, hazle un `line`.

8. **Calidad de KPIs**: prefiere bloques `K*` cuyo nombre **no** contenga "id", "código", "número", "folio", "consecutivo". Esos son identificadores, no KPIs.

9. **Títulos editorial**: títulos atractivos en español ("Tendencia mensual de ingresos", no "Gráfica de Datos"). Subtítulos opcionales con contexto adicional.

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

    parsed = _parse_json_loose(result.content)
    if not isinstance(parsed, dict) or "slides" not in parsed:
        raise PipelineError(
            ErrorCode.AI_RESPONSE_INVALID,
            "El planificador devolvió un JSON sin la forma esperada.",
            details=str(result.content)[:300],
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


def _parse_json_loose(text: str):
    raw = (text or "").strip()
    if not raw:
        return None
    raw = raw.replace("```json", "").replace("```", "").strip()
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if m:
        raw = m.group(0)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None
