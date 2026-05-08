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

PLANNER_VERSION = "p1"  # bump when prompt template changes — invalidates cache

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


PROMPT_TEMPLATE = """Eres un planificador de presentaciones ejecutivas que genera slides a partir de datos reales de Excel.

Reglas estrictas:
1. Cada slide DEBE referenciar bloques existentes por su `id` (campo `block_ref` o `block_refs`).
2. NO inventes números, nombres ni hechos. Si necesitas un dato, viene del bloque referenciado.
3. Si el prompt del usuario pide algo que los bloques no soportan, omite ese slide.
4. Bullets deben ser específicos (con cifras o nombres del bloque). No bullets genéricos como "los datos muestran variabilidad".
5. Devuelve EXCLUSIVAMENTE un JSON válido con esta forma exacta:

{{
  "presentation_meta": {{ "title": "...", "subtitle": "..." }},
  "slides": [
    {{ "type": "title", "title": "...", "subtitle": "..." }},
    {{ "type": "kpi_row", "title": "...", "block_refs": ["K1", "K2"] }},
    {{ "type": "chart", "chart_type": "bar|line|pie", "title": "...",
       "block_ref": "C1", "narrative": "una frase basada en el bloque" }},
    {{ "type": "table", "title": "...", "block_ref": "T1",
       "columns_subset": ["..."], "max_rows": 12 }},
    {{ "type": "text_bullets", "title": "...", "supports_block": "T1",
       "bullets": ["bullet con cifra", "bullet con nombre"] }}
  ],
  "prompt_suggestions": ["sugerencia 1", "sugerencia 2", "sugerencia 3"]
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
