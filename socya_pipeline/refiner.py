"""Optional second-pass AI critique of the plan before rendering.

Por qué existe: el planner es single-shot — emite el plan completo en una
sola pasada. A veces sale ruido: 2 slides con la misma narrativa, títulos
genéricos, slides marginales que el usuario va a borrar igual. Este
refinador le da al AI una segunda oportunidad de "tachar" sus errores
mirando el plan ya validado + auditado.

Es BEST-EFFORT y opt-in — controlado por env var SOCYA_AI_REFINE=1. Si
falla, devolvemos el plan intacto. Nunca rompe la generación.

Diseño:
  1. Tras `validate_plan` (ya filtró slides imposibles), construimos un
     resumen mínimo del plan (no todo el JSON, sólo {idx, type, title,
     provenance}).
  2. Lo mandamos al AI con prompt corto: "¿Fixes?".
  3. Parseamos respuesta JSON con shape {fixes: [{slide_index, action,
     ...}]}. Acciones: drop / retitle / reorder.
  4. Aplicamos defensive (no permitir drop del title, no permitir reorder
     fuera de rango, ignorar si fixes > 50% de slides).
"""
from __future__ import annotations

import json
import os
from typing import List, Optional

from socya_pipeline import insights


REFINE_ENABLED_ENV = "SOCYA_AI_REFINE"


def is_enabled() -> bool:
    """Default OFF para no afectar costos del usuario silenciosamente."""
    return os.environ.get(REFINE_ENABLED_ENV, "").strip().lower() in ("1", "true", "yes")


_REFINER_PROMPT = """Eres un revisor crítico de presentaciones ejecutivas.

Te paso el PLAN de un deck ya validado. Tu trabajo es proponer FIXES
QUIRÚRGICOS:

- DROP: slides duplicadas, redundantes o sin valor (ej. dos slides con la
  misma fuente y narrativa).
- RETITLE: títulos genéricos, no específicos, o que no reflejan el dato.
  El nuevo título debe ser CONCRETO: si el dato es "concentración por
  ciudad", no digas "Análisis"; di "Concentración por ciudad".
- REORDER: si una slide encaja mejor en otra posición temática.

REGLAS:
- NO toques la slide 0 (portada).
- NO propongas más de 3 fixes en total.
- Si el plan está bien, devuelve {"fixes": []}.
- Sólo emite JSON estricto, sin texto antes ni después.

Schema:
{
  "fixes": [
    {"slide_index": int, "action": "drop", "reason": "<corta>"},
    {"slide_index": int, "action": "retitle", "new_title": "<máx 80 chars>", "reason": "<corta>"},
    {"slide_index": int, "action": "reorder", "new_position": int, "reason": "<corta>"}
  ]
}

Plan a revisar:
{plan_json}
"""


def _slim_plan(plan: dict) -> List[dict]:
    """Construir el resumen mínimo del plan que mandamos al AI. Mantener
    bajo: el AI no necesita los `data` completos para decidir fixes."""
    slim = []
    for i, s in enumerate(plan.get("slides", [])):
        slim.append({
            "idx": i,
            "type": s.get("type"),
            "title": s.get("title"),
            "subtitle": s.get("subtitle"),
            "narrative": (s.get("narrative") or "")[:200],
            "provenance": s.get("provenance"),
        })
    return slim


def refine_plan(plan: dict, api_key: Optional[str] = None,
                  profile=None) -> dict:
    """Devuelve el plan modificado con fixes aplicados. Nunca lanza
    excepciones — si algo sale mal, devuelve el plan original."""
    if not is_enabled():
        return plan
    if not isinstance(plan, dict) or not plan.get("slides"):
        return plan
    if not api_key:
        return plan

    try:
        from socya_pipeline.ai_chain import AIChain, AIProfile
        slim = _slim_plan(plan)
        if len(slim) < 3:
            # Decks pequeños no se benefician — el AI tendería a dropear
            # algo y dejar al usuario sin nada.
            return plan
        prompt = _REFINER_PROMPT.format(plan_json=json.dumps(slim, ensure_ascii=False))
        chain = AIChain(api_key=api_key, profile=profile or AIProfile.FAST)
        result = chain.call(prompt, system_msg="Output strictly valid JSON.",
                              temperature=0.1)
        parsed = insights.parse_loose_json(result.content)
        if not isinstance(parsed, dict):
            return plan
        fixes = parsed.get("fixes")
        if not isinstance(fixes, list):
            return plan
        return _apply_fixes(plan, fixes)
    except Exception:
        # Refiner es polish — fallar silenciosamente y entregar el plan
        # original es preferible a romper la generación.
        return plan


def _apply_fixes(plan: dict, fixes: list) -> dict:
    """Aplica fixes con validación defensiva. Nunca toca la slide 0
    (portada). Limita a 3 fixes para evitar sobrecorrección."""
    slides = list(plan.get("slides", []))
    n = len(slides)
    if n < 2:
        return plan

    fixes = [f for f in fixes if isinstance(f, dict)][:3]
    if not fixes:
        return plan

    # Sanidad — si el AI quiere dropear más del 50%, ignoramos todo.
    drops_proposed = sum(1 for f in fixes if f.get("action") == "drop")
    if drops_proposed > n // 2:
        return plan

    # 1) RETITLE in-place (no afecta orden)
    for f in fixes:
        if f.get("action") != "retitle":
            continue
        idx = f.get("slide_index")
        new_title = (f.get("new_title") or "").strip()
        if not isinstance(idx, int) or idx <= 0 or idx >= n:
            continue
        if not new_title or len(new_title) > 140:
            continue
        slides[idx]["title"] = new_title

    # 2) DROP — marcar índices a remover (excepto idx 0)
    drop_set = set()
    for f in fixes:
        if f.get("action") != "drop":
            continue
        idx = f.get("slide_index")
        if isinstance(idx, int) and 0 < idx < n:
            drop_set.add(idx)

    # 3) REORDER — construye orden propuesto. Title slide siempre primero.
    desired_order: List[int] = list(range(n))
    for f in fixes:
        if f.get("action") != "reorder":
            continue
        idx = f.get("slide_index")
        new_pos = f.get("new_position")
        if not (isinstance(idx, int) and isinstance(new_pos, int)):
            continue
        if idx <= 0 or idx >= n or new_pos <= 0 or new_pos >= n:
            continue
        # Mover idx a new_pos en desired_order
        if idx in desired_order:
            desired_order.remove(idx)
            insert_at = min(max(new_pos, 1), len(desired_order))
            desired_order.insert(insert_at, idx)

    # Construir nueva lista respetando drops + orden
    new_slides = []
    for i in desired_order:
        if i in drop_set:
            continue
        new_slides.append(slides[i])

    if not new_slides:
        # Algo se descontroló — fallback al plan original
        return plan

    out = dict(plan)
    out["slides"] = new_slides
    # Marcar en _meta que hubo refinamiento (lo verá el audit JSON)
    meta = dict(out.get("_meta") or {})
    meta["refined"] = True
    meta["refine_fixes_applied"] = len(fixes)
    meta["refine_dropped"] = len(drop_set)
    out["_meta"] = meta
    return out
