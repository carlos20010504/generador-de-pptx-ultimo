"""CLI entry: `python -m socya_pipeline {analyze|generate} <args>`."""
import argparse
import json
import re
import sys
import os
from pathlib import Path
from socya_pipeline.errors import PipelineError, ErrorCode
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.planner import plan_presentation
from socya_pipeline.validator import validate_plan
from socya_pipeline.extractor import extract_for_render, auto_complete_slides
from socya_pipeline.ai_chain import AIProfile

# Strip absolute paths (Windows + POSIX) and home dirs from any user-facing
# error message. They leak filesystem layout and are useless to the user.
_PATH_REDACTORS = (
    re.compile(r"[A-Za-z]:\\[^\s'\"]+"),    # C:\foo\bar
    re.compile(r"/(?:home|Users|tmp|var)/[^\s'\"]+"),  # /home/x, /Users/x, etc.
)


def _sanitize_details(text) -> str:
    s = str(text or "")
    for pat in _PATH_REDACTORS:
        s = pat.sub("[ruta]", s)
    return s[:300]


def _emit_error(err: PipelineError):
    payload = err.to_dict()
    if isinstance(payload, dict) and payload.get("details"):
        payload["details"] = _sanitize_details(payload["details"])
    sys.stdout.write(json.dumps({"error": payload}, ensure_ascii=True))
    sys.exit(2)

def _load_request(raw: str) -> dict:
    try:
        return json.loads(raw or "{}")
    except json.JSONDecodeError:
        return {}

def cmd_plan(args):
    request = _load_request(args.request)
    api_key = _resolve_api_key()
    profile = (AIProfile.PATIENT if (os.environ.get("SOCYA_AI_PROFILE", "fast")
                                       .lower() == "patient")
                 else AIProfile.FAST)
    # Inicializo afuera del try para que el finally siempre vea xls (None si
    # nunca llegamos a abrirlo). Sin esto, Windows bloquea el archivo entre
    # uploads del mismo Excel hasta que el GC lo libere.
    xls = None
    try:
        wb = parse_workbook(args.input, api_key=api_key)
        inv = build_inventory(wb)
        plan = plan_presentation(
            wb, inv,
            user_prompt=request.get("prompt", ""),
            audience=request.get("audience", "ejecutivos"),
            language=request.get("language", "es"),
            api_key=api_key,
            profile=profile,
            file_path=Path(args.input),
        )
        outcome = validate_plan(plan, inv, wb)
        if not outcome.slides:
            raise PipelineError(
                ErrorCode.EXCEL_INSUFFICIENT_DATA,
                f"El Excel '{wb.filename}' no tiene datos suficientes para "
                "generar slides validados.",
                details=f"Slides planeados: {len(plan.get('slides', []))}, "
                f"descartados: {len(outcome.dropped)}.",
                user_action="improve_excel_or_change_prompt",
            )
        import pandas as pd
        from socya_pipeline.extractor import _build_dtype_map
        xls = pd.ExcelFile(Path(args.input))
        sheets_cache: dict = {}
        dtype_map = _build_dtype_map(wb)

        rendered, extraction_dropped = extract_for_render(
            outcome.slides, inv, wb, args.input,
            xls=xls, sheets_cache=sheets_cache, dtype_map=dtype_map)
        before_complete = len(rendered)
        rendered = auto_complete_slides(
            rendered, inv, wb, args.input, target_count=7,
            xls=xls, sheets_cache=sheets_cache, dtype_map=dtype_map)
        result = {
            "presentation_meta": plan.get("presentation_meta", {}),
            "slides": rendered,
            "prompt_suggestions": plan.get("prompt_suggestions", []),
            "ai_status": plan.get("_meta", {}),
            "audit": {
                "slides_planned": len(plan.get("slides", [])),
                "slides_validated": len(outcome.slides),
                "slides_dropped": outcome.dropped,
                "extraction_dropped": extraction_dropped,
                "slides_auto_added": len(rendered) - before_complete,
                "slides_final": len(rendered),
                "bullets_dropped": outcome.bullets_dropped,
            },
        }
        sys.stdout.write(json.dumps(result, ensure_ascii=True, default=str))
    except PipelineError as e:
        _emit_error(e)
    except Exception as e:
        _emit_error(PipelineError(ErrorCode.PYTHON_RUNTIME_ERROR,
                                    "Error inesperado en el planificador.",
                                    details=str(e)[:300]))
    finally:
        if xls is not None:
            try: xls.close()
            except Exception: pass

def cmd_generate(args):
    request = _load_request(args.request)
    api_key = _resolve_api_key()
    profile = AIProfile.PATIENT
    xls = None  # ver comentario en cmd_plan
    try:
        wb = parse_workbook(args.input, api_key=api_key)
        inv = build_inventory(wb)
        plan = plan_presentation(
            wb, inv,
            user_prompt=request.get("prompt", ""),
            audience=request.get("audience", "ejecutivos"),
            language=request.get("language", "es"),
            api_key=api_key,
            profile=profile,
            file_path=Path(args.input),
        )
        # Optional second-pass critique (refiner). Off por default; opt-in
        # via env SOCYA_AI_REFINE=1. Best-effort — fallos no rompen la
        # generación. Se aplica sólo en cmd_generate (no en cmd_plan)
        # para que la preview UI muestre lo mismo que el render final
        # cuando el refiner está apagado (caso default).
        from socya_pipeline.refiner import refine_plan
        plan = refine_plan(plan, api_key=api_key, profile=profile)
        outcome = validate_plan(plan, inv, wb)
        if not outcome.slides:
            raise PipelineError(
                ErrorCode.EXCEL_INSUFFICIENT_DATA,
                f"El Excel '{wb.filename}' no produjo slides validados.",
                details=f"Descartados: {len(outcome.dropped)}.",
                user_action="improve_excel_or_change_prompt",
            )
        # Open the Excel ONCE and share state between extraction phases
        # (without this, the file is parsed up to 4 times per request).
        import pandas as pd
        from socya_pipeline.extractor import _build_dtype_map
        xls = pd.ExcelFile(Path(args.input))
        sheets_cache: dict = {}
        dtype_map = _build_dtype_map(wb)

        rendered, extraction_dropped = extract_for_render(
            outcome.slides, inv, wb, args.input,
            xls=xls, sheets_cache=sheets_cache, dtype_map=dtype_map)
        before_complete = len(rendered)
        rendered = auto_complete_slides(
            rendered, inv, wb, args.input, target_count=7,
            xls=xls, sheets_cache=sheets_cache, dtype_map=dtype_map)

        # Optional user-driven filter: drop the slides the UI marked off in
        # the preview step. Indices are 0-based positions in the final list.
        # The title slide (index 0) is always kept regardless of input.
        excluded_raw = request.get("excludeSlideIndices") or []
        excluded = {int(i) for i in excluded_raw if isinstance(i, (int, float))}
        excluded.discard(0)  # title slide is mandatory
        excluded_count = 0
        if excluded:
            kept = [s for i, s in enumerate(rendered) if i not in excluded]
            excluded_count = len(rendered) - len(kept)
            rendered = kept

        # Optional user override: per-slide title edits. Same indexing
        # convention as `excludeSlideIndices` — positions in the rendered
        # list as the UI saw them in the plan preview.
        title_overrides_raw = request.get("slideTitles") or {}
        if isinstance(title_overrides_raw, dict) and title_overrides_raw:
            title_overrides = {}
            for k, v in title_overrides_raw.items():
                try:
                    title_overrides[int(k)] = str(v)
                except (TypeError, ValueError):
                    continue
            for pos, slide in enumerate(rendered):
                if pos in title_overrides:
                    new_title = title_overrides[pos].strip()
                    if new_title:
                        slide["title"] = new_title

        # Optional user override: reorder. Array of original positions in
        # the order the user wants. Title slide (pos 0) is always kept at
        # position 0 to preserve deck coherence. Falls back to no-op if the
        # input is malformed.
        order_raw = request.get("slideOrder")
        if isinstance(order_raw, list) and order_raw and len(rendered) > 1:
            try:
                order = [int(x) for x in order_raw]
                # Sólo aplicamos si las posiciones cubren TODAS las slides
                # supervivientes — evita reorders parciales que confunden.
                if sorted(order) == list(range(len(rendered))):
                    new_list = []
                    seen = set()
                    # Title slide first if it was at original position 0
                    if 0 not in seen:
                        new_list.append(rendered[0])
                        seen.add(0)
                    for pos in order:
                        if pos in seen:
                            continue
                        if 0 <= pos < len(rendered):
                            new_list.append(rendered[pos])
                            seen.add(pos)
                    if len(new_list) == len(rendered):
                        rendered = new_list
            except (TypeError, ValueError):
                pass

        from socya_pipeline.renderer import render_pptx
        template = Path(args.template)
        # Theme comes through the request from the UI selector. The renderer
        # falls back to its built-in Socya defaults if `theme` is missing.
        theme = request.get("theme") if isinstance(request.get("theme"), dict) else None
        render_pptx(rendered, plan.get("presentation_meta", {}),
                     template_path=template, output_path=Path(args.output),
                     theme=theme)

        # Write audit JSON next to output
        audit = {
            "model_used": plan.get("_meta", {}).get("model"),
            "cache_hit": plan.get("_meta", {}).get("cache_hit", False),
            "fallback_chain_steps": plan.get("_meta", {}).get("fallback_steps", []),
            "refined": plan.get("_meta", {}).get("refined", False),
            "refine_fixes_applied": plan.get("_meta", {}).get("refine_fixes_applied", 0),
            "refine_dropped": plan.get("_meta", {}).get("refine_dropped", 0),
            "slides_planned": len(plan.get("slides", [])),
            "slides_validated": len(outcome.slides),
            "slides_dropped": outcome.dropped,
            "slides_excluded_by_user": excluded_count,
            "extraction_dropped": extraction_dropped,
            "slides_auto_added": len(rendered) - before_complete,
            "slides_final": len(rendered),
            "bullets_dropped": outcome.bullets_dropped,
            "provenance_per_slide": [s.get("provenance") for s in rendered],
        }
        audit_path = Path(args.output).with_suffix(".audit.json")
        audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2),
                                encoding="utf-8")

        sys.stdout.write(json.dumps({"ok": True, "audit": audit}, ensure_ascii=True))
    except PipelineError as e:
        _emit_error(e)
    except Exception as e:
        _emit_error(PipelineError(ErrorCode.PYTHON_RUNTIME_ERROR,
                                    "Error inesperado en la generación.",
                                    details=str(e)[:300]))
    finally:
        if xls is not None:
            try: xls.close()
            except Exception: pass

def _resolve_api_key() -> str:
    """Look up the OpenRouter API key from env or .env files.

    Operate-then-handle (no os.path.exists pre-check) — TOCTOU-safe and one
    syscall less per file probed."""
    key = os.environ.get("OPENROUTER_API_KEY", "").strip()
    if key:
        return key
    for env_file in (".env", ".env.local"):
        try:
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("OPENROUTER_API_KEY="):
                        return line.split("=", 1)[1].strip().strip('"').strip("'")
        except OSError:
            continue
    return ""

def cmd_quick_summary(args):
    """Deterministic onboarding summary — ZERO AI calls. Fast (<1s for 10K
    rows). Used by /api/quick-summary to power the upload preview UI."""
    try:
        wb = parse_workbook(args.input)  # no api_key → no inspector AI
        inv = build_inventory(wb)
        from socya_pipeline.onboarding import quick_summary
        result = quick_summary(wb, inv)
        sys.stdout.write(json.dumps(result, ensure_ascii=True, default=str))
    except PipelineError as e:
        _emit_error(e)
    except Exception as e:
        _emit_error(PipelineError(ErrorCode.PYTHON_RUNTIME_ERROR,
                                    "Error inesperado en el resumen.",
                                    details=str(e)[:300]))


def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="command", required=True)
    plan_p = sub.add_parser("plan")
    plan_p.add_argument("--input", required=True)
    plan_p.add_argument("--request", default="{}")
    plan_p.set_defaults(func=cmd_plan)
    gen_p = sub.add_parser("generate")
    gen_p.add_argument("--input", required=True)
    gen_p.add_argument("--output", required=True)
    gen_p.add_argument("--template", required=True)
    gen_p.add_argument("--request", default="{}")
    gen_p.set_defaults(func=cmd_generate)
    sum_p = sub.add_parser("quick-summary",
                             help="Deterministic onboarding summary (no AI)")
    sum_p.add_argument("--input", required=True)
    sum_p.set_defaults(func=cmd_quick_summary)
    # PPTX → PNG preview rendering. Best-effort; fails soft if PowerPoint
    # is not available so the caller can still serve the download.
    from socya_pipeline.preview import cmd_preview_pptx, cmd_export_pdf
    prev_p = sub.add_parser("preview-pptx",
                              help="Render PNG previews of each slide in a PPTX")
    prev_p.add_argument("--input", required=True)
    prev_p.add_argument("--output-dir", required=True)
    prev_p.add_argument("--width", type=int, default=None)
    prev_p.set_defaults(func=cmd_preview_pptx)

    pdf_p = sub.add_parser("export-pdf",
                             help="Export a PPTX to PDF via PowerPoint COM (Windows-only)")
    pdf_p.add_argument("--input", required=True)
    pdf_p.add_argument("--output", required=True)
    pdf_p.set_defaults(func=cmd_export_pdf)
    args = p.parse_args()
    args.func(args)

if __name__ == "__main__":
    main()
