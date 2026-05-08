"""CLI entry: `python -m socya_pipeline {analyze|generate} <args>`."""
import argparse
import json
import sys
import os
from pathlib import Path
from socya_pipeline.errors import PipelineError, ErrorCode
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.planner import plan_presentation
from socya_pipeline.validator import validate_plan
from socya_pipeline.extractor import extract_for_render
from socya_pipeline.ai_chain import AIProfile

def _emit_error(err: PipelineError):
    sys.stdout.write(json.dumps({"error": err.to_dict()}, ensure_ascii=False))
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
    try:
        wb = parse_workbook(args.input)
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
        rendered = extract_for_render(outcome.slides, inv, wb, args.input)
        result = {
            "presentation_meta": plan.get("presentation_meta", {}),
            "slides": rendered,
            "prompt_suggestions": plan.get("prompt_suggestions", []),
            "ai_status": plan.get("_meta", {}),
            "audit": {
                "slides_planned": len(plan.get("slides", [])),
                "slides_validated": len(outcome.slides),
                "slides_dropped": outcome.dropped,
                "bullets_dropped": outcome.bullets_dropped,
            },
        }
        sys.stdout.write(json.dumps(result, ensure_ascii=False, default=str))
    except PipelineError as e:
        _emit_error(e)
    except Exception as e:
        _emit_error(PipelineError(ErrorCode.PYTHON_RUNTIME_ERROR,
                                    "Error inesperado en el planificador.",
                                    details=str(e)[:300]))

def _resolve_api_key() -> str:
    key = os.environ.get("OPENROUTER_API_KEY", "").strip()
    if key:
        return key
    for env_file in (".env", ".env.local"):
        if os.path.exists(env_file):
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("OPENROUTER_API_KEY="):
                        return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""

def main():
    p = argparse.ArgumentParser()
    sub = p.add_subparsers(dest="command", required=True)
    plan_p = sub.add_parser("plan")
    plan_p.add_argument("--input", required=True)
    plan_p.add_argument("--request", default="{}")
    plan_p.set_defaults(func=cmd_plan)
    args = p.parse_args()
    args.func(args)

if __name__ == "__main__":
    main()
