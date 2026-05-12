"""Wave 2: smoke test (sin AI) sobre cada fixture.

Para cada fixture, ejecutamos:
  1. parse_workbook
  2. build_inventory
  3. quick_summary (onboarding)

Capturamos: tiempo, conteos, primer error si rompe. Reporte JSON +
human readable. NO ejecutamos planner/render aquí — eso es Wave 3.

Usage: python scripts/stress_test_smoke.py
"""
from __future__ import annotations

import json
import sys
import time
import traceback
from pathlib import Path

# Run from repo root: importar socya_pipeline desde el cwd
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

FIXTURES_DIR = Path("__tests__/stress_fixtures")


def smoke_one(p: Path) -> dict:
    out = {"file": p.name, "size_kb": p.stat().st_size // 1024,
            "stages": {}, "ok": True, "first_error": None}
    try:
        # Stage 1: parser
        t0 = time.perf_counter()
        from socya_pipeline.parser import parse_workbook
        wb = parse_workbook(p, api_key=None)
        dt = (time.perf_counter() - t0) * 1000
        out["stages"]["parse"] = {
            "ms": round(dt, 1),
            "sheets": len(wb.sheets),
            "total_rows": sum(s.shape[0] for s in wb.sheets),
            "total_cols": sum(s.shape[1] for s in wb.sheets),
            "named_ranges": len(wb.named_ranges),
            "units_detected": sum(
                1 for s in wb.sheets for c in s.columns
                if getattr(c, "unit", None)
            ),
        }
    except Exception as e:
        out["ok"] = False
        out["first_error"] = {"stage": "parse",
                                "msg": str(e)[:300],
                                "tb": traceback.format_exc()[-800:]}
        return out

    try:
        # Stage 2: inventory
        t0 = time.perf_counter()
        from socya_pipeline.inventory import build_inventory
        inv = build_inventory(wb)
        dt = (time.perf_counter() - t0) * 1000
        out["stages"]["inventory"] = {
            "ms": round(dt, 1),
            "n_blocks": len(inv),
            "kpis": sum(1 for b in inv if b.kind == "kpi_candidate"),
            "cats": sum(1 for b in inv if b.kind == "categorical_distribution"),
            "ts": sum(1 for b in inv if b.kind == "time_series_candidate"),
            "tables": sum(1 for b in inv if b.kind == "table"),
        }
    except Exception as e:
        out["ok"] = False
        out["first_error"] = {"stage": "inventory",
                                "msg": str(e)[:300],
                                "tb": traceback.format_exc()[-800:]}
        return out

    try:
        # Stage 3: onboarding
        t0 = time.perf_counter()
        from socya_pipeline.onboarding import quick_summary
        summary = quick_summary(wb, inv)
        dt = (time.perf_counter() - t0) * 1000
        out["stages"]["onboarding"] = {
            "ms": round(dt, 1),
            "warnings": len(summary.get("warnings", [])),
            "suggestions": len(summary.get("suggestions", [])),
            "kpis_preview": len(summary.get("kpis_preview", [])),
            "pii_findings": len(summary.get("pii_findings", []) or []),
            "deck_estimate": summary.get("deck_estimate", {}),
        }
    except Exception as e:
        out["ok"] = False
        out["first_error"] = {"stage": "onboarding",
                                "msg": str(e)[:300],
                                "tb": traceback.format_exc()[-800:]}

    return out


def main():
    fixtures = sorted(FIXTURES_DIR.glob("*.xlsx"))
    if not fixtures:
        print("No fixtures found. Run build_stress_fixtures.py first.")
        return

    results = []
    for p in fixtures:
        r = smoke_one(p)
        results.append(r)
        # Reporte de una linea por fixture
        if r["ok"]:
            stages = r["stages"]
            ms_total = sum(s.get("ms", 0) for s in stages.values())
            print(f"[OK]   {r['file']:32s}  total={ms_total:6.1f}ms  "
                   f"sheets={stages.get('parse', {}).get('sheets', 0)}  "
                   f"blocks={stages.get('inventory', {}).get('n_blocks', 0)}  "
                   f"warns={stages.get('onboarding', {}).get('warnings', 0)}")
        else:
            err = r["first_error"]
            print(f"[FAIL] {r['file']:32s}  STAGE={err['stage']}  "
                   f"MSG={err['msg'][:120]}")

    # Detalle por fixture
    print("\n" + "=" * 80)
    print("DETALLE POR FIXTURE")
    print("=" * 80)
    for r in results:
        print(f"\n--- {r['file']} ({r['size_kb']} KB) ---")
        if not r["ok"]:
            print(f"  FAIL en {r['first_error']['stage']}: {r['first_error']['msg']}")
            print(f"  Traceback (final):")
            for line in r["first_error"]["tb"].splitlines()[-5:]:
                print(f"    {line}")
            continue
        for stage, info in r["stages"].items():
            keys = ", ".join(f"{k}={v}" for k, v in info.items() if k != "ms")
            print(f"  {stage:12s}  {info.get('ms', 0):6.1f}ms   {keys}")

    # Resumen global
    n_ok = sum(1 for r in results if r["ok"])
    n_fail = sum(1 for r in results if not r["ok"])
    print(f"\n{'=' * 80}")
    print(f"Total: {n_ok} OK / {n_fail} FAIL  de {len(results)} fixtures.")
    if n_fail > 0:
        print("\nFAILURES:")
        for r in results:
            if not r["ok"]:
                print(f"  - {r['file']}: stage={r['first_error']['stage']} "
                       f"msg={r['first_error']['msg'][:100]}")

    # Persist JSON for later diff
    out_json = Path("__tests__/stress_fixtures/_smoke_report.json")
    out_json.write_text(json.dumps(results, indent=2, ensure_ascii=False, default=str),
                          encoding="utf-8")
    print(f"\nReporte JSON: {out_json}")


if __name__ == "__main__":
    main()
