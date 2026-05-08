from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.validator import validate_plan, ValidationOutcome

def _plan_with(slide_list):
    return {"presentation_meta": {"title": "T", "subtitle": "S"},
            "slides": slide_list,
            "prompt_suggestions": []}

def test_drops_slides_with_unknown_block_ref(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    plan = _plan_with([
        {"type": "chart", "chart_type": "bar", "title": "x",
         "block_ref": "NONEXISTENT", "narrative": "n"},
    ])
    outcome = validate_plan(plan, inv, wb)
    assert outcome.slides == []
    assert outcome.dropped[0]["reason"] == "block_ref_not_found"

def test_keeps_valid_chart(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    cat_block = next(b for b in inv if b.kind == "categorical_distribution")
    plan = _plan_with([
        {"type": "chart", "chart_type": "bar", "title": "x",
         "block_ref": cat_block.id, "narrative": "n"},
    ])
    outcome = validate_plan(plan, inv, wb)
    assert len(outcome.slides) == 1

def test_drops_chart_when_block_flagged_low_quality(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    # Force a single_dominant_category flag on first cat block
    cat_block = next(b for b in inv if b.kind == "categorical_distribution")
    cat_block.quality_flags.add("single_dominant_category")
    plan = _plan_with([
        {"type": "chart", "chart_type": "pie", "title": "x",
         "block_ref": cat_block.id, "narrative": "n"},
    ])
    outcome = validate_plan(plan, inv, wb)
    assert outcome.slides == []
    assert outcome.dropped[0]["reason"] == "single_dominant_category"

def test_returns_at_least_one_slide_or_raises(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    plan = _plan_with([])  # empty plan
    outcome = validate_plan(plan, inv, wb)
    assert outcome.slides == []
    # Caller decides whether to raise EXCEL_INSUFFICIENT_DATA
