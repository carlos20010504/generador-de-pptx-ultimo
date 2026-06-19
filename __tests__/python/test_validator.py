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


def test_fuzzy_resolve_recovers_case_mismatch(ventas_simple_path):
    """El AI a veces manda block_ref con case wrong. Antes se dropeaba.
    Ahora se recupera al ID real sin perder la slide."""
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    cat_block = next(b for b in inv if b.kind == "categorical_distribution")
    # Mandamos el ID en lowercase — antes hubiera fallado
    wrong_case = cat_block.id.lower()
    if wrong_case == cat_block.id:
        # Si el id ya es lowercase, fuerza un upper para que el test tenga sentido
        wrong_case = cat_block.id.upper()
    plan = _plan_with([
        {"type": "chart", "chart_type": "bar", "title": "x",
         "block_ref": wrong_case, "narrative": "n"},
    ])
    outcome = validate_plan(plan, inv, wb)
    assert len(outcome.slides) == 1, f"Expected slide recovered, got dropped: {outcome.dropped}"
    # El slide debe tener el ID canónico, no el wrong_case
    assert outcome.slides[0]["block_ref"] == cat_block.id


def test_fuzzy_does_not_match_unrelated_blocks(ventas_simple_path):
    """Garantía: NO confundimos K1 con K11 ni IDs disímiles. Sin esto, el
    fuzzy resolver podría recuperar refs falsos que entreguen data wrong."""
    from socya_pipeline.validator import _fuzzy_resolve_block_ref
    valid = ["K1", "K11", "K12", "T1", "C1"]
    # K1 → debe matchear K1 exacto, no K11 ni K12
    assert _fuzzy_resolve_block_ref("K1", valid) == "K1"
    # ID completamente distinto → None
    assert _fuzzy_resolve_block_ref("XYZ99", valid) is None
    # case mismatch → recupera
    assert _fuzzy_resolve_block_ref("k1", valid) == "K1"
    # ID con separator typo → recupera
    assert _fuzzy_resolve_block_ref("K_1", valid) == "K1"
