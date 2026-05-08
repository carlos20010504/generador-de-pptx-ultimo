from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.validator import validate_plan
from socya_pipeline.extractor import extract_for_render

def test_extracts_table_with_clean_cells(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    table_block = next(b for b in inv if b.kind == "table")
    plan = {"presentation_meta": {"title":"T","subtitle":"S"},
            "slides": [{"type": "table", "title": "x",
                         "block_ref": table_block.id,
                         "columns_subset": [c for c in table_block.provenance.columns[:3]],
                         "max_rows": 5}]}
    outcome = validate_plan(plan, inv, wb)
    rendered = extract_for_render(outcome.slides, inv, wb,
                                    file_path=ventas_simple_path)
    assert len(rendered) == 1
    table_slide = rendered[0]
    assert "rows" in table_slide["data"]
    assert len(table_slide["data"]["rows"]) <= 5
    # No ugly literals
    flat = [str(c) for row in table_slide["data"]["rows"] for c in row]
    for forbidden in ("nan", "None", "NaT", "null"):
        assert forbidden not in flat

def test_kpi_extraction(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    kpi_block = next(b for b in inv if b.kind == "kpi_candidate")
    plan = {"presentation_meta":{"title":"T","subtitle":"S"},
            "slides": [{"type": "kpi_row", "title": "k",
                         "block_refs": [kpi_block.id]}]}
    outcome = validate_plan(plan, inv, wb)
    rendered = extract_for_render(outcome.slides, inv, wb,
                                    file_path=ventas_simple_path)
    assert rendered[0]["data"]["kpis"][0]["label"]
    assert rendered[0]["data"]["kpis"][0]["value"]
