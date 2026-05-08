from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory, Block

def test_inventory_basic(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    ids = {b.id for b in inv}
    assert any(b.kind == "kpi_candidate" for b in inv)
    assert any(b.kind == "categorical_distribution" for b in inv)
    assert any(b.kind == "table" for b in inv)
    # IDs must be unique
    assert len(ids) == len(inv)

def test_inventory_provenance(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    inv = build_inventory(wb)
    for b in inv:
        assert b.provenance.sheet
        assert b.provenance.columns
        assert b.provenance.rows[1] >= b.provenance.rows[0]

def test_inventory_marks_quality_flags(casi_vacio_path):
    wb = parse_workbook(casi_vacio_path)
    inv = build_inventory(wb)
    for b in inv:
        if b.kind == "table":
            assert "low_fill_ratio" in b.quality_flags
