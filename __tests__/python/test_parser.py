import pytest
from pathlib import Path
from socya_pipeline.parser import parse_workbook, WorkbookData
from socya_pipeline.errors import PipelineError, ErrorCode

def test_parse_ventas_simple(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    assert isinstance(wb, WorkbookData)
    assert wb.filename == "ventas_simple.xlsx"
    assert "Ventas" in [s.name for s in wb.sheets]
    ventas = next(s for s in wb.sheets if s.name == "Ventas")
    assert ventas.shape == (100, 5)
    cols = {c.name for c in ventas.columns}
    assert {"Fecha", "Total", "Ciudad", "Vendedor", "Producto"}.issubset(cols)

def test_parse_corrupto_raises(fixtures_dir):
    with pytest.raises(PipelineError) as exc:
        parse_workbook(fixtures_dir / "corrupto.xlsx")
    assert exc.value.code == ErrorCode.EXCEL_INVALID

def test_parse_casi_vacio_succeeds(casi_vacio_path):
    wb = parse_workbook(casi_vacio_path)
    assert wb.sheets
    sheet = wb.sheets[0]
    assert sheet.fill_ratio < 0.5  # mostly empty
