from pathlib import Path
import pytest
from pptx import Presentation
from socya_pipeline.renderer import render_pptx

TEMPLATE = Path(__file__).resolve().parent.parent.parent / "Plantilla_Presentacion_Socya (1) (1).pptx"

@pytest.mark.skipif(not TEMPLATE.exists(), reason="template not available")
def test_renders_minimal_deck(tmp_path):
    rendered_slides = [
        {"type": "title", "data": {"title": "Test", "subtitle": "Sub"}},
        {"type": "kpi_row", "data": {"kpis": [{"label": "Total", "value": "12.4M"}]}},
        {"type": "table", "data": {"headers": ["Ciudad", "Total"],
                                     "rows": [["Bogotá", "100"], ["Medellín", "80"]]}},
    ]
    out = tmp_path / "test.pptx"
    render_pptx(rendered_slides,
                presentation_meta={"title": "T", "subtitle": "S"},
                template_path=TEMPLATE, output_path=out)
    assert out.exists() and out.stat().st_size > 1000
    prs = Presentation(out)
    assert len(prs.slides) == 3
