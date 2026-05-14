"""Tests para el enforcement post-AI del planner — garantiza slide_count
y required_sheets sin depender de la AI."""
import pytest
from socya_pipeline.planner import _enforce_intent
from socya_pipeline.prompt_intent import PromptIntent, SheetMatch
from socya_pipeline.parser import WorkbookData, SheetData, ColumnData
from socya_pipeline.inventory import build_inventory


def _make_wb_with_sheets(sheet_names):
    """Construye un WorkbookData mínimo con N hojas sin datos."""
    sheets = []
    for name in sheet_names:
        sheets.append(SheetData(
            name=name,
            shape=(10, 2),
            fill_ratio=0.9,
            columns=[
                ColumnData(name="Categoría", dtype="categorical",
                            n_unique=3, fill_ratio=1.0,
                            samples=["A", "B", "C"],
                            top_values=[("A", 5), ("B", 3), ("C", 2)]),
                ColumnData(name="Monto", dtype="currency",
                            n_unique=10, fill_ratio=1.0,
                            samples=[100.0, 200.0, 300.0],
                            sum=2000.0, mean=200.0, min=50.0, max=500.0),
            ],
            first_rows=[["A", 100], ["B", 200]],
        ))
    return WorkbookData(filename="test.xlsx", sheets=sheets)


# ─── Slide count enforcement ──────────────────────────────────────

def test_enforce_trim_to_exact_count():
    wb = _make_wb_with_sheets(["Hoja1"])
    blocks = build_inventory(wb)
    plan = {
        "slides": [
            {"type": "title", "data": {"title": "T"}},
            {"type": "kpi_row", "block_refs": [b.id for b in blocks if b.kind == "kpi_candidate"][:2]},
            {"type": "chart", "block_ref": next(b.id for b in blocks if b.kind == "categorical_distribution"),
             "provenance": {"sheet": "Hoja1"}},
            {"type": "chart", "block_ref": next(b.id for b in blocks if b.kind == "categorical_distribution"),
             "provenance": {"sheet": "Hoja1"}},
            {"type": "table", "block_ref": next(b.id for b in blocks if b.kind == "table"),
             "provenance": {"sheet": "Hoja1"}},
        ]
    }
    intent = PromptIntent(requested_slide_count=3)
    result = _enforce_intent(plan, intent, blocks, wb)
    assert len(result["slides"]) == 3
    assert result["slides"][0]["type"] == "title"
    assert result["_intent_report"]["actual_slide_count"] == 3
    assert result["_intent_report"]["count_honored"] is True


def test_enforce_pad_to_count():
    wb = _make_wb_with_sheets(["Hoja1", "Hoja2"])
    blocks = build_inventory(wb)
    plan = {
        "slides": [
            {"type": "title", "data": {"title": "T"}},
            {"type": "kpi_row", "block_refs": [b.id for b in blocks if b.kind == "kpi_candidate"][:1]},
        ]
    }
    intent = PromptIntent(requested_slide_count=5)
    result = _enforce_intent(plan, intent, blocks, wb)
    assert len(result["slides"]) == 5
    assert result["_intent_report"]["count_honored"] is True


def test_enforce_pad_caps_at_available():
    wb = _make_wb_with_sheets(["Hoja1"])
    blocks = build_inventory(wb)
    plan = {"slides": [{"type": "title", "data": {"title": "T"}}]}
    intent = PromptIntent(requested_slide_count=100)
    result = _enforce_intent(plan, intent, blocks, wb)
    assert len(result["slides"]) < 100
    assert result["_intent_report"]["count_honored"] is False


def test_enforce_count_none_is_noop():
    wb = _make_wb_with_sheets(["Hoja1"])
    blocks = build_inventory(wb)
    plan = {"slides": [{"type": "title", "data": {"title": "T"}}]}
    intent = PromptIntent(requested_slide_count=None)
    result = _enforce_intent(plan, intent, blocks, wb)
    assert len(result["slides"]) == 1
    assert result["_intent_report"]["count_honored"] is True


# ─── Required sheets enforcement ──────────────────────────────────

def test_enforce_inject_missing_required_sheet():
    wb = _make_wb_with_sheets(["Riesgos Core", "Otra"])
    blocks = build_inventory(wb)
    plan = {
        "slides": [
            {"type": "title", "data": {"title": "T"}},
            {"type": "chart",
             "block_ref": next(b.id for b in blocks if b.kind == "categorical_distribution" and b.provenance.sheet == "Otra"),
             "provenance": {"sheet": "Otra"}},
        ]
    }
    intent = PromptIntent(
        required_sheets=[SheetMatch(requested="riesgos core",
                                       matched="Riesgos Core",
                                       closest=None, ratio=1.0)],
    )
    result = _enforce_intent(plan, intent, blocks, wb)
    sheets_in_plan = {s.get("provenance", {}).get("sheet")
                      for s in result["slides"] if "provenance" in s}
    assert "Riesgos Core" in sheets_in_plan
    sheet_report = result["_intent_report"]["required_sheets"][0]
    assert sheet_report["matched"] == "Riesgos Core"
    assert len(sheet_report["slide_indices"]) >= 1


def test_enforce_already_covered_sheet_no_op():
    wb = _make_wb_with_sheets(["Riesgos Core"])
    blocks = build_inventory(wb)
    plan = {
        "slides": [
            {"type": "title", "data": {"title": "T"}},
            {"type": "chart",
             "block_ref": next(b.id for b in blocks if b.kind == "categorical_distribution"),
             "provenance": {"sheet": "Riesgos Core"}},
        ]
    }
    intent = PromptIntent(
        required_sheets=[SheetMatch(requested="riesgos core",
                                       matched="Riesgos Core",
                                       closest=None, ratio=1.0)],
    )
    result = _enforce_intent(plan, intent, blocks, wb)
    assert sum(1 for s in result["slides"]
                if s.get("provenance", {}).get("sheet") == "Riesgos Core") == 1


def test_enforce_intent_report_has_unmatched_sheet():
    wb = _make_wb_with_sheets(["Otra"])
    blocks = build_inventory(wb)
    plan = {"slides": [{"type": "title", "data": {"title": "T"}}]}
    intent = PromptIntent(
        required_sheets=[SheetMatch(requested="proyecciones",
                                       matched=None,
                                       closest="Otra", ratio=0.5)],
    )
    result = _enforce_intent(plan, intent, blocks, wb)
    sheet_report = result["_intent_report"]["required_sheets"][0]
    assert sheet_report["matched"] is None
    assert sheet_report["closest"] == "Otra"
    assert sheet_report["slide_indices"] == []


def test_enforce_idempotent():
    """Aplicar enforce dos veces produce el mismo resultado."""
    wb = _make_wb_with_sheets(["Riesgos Core"])
    blocks = build_inventory(wb)
    plan = {"slides": [{"type": "title", "data": {"title": "T"}}]}
    intent = PromptIntent(
        requested_slide_count=3,
        required_sheets=[SheetMatch(requested="riesgos core",
                                       matched="Riesgos Core",
                                       closest=None, ratio=1.0)],
    )
    once = _enforce_intent(plan, intent, blocks, wb)
    twice = _enforce_intent(once, intent, blocks, wb)
    assert len(twice["slides"]) == len(once["slides"])
    assert twice["_intent_report"] == once["_intent_report"]
