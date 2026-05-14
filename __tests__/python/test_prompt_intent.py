"""Tests para socya_pipeline/prompt_intent.py — parser puro del prompt."""
import pytest
from socya_pipeline.prompt_intent import extract, PromptIntent, SheetMatch


# ─── slide count regex ─────────────────────────────────────────────

@pytest.mark.parametrize("prompt,expected", [
    ("hazme 9 slides con riesgos core", 9),
    ("9 slides", 9),
    ("dame 12 diapositivas", 12),
    ("deck de 15", 15),
    ("quiero un deck de 7 slides exactamente", 7),
    ("11 láminas", 11),
    ("Hazme 10 SLIDES", 10),  # case insensitive
])
def test_slide_count_extracted(prompt, expected):
    intent = extract(prompt, available_sheet_names=[])
    assert intent.requested_slide_count == expected


@pytest.mark.parametrize("prompt", [
    "",
    "muéstrame riesgos core sin más",
    "hazme un deck bonito",
    "incluye 3 columnas de datos",  # 3 no es slide count
])
def test_slide_count_none_when_not_specified(prompt):
    intent = extract(prompt, available_sheet_names=[])
    assert intent.requested_slide_count is None


def test_slide_count_first_match_wins():
    intent = extract("hazme 9 slides con 5 columnas", [])
    assert intent.requested_slide_count == 9


def test_slide_count_zero_or_negative_ignored():
    intent = extract("hazme 0 slides", [])
    assert intent.requested_slide_count is None


# ─── required_sheets fuzzy match ──────────────────────────────────

def test_exact_sheet_match_returns_sheetmatch():
    intent = extract("incluye Riesgos Core",
                     available_sheet_names=["Riesgos Core", "Ventas"])
    assert len(intent.required_sheets) == 1
    m = intent.required_sheets[0]
    assert m.requested == "riesgos core"
    assert m.matched == "Riesgos Core"
    assert m.closest is None
    assert m.ratio == 1.0


def test_case_insensitive_match():
    intent = extract("incluye RIESGOS core", ["Riesgos Core"])
    assert intent.required_sheets[0].matched == "Riesgos Core"


def test_accent_insensitive_match():
    intent = extract("incluye operacion", ["Operación"])
    assert intent.required_sheets[0].matched == "Operación"


def test_fuzzy_match_above_threshold():
    intent = extract("incluye riesgos core", ["Riesgos-Core"])
    m = intent.required_sheets[0]
    assert m.matched == "Riesgos-Core"
    assert m.ratio >= 0.65


def test_no_match_returns_closest():
    intent = extract("incluye Proyecciones 2027",
                     ["Riesgos Core", "Acciones", "Ventas"])
    assert len(intent.required_sheets) == 1
    m = intent.required_sheets[0]
    assert m.matched is None
    assert m.closest in {"Riesgos Core", "Acciones", "Ventas"}
    assert m.ratio < 0.65


def test_multiple_sheets_extracted():
    intent = extract("9 slides con Riesgos Core y Riesgos Acciones",
                     ["Riesgos Core", "Riesgos Acciones", "Otra"])
    matched = {m.matched for m in intent.required_sheets}
    assert "Riesgos Core" in matched
    assert "Riesgos Acciones" in matched


def test_no_sheet_mentions_returns_empty():
    intent = extract("hazme 9 slides bonitas",
                     ["Riesgos Core", "Ventas"])
    assert intent.required_sheets == []


def test_empty_sheet_list_returns_empty():
    intent = extract("incluye Riesgos Core", available_sheet_names=[])
    assert intent.required_sheets == []
