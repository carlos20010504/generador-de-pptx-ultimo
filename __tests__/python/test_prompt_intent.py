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
