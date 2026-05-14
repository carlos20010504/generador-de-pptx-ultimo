import os
import pytest
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent.parent / "fixtures"


# Aislar tests de cualquier API key real que tenga el usuario en el shell o
# .env. Sin esto, tests que esperan "openrouter/..." se rompen cuando hay
# CEREBRAS_API_KEY/GROQ_API_KEY en env (Cerebras se prueba primero).
@pytest.fixture(autouse=True)
def _isolate_provider_env(monkeypatch):
    for var in ("CEREBRAS_API_KEY", "GROQ_API_KEY"):
        monkeypatch.delenv(var, raising=False)
    yield

@pytest.fixture
def fixtures_dir():
    return FIXTURES_DIR

@pytest.fixture
def ventas_simple_path():
    return FIXTURES_DIR / "ventas_simple.xlsx"

@pytest.fixture
def casi_vacio_path():
    return FIXTURES_DIR / "casi_vacio.xlsx"

@pytest.fixture
def stub_planner_response():
    return {
        "presentation_meta": {"title": "Ventas 2024", "subtitle": "Resumen"},
        "slides": [
            {"type": "title", "title": "Ventas 2024", "subtitle": "Resumen"},
            {"type": "kpi_row", "title": "KPIs principales", "block_refs": ["K1"]},
            {"type": "chart", "chart_type": "bar", "title": "Por ciudad",
             "block_ref": "C1", "narrative": "Bogotá lidera con 60 ventas."},
        ],
        "prompt_suggestions": ["Compara Bogotá vs Medellín"],
    }
