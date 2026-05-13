"""Tests for the deterministic onboarding summary."""
import pandas as pd
import pytest

from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.onboarding import quick_summary


def _save(tmp_path, name, **sheets):
    p = tmp_path / f"{name}.xlsx"
    with pd.ExcelWriter(p, engine="openpyxl") as w:
        for sheet_name, df in sheets.items():
            df.to_excel(w, sheet_name=sheet_name, index=False)
    return p


@pytest.fixture
def commerce_excel(tmp_path):
    N = 40
    cities = ["Bogotá", "Medellín", "Cali"]
    df = pd.DataFrame({
        "Cliente":    [f"Cliente {chr(65 + i % 5)}" for i in range(N)],
        "Ciudad":     [cities[i % 3] for i in range(N)],
        "Fecha":      pd.date_range("2024-01-01", periods=N, freq="W"),
        "Solicitado": [(i + 1) * 10000 for i in range(N)],
        "Aprobado":   [(i + 1) * 8500 for i in range(N)],
        "Estado":     (["Aprobado"] * 4 + ["Rechazado"]) * (N // 5),
    })
    return _save(tmp_path, "commerce", Operaciones=df)


class TestQuickSummary:
    def test_zero_ai_calls_required(self, commerce_excel):
        """Must work without any api_key — pure deterministic."""
        wb = parse_workbook(commerce_excel)  # no api_key → no inspector
        inv = build_inventory(wb)
        summary = quick_summary(wb, inv)
        # Just verify it produced a sensible result
        assert summary["totals"]["sheets"] == 1
        assert summary["totals"]["rows"] == 40
        assert summary["totals"]["kpi_candidates"] >= 1

    def test_top_kpis_have_compact_value(self, commerce_excel):
        wb = parse_workbook(commerce_excel)
        inv = build_inventory(wb)
        summary = quick_summary(wb, inv)
        assert summary["kpis_preview"]
        for k in summary["kpis_preview"]:
            assert isinstance(k["value"], str)
            assert any(s in k["value"] for s in ("$", "K", "M", "B", "%"))

    def test_warnings_for_missing_dates(self, tmp_path):
        df = pd.DataFrame({
            "ciudad": ["A", "B", "C"] * 10,
            "monto":  [(i + 1) * 100 for i in range(30)],
        })
        p = _save(tmp_path, "no_dates", Hoja1=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        summary = quick_summary(wb, inv)
        # Should warn about missing dates
        icons = [w["icon"] for w in summary["warnings"]]
        assert "missing_dates" in icons

    def test_warnings_for_low_fill(self, tmp_path):
        df = pd.DataFrame({
            "a": ["x", None, None, None, "y", None, None, None],
            "b": [1, None, None, None, None, 2, None, None],
        })
        p = _save(tmp_path, "sparse", Hoja1=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        summary = quick_summary(wb, inv)
        icons = [w["icon"] for w in summary["warnings"]]
        assert "low_fill" in icons

    def test_suggestions_match_data_shape(self, commerce_excel):
        wb = parse_workbook(commerce_excel)
        inv = build_inventory(wb)
        summary = quick_summary(wb, inv)
        # Has Cliente+Ciudad+Solicitado/Aprobado+Fecha+Estado → suggestions cover
        # all the patterns
        ids = [s["id"] for s in summary["suggestions"]]
        assert "geo_breakdown" in ids
        assert "trend_analysis" in ids
        assert "status_funnel" in ids
        assert "executive_summary" in ids

    def test_deck_estimate_is_sane(self, commerce_excel):
        wb = parse_workbook(commerce_excel)
        inv = build_inventory(wb)
        summary = quick_summary(wb, inv)
        de = summary["deck_estimate"]
        assert de["min_slides"] == 7
        assert 7 <= de["max_slides"] <= 11
        assert "Portada" in de["expected_sections"]
        assert "Hallazgos" in de["expected_sections"]

    def test_redundant_sheets_flagged(self, tmp_path):
        master = pd.DataFrame({"id": range(50), "v": [i * 10 for i in range(50)]})
        sample = master.head(10).copy()
        p = tmp_path / "dup.xlsx"
        with pd.ExcelWriter(p, engine="openpyxl") as w:
            master.to_excel(w, sheet_name="Master", index=False)
            sample.to_excel(w, sheet_name="Sample", index=False)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        summary = quick_summary(wb, inv)
        sample_meta = next(s for s in summary["sheets"] if s["name"] == "Sample")
        assert sample_meta["is_redundant"] is True

    def test_first_cols_humanized(self, tmp_path):
        df = pd.DataFrame({
            "valor_total":     [100, 200],
            "fecha_creacion":  ["2024-01-01", "2024-01-02"],
            "estado_proceso":  ["A", "B"],
        })
        p = _save(tmp_path, "snake", Hoja1=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        summary = quick_summary(wb, inv)
        first_cols = summary["sheets"][0]["first_cols"]
        # snake_case → "Snake Case"
        assert any(" " in c for c in first_cols)
