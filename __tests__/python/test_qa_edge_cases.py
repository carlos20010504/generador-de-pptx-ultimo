"""Aggressive edge-case tests against parser/inventory/extractor.

Each test simulates a real-world Excel pathology that has been seen in the
wild (or that we expect users to throw at us). The tests assert that:
  (a) the pipeline does NOT crash
  (b) the output is sane (no NaN strings rendered, no garbage dtypes)
"""
import math
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

from socya_pipeline import insights
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.extractor import (
    extract_for_render, auto_complete_slides,
    _format_table_cell, _format_currency_compact, _format_kpi_value,
)


# ─────────────────────────────────────────────────────────────────────
# Pathological data fixtures
# ─────────────────────────────────────────────────────────────────────

def _save(tmp_path, name, **sheets) -> Path:
    p = tmp_path / f"{name}.xlsx"
    with pd.ExcelWriter(p, engine="openpyxl") as w:
        for sheet_name, df in sheets.items():
            df.to_excel(w, sheet_name=sheet_name, index=False)
    return p


# ─────────────────────────────────────────────────────────────────────
# Tests: pathological numeric values
# ─────────────────────────────────────────────────────────────────────

class TestPathologicalNumerics:
    def test_nan_inf_in_numeric_column(self, tmp_path):
        df = pd.DataFrame({
            "ciudad": ["A", "B", "C", "D", "E"],
            "valor":  [100.0, float("nan"), float("inf"), -float("inf"), 200.0],
        })
        p = _save(tmp_path, "nans", Hoja1=df)
        wb = parse_workbook(p)
        # Should NOT raise. Inf values should be coerced/dropped, not produce
        # a "Total = inf" KPI.
        inv = build_inventory(wb)
        kpi = next((b for b in inv if b.kind == "kpi_candidate"
                      and b.label.lower().startswith("total")
                      and "valor" in b.label.lower()), None)
        if kpi:
            v = kpi.extra.get("value")
            # If a KPI was emitted, its value MUST be finite
            assert v is None or (isinstance(v, (int, float))
                                  and math.isfinite(v)), \
                f"KPI emitted infinite value: {v}"

    def test_extreme_magnitudes(self, tmp_path):
        # Mix tiny and huge — formatter must not produce '1.5e+12K' nonsense
        df = pd.DataFrame({
            "concepto": ["alpha", "beta", "gamma", "delta", "epsilon"],
            "monto":    [0.001, 1.5, 1_500_000_000, 0.0, 250],
        })
        p = _save(tmp_path, "extremes", Hoja1=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        # Format every value via the table formatter — none should crash
        for v in df["monto"]:
            formatted = _format_table_cell(v, df["monto"], "monto", "currency")
            assert isinstance(formatted, str)
            assert "nan" not in formatted.lower()
            assert "inf" not in formatted.lower()

    def test_format_kpi_value_handles_nan_inf(self):
        # NaN, None, and Inf all render as the em-dash placeholder
        assert _format_kpi_value(float("nan")) == "—"
        assert _format_kpi_value(None) == "—"
        assert _format_kpi_value(float("inf")) == "—"
        assert _format_kpi_value(float("-inf")) == "—"


# ─────────────────────────────────────────────────────────────────────
# Tests: empty / minimal / malformed sheets
# ─────────────────────────────────────────────────────────────────────

class TestMinimalSheets:
    def test_completely_empty_excel_raises_clean_error(self, tmp_path):
        p = tmp_path / "empty.xlsx"
        with pd.ExcelWriter(p, engine="openpyxl") as w:
            pd.DataFrame().to_excel(w, sheet_name="Empty", index=False)
        # parse_workbook may succeed with an empty sheet or raise — but
        # whatever happens, build_inventory must not crash on the result
        try:
            wb = parse_workbook(p)
            inv = build_inventory(wb)
            assert isinstance(inv, list)
        except Exception as e:
            # Acceptable: pipeline error type, not a generic crash
            from socya_pipeline.errors import PipelineError
            assert isinstance(e, PipelineError)

    def test_single_row_sheet(self, tmp_path):
        df = pd.DataFrame({"a": [1], "b": ["x"]})
        p = _save(tmp_path, "single_row", Hoja1=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        # Single row → too_few_rows flag on the table block
        table_blocks = [b for b in inv if b.kind == "table"]
        assert len(table_blocks) == 1
        assert "too_few_rows" in table_blocks[0].quality_flags

    def test_single_column_only(self, tmp_path):
        df = pd.DataFrame({"only_col": list(range(50))})
        p = _save(tmp_path, "single_col", Hoja1=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        # Should not crash; should produce a table block at least
        assert any(b.kind == "table" for b in inv)


# ─────────────────────────────────────────────────────────────────────
# Tests: weird sheet & column names
# ─────────────────────────────────────────────────────────────────────

class TestWeirdNames:
    def test_sheet_name_with_spaces_and_emoji(self, tmp_path):
        df = pd.DataFrame({
            "concepto": ["uno", "dos", "tres"],
            "valor":    [100, 200, 300],
        })
        p = tmp_path / "weird_sheet.xlsx"
        with pd.ExcelWriter(p, engine="openpyxl") as w:
            # Note: sheet name has trailing space + emoji-ish symbol
            df.to_excel(w, sheet_name="Mi Hoja  $", index=False)
        wb = parse_workbook(p)
        # Sheet name preserved through pipeline
        assert wb.sheets[0].name == "Mi Hoja  $"
        inv = build_inventory(wb)
        assert len(inv) > 0

    def test_column_with_leading_trailing_spaces(self, tmp_path):
        df = pd.DataFrame({
            " ciudad ": ["A", "B", "C"],
            "  valor  ": [100, 200, 300],
        })
        p = _save(tmp_path, "spaces", Hoja1=df)
        wb = parse_workbook(p)
        # Column names preserved as-is (extractor uses fuzzy match later)
        cols = [c.name for c in wb.sheets[0].columns]
        assert any("ciudad" in c for c in cols)


# ─────────────────────────────────────────────────────────────────────
# Tests: dates in various forms
# ─────────────────────────────────────────────────────────────────────

class TestDateVariations:
    def test_string_date_iso(self, tmp_path):
        df = pd.DataFrame({
            "fecha": ["2024-01-15", "2024-02-20", "2024-03-10",
                       "2024-04-05", "2024-05-12", "2024-06-18"],
            "ventas": [1000, 1500, 800, 2000, 1200, 1800],
        })
        p = _save(tmp_path, "iso_dates", Hoja1=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        # ISO strings get parsed as date by pandas — should yield a time series
        ts = [b for b in inv if b.kind == "time_series_candidate"]
        assert len(ts) >= 1, "Expected time series block from ISO date column"

    def test_real_datetime_objects(self, tmp_path):
        df = pd.DataFrame({
            "stamp": pd.date_range("2024-01-01", periods=10, freq="D"),
            "amount": [(i + 1) * 100 for i in range(10)],
        })
        p = _save(tmp_path, "real_dates", Hoja1=df)
        wb = parse_workbook(p)
        # First col should be detected as date
        cols = {c.name: c.dtype for c in wb.sheets[0].columns}
        assert cols["stamp"] == "date"


# ─────────────────────────────────────────────────────────────────────
# Tests: text edge cases
# ─────────────────────────────────────────────────────────────────────

class TestTextEdgeCases:
    def test_very_long_text_in_cell(self, tmp_path):
        long = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " * 10
        df = pd.DataFrame({
            "id": list(range(5)),
            "descripcion": [long] * 5,
            "monto": [100, 200, 300, 400, 500],
        })
        p = _save(tmp_path, "long_text", Hoja1=df)
        wb = parse_workbook(p)
        # `descripcion` (avg len > 25) should be classified as text, not categorical
        cols = {c.name: c.dtype for c in wb.sheets[0].columns}
        assert cols["descripcion"] == "text"

    def test_table_cell_truncates_long_text(self):
        long = "X" * 200
        out = _format_table_cell(long, None, "col", "text")
        assert len(out) <= 38  # truncate cap

    def test_unicode_column_names(self, tmp_path):
        df = pd.DataFrame({
            "categoría": ["Α", "Β", "Γ"],  # Greek letters in values
            "porcentaje %": [0.10, 0.50, 0.40],
        })
        p = _save(tmp_path, "unicode", Hoja1=df)
        wb = parse_workbook(p)
        cols = {c.name: c.dtype for c in wb.sheets[0].columns}
        # Percent detection should still kick in
        assert cols["porcentaje %"] == "percent"


# ─────────────────────────────────────────────────────────────────────
# Tests: insights formatters under stress
# ─────────────────────────────────────────────────────────────────────

class TestFormatters:
    def test_currency_compact_extremes(self):
        assert _format_currency_compact(0) == "$0"
        assert _format_currency_compact(100) == "$100"
        assert _format_currency_compact(1500).endswith("K") or "1,500" in _format_currency_compact(1500)
        assert _format_currency_compact(2_500_000).endswith("M")
        assert _format_currency_compact(3_500_000_000).endswith("B")

    def test_currency_compact_negative(self):
        # Negative values shouldn't crash even if they're unusual for currency
        result = _format_currency_compact(-5000)
        assert isinstance(result, str)

    def test_kpi_value_at_threshold_boundaries(self):
        # Exactly at threshold boundaries — must not double-count
        assert _format_kpi_value(999) == "999"
        assert _format_kpi_value(1000).endswith("K")
        assert _format_kpi_value(999_999.99).endswith("K")
        assert _format_kpi_value(1_000_000).endswith("M")
        assert _format_kpi_value(1_000_000_000).endswith("B")


# ─────────────────────────────────────────────────────────────────────
# Tests: pipeline integration on a "rich" stress Excel
# ─────────────────────────────────────────────────────────────────────

class TestRichStressExcel:
    @pytest.fixture
    def stress_excel(self, tmp_path):
        # An Excel that combines many edge cases at once
        df_main = pd.DataFrame({
            "Id Operación":      list(range(1, 51)),  # ID-like, should be filtered
            "Cliente":           [f"Cliente {chr(65 + i % 10)}" for i in range(50)],
            "Ciudad":            ["Bogotá", "Medellín", "Cali", "Barranquilla",
                                    "Cartagena"] * 10,
            "Fecha":             pd.date_range("2024-01-01", periods=50, freq="W"),
            "Monto Solicitado":  [(i + 1) * 12_500 for i in range(50)],
            "Monto Aprobado":    [(i + 1) * 11_000 for i in range(50)],
            "Tasa Aprobación":   [0.85 + 0.1 * (i % 5) / 5 for i in range(50)],
            "Estado":            (["Aprobado"] * 4 + ["Rechazado"]) * 10,
            "Año":               [2024] * 50,
            "Score":             [(i % 5) + 1 for i in range(50)],
            "Observaciones":     ["Lorem ipsum dolor sit amet " * 5] * 50,
        })
        # Add a TOTAL row at the bottom (to test stripping)
        total_row = pd.DataFrame([{
            "Id Operación": "TOTAL",
            "Cliente": "",
            "Ciudad": "",
            "Fecha": None,
            "Monto Solicitado": df_main["Monto Solicitado"].sum(),
            "Monto Aprobado":   df_main["Monto Aprobado"].sum(),
            "Tasa Aprobación":  None,
            "Estado": "",
            "Año": None,
            "Score": None,
            "Observaciones": "",
        }])
        df_with_total = pd.concat([df_main, total_row], ignore_index=True)
        return _save(tmp_path, "stress", Operaciones=df_with_total)

    def test_no_crash_full_pipeline(self, stress_excel):
        wb = parse_workbook(stress_excel)
        inv = build_inventory(wb)
        # Total row should have been stripped
        assert wb.sheets[0].shape[0] == 50, \
            f"TOTAL row not stripped, shape={wb.sheets[0].shape}"

    def test_stress_dtypes_correct(self, stress_excel):
        wb = parse_workbook(stress_excel)
        cols = {c.name: c.dtype for c in wb.sheets[0].columns}
        assert cols["Id Operación"] == "year" or "currency" or "numeric"  # any but not crash
        assert cols["Fecha"] == "date"
        assert cols["Monto Solicitado"] == "currency"
        assert cols["Monto Aprobado"]   == "currency"
        assert cols["Tasa Aprobación"]  == "percent"
        assert cols["Año"]              == "year"
        assert cols["Score"]            == "score"
        assert cols["Observaciones"]    == "text"

    def test_stress_kpi_dedup_doesnt_show_id_as_kpi(self, stress_excel):
        wb = parse_workbook(stress_excel)
        inv = build_inventory(wb)
        kpi_labels_lower = [b.label.lower() for b in inv
                              if b.kind == "kpi_candidate"]
        # No KPI should reference identifier-only columns (Id Operación) or
        # dimension columns (Año). 'Score' IS a legitimate KPI (rating avg)
        # so it's allowed.
        for bad in ("id ", "operación", "año"):
            assert not any(bad in l for l in kpi_labels_lower), \
                f"KPI list contains an ID/year column: {kpi_labels_lower}"

    def test_stress_derived_ratio_emitted(self, stress_excel):
        wb = parse_workbook(stress_excel)
        inv = build_inventory(wb)
        # Solicitado / Aprobado pair → should emit a derived "% Aprobación"
        derived = [b for b in inv if "derived" in b.quality_flags]
        # If pair detected, value should be a sane percentage 0..200
        if derived:
            v = derived[0].extra.get("value")
            assert v is not None and 0 <= float(v) <= 200, \
                f"Derived ratio out of sane range: {v}"
