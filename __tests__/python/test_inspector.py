"""Tests for the conditional AI inspector.

Critical guarantees:
- Normal Excels trigger ZERO inspections (no extra AI calls).
- Without an API key, the parser works exactly as before (offline-safe).
- Inspector failures NEVER break the pipeline (best-effort only).
"""
import pandas as pd
import pytest

from socya_pipeline.parser import parse_workbook
from socya_pipeline.inspector import (
    needs_inspection, apply_inspector_hints, _parse_response,
)


def _save(tmp_path, name, **sheets):
    p = tmp_path / f"{name}.xlsx"
    with pd.ExcelWriter(p, engine="openpyxl") as w:
        for sheet_name, df in sheets.items():
            df.to_excel(w, sheet_name=sheet_name, index=False)
    return p


# ─────────────────────────────────────────────────────────────────────
# Trigger gate — must be conservative (false negatives > false positives)
# ─────────────────────────────────────────────────────────────────────

class TestTriggerGate:
    def test_normal_excel_does_not_trigger(self, tmp_path):
        """Healthy Excel with currency, dates, categories → ZERO inspection."""
        df = pd.DataFrame({
            "Cliente": [f"Cliente {i % 5}" for i in range(30)],
            "Ciudad":  ["A", "B", "C"] * 10,
            "Fecha":   pd.date_range("2024-01-01", periods=30, freq="D"),
            "Monto":   [(i + 1) * 1000 for i in range(30)],
        })
        p = _save(tmp_path, "normal", Hoja1=df)
        wb = parse_workbook(p)
        assert needs_inspection(wb) == [], \
            "A normal, well-typed Excel must not trigger AI inspection"

    def test_all_text_columns_triggers(self, tmp_path):
        """Sheet with no typed numerics → trigger."""
        df = pd.DataFrame({
            "a": [f"alpha {i}" for i in range(20)],
            "b": [f"beta {i}"  for i in range(20)],
            "c": [f"gamma {i}" for i in range(20)],
        })
        p = _save(tmp_path, "alltext", Hoja1=df)
        wb = parse_workbook(p)
        issues = needs_inspection(wb)
        assert any(r == "no_typed_columns" for _, rs in issues for r in rs), \
            f"Expected no_typed_columns trigger, got {issues}"

    def test_single_column_triggers(self, tmp_path):
        df = pd.DataFrame({"only": [f"item {i}" for i in range(15)]})
        p = _save(tmp_path, "onecol", Hoja1=df)
        wb = parse_workbook(p)
        issues = needs_inspection(wb)
        # Either single_column_with_data OR no_typed_columns → just non-empty
        assert issues, "Single-column sheet should trigger"


# ─────────────────────────────────────────────────────────────────────
# Offline safety — pipeline must work without an API key
# ─────────────────────────────────────────────────────────────────────

class TestOfflineSafety:
    def test_parse_workbook_without_api_key(self, tmp_path):
        """parse_workbook(path) without api_key → identical to before."""
        df = pd.DataFrame({
            "ciudad": ["A", "B", "C"] * 10,
            "monto":  [(i + 1) * 100 for i in range(30)],
        })
        p = _save(tmp_path, "no_key", Hoja1=df)
        wb = parse_workbook(p)  # no api_key
        assert wb.sheets[0].shape == (30, 2)

    def test_parse_workbook_with_empty_api_key(self, tmp_path):
        """Empty string api_key should be treated as no-key (no inspection)."""
        df = pd.DataFrame({
            "x": list(range(10)),
            "y": [f"v{i}" for i in range(10)],
        })
        p = _save(tmp_path, "empty_key", Hoja1=df)
        wb = parse_workbook(p, api_key="")
        assert wb.sheets[0].shape == (10, 2)

    def test_parse_workbook_with_invalid_api_key_does_not_crash(self, tmp_path):
        """Even with a bogus api_key on a problematic sheet, must not raise."""
        # All-text sheet → would trigger inspection if key were real
        df = pd.DataFrame({
            "a": [f"alpha {i}" for i in range(20)],
            "b": [f"beta {i}"  for i in range(20)],
        })
        p = _save(tmp_path, "bad_key", Hoja1=df)
        # Invalid key: AI call will fail, but parser must not crash
        wb = parse_workbook(p, api_key="sk-totally-not-a-real-key")
        assert wb.sheets[0].shape == (20, 2)


# ─────────────────────────────────────────────────────────────────────
# Hints application
# ─────────────────────────────────────────────────────────────────────

class TestHintsApplication:
    def test_dtype_overrides_applied(self, tmp_path):
        df = pd.DataFrame({
            "code": ["X1", "X2", "X3"],
            "n":    [1000, 2000, 3000],
        })
        p = _save(tmp_path, "hints", Hoja1=df)
        wb = parse_workbook(p)
        # Pre-conditions: 'n' might be currency-by-magnitude or numeric
        original_dtype = next(c.dtype for c in wb.sheets[0].columns
                                 if c.name == "n")

        # Fake hint: force 'n' to percent
        hints = {"sheets": {"Hoja1": {"dtype_overrides": {"n": "percent"}}}}
        apply_inspector_hints(wb, hints)
        new_dtype = next(c.dtype for c in wb.sheets[0].columns if c.name == "n")
        assert new_dtype == "percent"
        assert new_dtype != original_dtype

    def test_invalid_hints_silently_ignored(self, tmp_path):
        df = pd.DataFrame({"a": list(range(5))})
        p = _save(tmp_path, "junk_hints", Hoja1=df)
        wb = parse_workbook(p)
        # Various malformed hints — none should crash
        apply_inspector_hints(wb, None)
        apply_inspector_hints(wb, {})
        apply_inspector_hints(wb, "not a dict")
        apply_inspector_hints(wb, {"sheets": "not a dict either"})
        apply_inspector_hints(wb, {"sheets": {"NonExistentSheet": {}}})
        apply_inspector_hints(wb, {"sheets": {"Hoja1": {"dtype_overrides":
                                                            {"a": "bogus_type"}}}})
        # Sheet still intact
        assert wb.sheets[0].shape == (5, 1)


# ─────────────────────────────────────────────────────────────────────
# Response parsing
# ─────────────────────────────────────────────────────────────────────

class TestResponseParsing:
    def test_clean_json(self):
        text = '{"sheets": {"X": {"comment": "ok"}}}'
        assert _parse_response(text) == {"sheets": {"X": {"comment": "ok"}}}

    def test_with_markdown_fences(self):
        text = '```json\n{"sheets": {}}\n```'
        assert _parse_response(text) == {"sheets": {}}

    def test_garbage_returns_empty(self):
        assert _parse_response("not json at all") == {}
        assert _parse_response("") == {}
        assert _parse_response(None) == {}

    def test_array_response_returns_empty(self):
        # If model returns a JSON array instead of object
        assert _parse_response('[1,2,3]') == {}
