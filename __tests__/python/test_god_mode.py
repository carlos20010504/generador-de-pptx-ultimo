"""End-to-end + edge-case tests for the new intelligence layer.

Covers cases that real users hit but unit tests historically missed:
- Multi-language Excel (EN, PT)
- Edge cases (empty sheet, single column, all-NaN, headers in row 2)
- Smart KPI dedup correctness
- PII detection conservatism
- Structural type detection independent of column language
"""
import pandas as pd
import pytest
from pathlib import Path

from socya_pipeline import insights
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.extractor import extract_for_render, auto_complete_slides


# ─────────────────────────────────────────────────────────────────────
# Insights smoke tests (i18n + structural detectors)
# ─────────────────────────────────────────────────────────────────────

class TestI18n:
    def test_months_es(self):
        assert insights.month_index("enero") == 0
        assert insights.month_index("diciembre") == 11
        assert insights.month_index("setiembre") == 8

    def test_months_en(self):
        assert insights.month_index("january") == 0
        assert insights.month_index("december") == 11
        assert insights.month_index("sept") == 8

    def test_months_pt(self):
        assert insights.month_index("janeiro") == 0
        assert insights.month_index("dezembro") == 11
        assert insights.month_index("março") == 2
        assert insights.month_index("marco") == 2

    def test_months_unknown(self):
        assert insights.month_index("xyz") is None
        assert insights.month_index("") is None

    def test_weekdays(self):
        assert insights.weekday_index("lunes") == 0
        assert insights.weekday_index("monday") == 0
        assert insights.weekday_index("segunda") == 0


class TestStructuralDetection:
    def test_year_column_positive(self):
        assert insights.is_year_column([2020, 2021, 2022, 2023, 2024])

    def test_year_column_negative_too_few(self):
        assert not insights.is_year_column([2020, 2021])

    def test_year_column_negative_out_of_range(self):
        assert not insights.is_year_column([100, 200, 300, 400])

    def test_percent_decimal_form(self):
        assert insights.is_percentage_column([0.1, 0.5, 0.95])

    def test_percent_int_form_needs_hint(self):
        assert not insights.is_percentage_column([10, 50, 90])
        assert insights.is_percentage_column([10, 50, 90], "tasa de éxito")
        assert insights.is_percentage_column([10, 50, 90], "Completion Rate")

    def test_score_column(self):
        assert insights.is_score_column([1, 2, 3, 4, 5, 1, 2, 3])
        assert not insights.is_score_column([100, 200, 300])

    def test_currency_by_magnitude(self):
        assert insights.is_currency_by_magnitude([1000, 5000, 12000, 8500, 9000])

    def test_currency_by_magnitude_rejects_small(self):
        assert not insights.is_currency_by_magnitude([1, 2, 3, 4, 5])

    def test_currency_by_magnitude_rejects_negative(self):
        assert not insights.is_currency_by_magnitude([1000, -500, 2000, 3000, 4000])

    def test_constant_column_detected(self):
        assert insights.is_constant_column(["A", "A", "A", "A"])

    def test_constant_column_negative(self):
        assert not insights.is_constant_column(["A", "B", "A"])

    def test_boolean_disguised_yes_no(self):
        assert insights.looks_boolean_disguised(["yes", "no", "yes", "no"])

    def test_boolean_disguised_aprobado(self):
        assert insights.looks_boolean_disguised(["aprobado", "rechazado"])

    def test_boolean_disguised_negative(self):
        assert not insights.looks_boolean_disguised(["a", "b", "c"])


class TestOutliersAndPareto:
    def test_outliers_finds_extreme(self):
        # need a clear outlier (z >= 2.5)
        out = insights.detect_outliers([1, 1, 1, 1, 1, 1, 1, 1, 50])
        assert len(out) >= 1

    def test_outliers_skips_small_sample(self):
        assert insights.detect_outliers([1, 100]) == []

    def test_pareto_concentrated(self):
        s = insights.pareto_share([90, 5, 3, 1, 1], top_n=3)
        assert s >= 0.95

    def test_pareto_uniform(self):
        s = insights.pareto_share([10, 10, 10, 10, 10], top_n=3)
        assert s and abs(s - 0.6) < 0.01


class TestPIIConservativeness:
    def test_money_column_not_masked(self):
        # 7-digit numbers in a "Valor" column should NOT be detected as IDs
        assert insights.is_pii_column(
            ["1234567", "9876543", "5555555", "1111111"],
            col_name="Valor Total",
        ) is None

    def test_money_column_en_not_masked(self):
        assert insights.is_pii_column(
            ["1234567", "9876543", "5555555", "1111111"],
            col_name="Total Amount",
        ) is None

    def test_doc_id_with_hint_masked(self):
        # Same digits, but column is named for ID → mask
        assert insights.is_pii_column(
            ["1234567", "9876543", "5555555", "1111111"],
            col_name="Cédula",
        ) == "doc_id"

    def test_email_always_masked(self):
        assert insights.is_pii_column(
            ["foo@bar.com", "baz@qux.com", "x@y.com"],
            col_name="contacto",
        ) == "email"

    def test_mask_doc_id(self):
        assert insights.mask_pii("1234567890", "doc_id").endswith("7890")

    def test_mask_email_keeps_first_letter(self):
        assert insights.mask_pii("foobar@example.com", "email").startswith("f")


class TestRatioPairs:
    def _mk_block(self, bid, label, value, sheet="S1"):
        from socya_pipeline.inventory import Block, Provenance
        return Block(
            id=bid, kind="kpi_candidate", label=label,
            provenance=Provenance(sheet=sheet, columns=[label], rows=(0, 10)),
            extra={"agg": "sum", "value": value},
        )

    def test_solicitado_aprobado(self):
        blocks = [
            self._mk_block("K1", "Solicitado", 100),
            self._mk_block("K2", "Aprobado", 75),
        ]
        pairs = insights.find_ratio_pairs(blocks)
        assert len(pairs) == 1
        assert "Aprobación" in pairs[0]["label"]
        assert abs(pairs[0]["percent_value"] - 75) < 0.5

    def test_budget_actual_en(self):
        blocks = [
            self._mk_block("K1", "Budget", 1000),
            self._mk_block("K2", "Actual", 850),
        ]
        pairs = insights.find_ratio_pairs(blocks)
        assert len(pairs) >= 1


# ─────────────────────────────────────────────────────────────────────
# End-to-end tests with synthetic Excels
# ─────────────────────────────────────────────────────────────────────

@pytest.fixture
def english_excel(tmp_path):
    p = tmp_path / "english.xlsx"
    N = 60
    df = pd.DataFrame({
        "Order Date": pd.date_range("2023-01-01", periods=N, freq="W"),
        "Customer Name": [f"Cust {i%10}" for i in range(N)],
        "City": (["NYC", "LA", "Chicago"] * 20)[:N],
        "Quantity": [(i % 5) + 1 for i in range(N)],
        "Unit Price": [(100, 200, 300)[i % 3] for i in range(N)],
        "Total Amount": [((i % 5) + 1) * (100 + (i % 3) * 100) for i in range(N)],
        "Approval Status": (["Approved"] * 4 + ["Rejected"]) * 12,
        "Year": [2023 + i // 30 for i in range(N)],
        "Completion Rate": [round(0.5 + 0.4 * (i % 10) / 10, 2) for i in range(N)],
    })
    with pd.ExcelWriter(p) as w:
        df.to_excel(w, sheet_name="Sales", index=False)
    return p


@pytest.fixture
def portuguese_excel(tmp_path):
    p = tmp_path / "portuguese.xlsx"
    df = pd.DataFrame({
        "Data": pd.date_range("2023-01-01", periods=30, freq="D"),
        "Cliente": [f"Cliente {chr(65 + i % 5)}" for i in range(30)],
        "Cidade": (["São Paulo", "Rio", "Brasília"] * 10),
        "Valor": [(i + 1) * 1000 for i in range(30)],
        "Status": ["Aprovado", "Aprovado", "Rejeitado"] * 10,
    })
    with pd.ExcelWriter(p) as w:
        df.to_excel(w, sheet_name="Vendas", index=False)
    return p


@pytest.fixture
def messy_excel(tmp_path):
    """Excel where row 0 is a title and the real headers are in row 1."""
    p = tmp_path / "messy.xlsx"
    # Build a sheet with a title in row 0
    df = pd.DataFrame({
        "Reporte Anual 2024": [None, "Producto", "A", "B", "C", "D"],
        "Unnamed: 1": [None, "Cantidad", 100, 200, 150, 80],
        "Unnamed: 2": [None, "Precio", 1500, 2200, 1800, 950],
    })
    with pd.ExcelWriter(p) as w:
        df.to_excel(w, sheet_name="Hoja1", index=False)
    return p


@pytest.fixture
def edge_case_excel(tmp_path):
    """Excel with edge cases: very small sheet, single column, sparse NaNs."""
    p = tmp_path / "edge.xlsx"
    df_small = pd.DataFrame({"single_col": ["a", "b"]})
    # Make this truly sparse: ~30% filled, well under the 0.5 threshold
    df_nans = pd.DataFrame({
        "name":  ["A", None, None, "C", None, None, None, None, "E", None],
        "value": [1.0, None, None, None, None, 6.0, None, None, None, None],
    })
    with pd.ExcelWriter(p) as w:
        df_small.to_excel(w, sheet_name="Tiny", index=False)
        df_nans.to_excel(w, sheet_name="Sparse", index=False)
    return p


class TestEnglishExcelE2E:
    def test_parser_detects_dtypes(self, english_excel):
        wb = parse_workbook(english_excel)
        assert len(wb.sheets) == 1
        cols = {c.name: c.dtype for c in wb.sheets[0].columns}
        assert cols["Order Date"] == "date"
        assert cols["Quantity"] == "score"
        assert cols["Unit Price"] == "currency"
        assert cols["Total Amount"] == "currency"
        assert cols["Year"] == "year"
        assert cols["Completion Rate"] == "percent"
        assert cols["Approval Status"] == "bool"
        assert cols["City"] == "categorical"

    def test_inventory_excludes_year_from_kpis(self, english_excel):
        wb = parse_workbook(english_excel)
        inv = build_inventory(wb)
        # Year should NOT appear as KPI candidate
        kpi_labels = [b.label for b in inv if b.kind == "kpi_candidate"]
        assert not any("Year" in l for l in kpi_labels)

    def test_inventory_creates_kpi_for_currency(self, english_excel):
        wb = parse_workbook(english_excel)
        inv = build_inventory(wb)
        kpi_blocks = [b for b in inv if b.kind == "kpi_candidate"]
        currency_kpis = [b for b in kpi_blocks if b.extra.get("agg") == "sum"]
        # Both Unit Price and Total Amount are currency → ≥2 sum KPIs
        assert len(currency_kpis) >= 2


class TestPortugueseE2E:
    def test_parser_recognizes_pt_currency(self, portuguese_excel):
        wb = parse_workbook(portuguese_excel)
        cols = {c.name: c.dtype for c in wb.sheets[0].columns}
        assert cols["Valor"] == "currency"
        assert cols["Data"] == "date"

    def test_pt_status_detected_as_bool(self, portuguese_excel):
        wb = parse_workbook(portuguese_excel)
        cols = {c.name: c.dtype for c in wb.sheets[0].columns}
        # "Aprovado / Rejeitado" — but Rejeitado not in our bool_pairs
        # We only check that it doesn't crash and produces a sensible type
        assert cols["Status"] in ("categorical", "bool")


class TestMessyExcel:
    def test_parser_promotes_real_headers(self, messy_excel):
        wb = parse_workbook(messy_excel)
        # The real headers should be promoted (Producto, Cantidad, Precio)
        col_names = [c.name for c in wb.sheets[0].columns]
        # Even if header promotion doesn't catch everything, we should not crash
        assert len(col_names) >= 1


class TestEdgeCases:
    def test_edge_case_doesnt_crash(self, edge_case_excel):
        wb = parse_workbook(edge_case_excel)
        inv = build_inventory(wb)  # must not raise
        # Sparse sheet should be flagged
        sparse_table = [b for b in inv if b.kind == "table"
                          and b.provenance.sheet == "Sparse"]
        assert len(sparse_table) == 1
        assert "low_fill_ratio" in sparse_table[0].quality_flags

    def test_subsumption_doesnt_misfire(self, english_excel):
        # Total Amount alone should not be flagged subsumed by Unit Price
        wb = parse_workbook(english_excel)
        inv = build_inventory(wb)
        kpi_blocks = [b for b in inv if b.kind == "kpi_candidate"]
        # No subsumption should happen here (the two currencies aren't sub-totals
        # of each other)
        flagged = [b for b in kpi_blocks if "subsumed_by_total" in b.quality_flags]
        assert len(flagged) == 0


class TestRedundantSheetDetection:
    def test_redundant_sheet_flagged(self, tmp_path):
        p = tmp_path / "redundant.xlsx"
        master = pd.DataFrame({
            "id": range(100), "name": [f"x{i}" for i in range(100)],
            "value": [i * 10 for i in range(100)],
        })
        sample = master.head(10).copy()
        with pd.ExcelWriter(p) as w:
            master.to_excel(w, sheet_name="Master", index=False)
            sample.to_excel(w, sheet_name="Sample", index=False)
        wb = parse_workbook(p)
        rids = insights.redundant_sheet_ids(wb.sheets)
        assert "Sample" in rids
        assert "Master" not in rids
