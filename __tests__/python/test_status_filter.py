"""Tests for status-filter detection — the fix for the user-reported bug
where 'VALOR TOTAL SOLICITADO $1.50B' was shown as a KPI without revealing
that $258M of that sum came from rejected rows.

Layers exercised:
  - insights.classify_status_value / is_status_filter_column / compute_status_breakdown
  - parser._summarize_sheet populates SheetData.status_breakdowns
  - inventory.build_inventory marks raw KPI with `needs_status_disclaimer`
    and emits sibling `effective_total` / `rejected_total` KPIs
  - validator._bullet_misleads_on_status drops bullets that cite the raw
    sum without mentioning the breakdown
"""
import os
import pandas as pd
import pytest

from socya_pipeline import insights
from socya_pipeline.parser import parse_workbook, _summarize_sheet
from socya_pipeline.inventory import build_inventory
from socya_pipeline.validator import _bullet_misleads_on_status


# ───────────────────── unit: insights ─────────────────────

def test_classify_status_value_negative():
    assert insights.classify_status_value("RECHAZADO") == "negative"
    assert insights.classify_status_value("Rechazado Desembolso") == "negative"
    assert insights.classify_status_value("ANULADO") == "negative"
    assert insights.classify_status_value("CANCELADO") == "negative"


def test_classify_status_value_positive():
    assert insights.classify_status_value("CONTABILIZADO") == "positive"
    assert insights.classify_status_value("Desembolso Realizado") == "positive"
    assert insights.classify_status_value("LEGALIZADO") == "positive"


def test_classify_status_value_pending():
    assert insights.classify_status_value("PENDIENTE") == "pending"
    assert insights.classify_status_value("EN PROCESO") == "pending"
    assert insights.classify_status_value("APROBADO LIDER") == "pending"


def test_classify_status_value_unknown_returns_unknown_never_guesses():
    assert insights.classify_status_value("ABRACADABRA") == "unknown"
    assert insights.classify_status_value(None) == "unknown"
    assert insights.classify_status_value("") == "unknown"


def test_is_status_filter_column_requires_both_name_and_negative_value():
    # Both gates fire → True
    assert insights.is_status_filter_column(
        "Estado", [["RECHAZADO", 10], ["CONTABILIZADO", 5]],
    ) is True
    # Wrong name → False even though values look like a status
    assert insights.is_status_filter_column(
        "Ciudad", [["RECHAZADO", 10]],
    ) is False
    # Right name but no negative value → False (don't trigger disclaimer
    # for a column whose data is all-positive)
    assert insights.is_status_filter_column(
        "Estado", [["CONTABILIZADO", 10], ["PAGADO", 5]],
    ) is False


def test_compute_status_breakdown_buckets_correctly():
    df = pd.DataFrame({
        "Monto": [100, 200, 50, 75, 1000],
        "Estado": ["CONTABILIZADO", "RECHAZADO", "CONTABILIZADO",
                   "PENDIENTE", "DESEMBOLSO REALIZADO"],
    })
    bd = insights.compute_status_breakdown(df, "Monto", "Estado")
    assert bd["positive_sum"] == 1150  # 100 + 50 + 1000
    assert bd["positive_count"] == 3
    assert bd["negative_sum"] == 200
    assert bd["negative_count"] == 1
    assert bd["pending_sum"] == 75
    assert bd["pending_count"] == 1
    assert bd["gross_sum"] == 1425


def test_compute_status_breakdown_handles_missing_columns():
    df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
    # missing currency column
    assert insights.compute_status_breakdown(df, "MISSING", "B") == {}
    # missing status column
    assert insights.compute_status_breakdown(df, "A", "MISSING") == {}


# ───────────────────── unit: parser ─────────────────────

def test_summarize_sheet_populates_status_breakdowns():
    df = pd.DataFrame({
        "ValorTotalSolicitado": [100.0, 200.0, 50.0, 75.0],
        "Estado": ["CONTABILIZADO", "RECHAZADO",
                   "CONTABILIZADO", "PENDIENTE"],
    })
    # Tell the parser these are currency / categorical via a hint
    sheet = _summarize_sheet("Test", df,
                              column_formats={"ValorTotalSolicitado": "$#,##0.00"})
    assert sheet.status_breakdowns, "status_breakdowns should be populated"
    bd = sheet.status_breakdowns.get("ValorTotalSolicitado")
    assert bd is not None
    assert bd["status_column"] == "Estado"
    assert bd["positive_sum"] == 150.0
    assert bd["negative_sum"] == 200.0


def test_summarize_sheet_no_status_breakdown_when_no_status_column():
    df = pd.DataFrame({
        "Monto": [100.0, 200.0],
        "Ciudad": ["Bogota", "Medellin"],
    })
    sheet = _summarize_sheet("Test", df,
                              column_formats={"Monto": "$#,##0.00"})
    assert sheet.status_breakdowns == {}


# ───────────────────── integration: inventory ─────────────────────

@pytest.fixture(scope="module")
def comisiones_wb():
    """Loads the real Comisiones V1.xlsx if present in the repo root.
    Skipped on CI / fresh checkouts where the fixture is not committed."""
    path = "Comisiones V1.xlsx"
    if not os.path.exists(path):
        # fallback to the fixtures version if present
        for alt in ("__tests__/fixtures/comisiones_real.xlsx",
                     "tests/fixtures/comisiones.xlsx"):
            if os.path.exists(alt):
                path = alt
                break
        else:
            pytest.skip("Comisiones fixture not available")
    return parse_workbook(path)


def test_inventory_marks_raw_total_with_disclaimer(comisiones_wb):
    """The K* block that sums 'Valor Total Solicitado' (= $1.50B) must carry
    the `needs_status_disclaimer` flag now."""
    blocks = build_inventory(comisiones_wb)
    raw_total = [
        b for b in blocks
        if b.kind == "kpi_candidate"
        and b.extra.get("agg") == "sum"
        and "Valor Total Solicitado" in (b.provenance.columns[0] if b.provenance.columns else "")
        and b.provenance.sheet == "Comisiones- Base"
    ]
    assert raw_total, "expected a raw-sum KPI for 'Valor Total Solicitado'"
    bruto = raw_total[0]
    assert "needs_status_disclaimer" in bruto.quality_flags
    assert bruto.extra.get("status_filter") is not None
    sf = bruto.extra["status_filter"]
    assert sf["negative_count"] > 0
    # Sanity: gross_sum should match the column's raw sum
    assert abs(sf["gross_sum"] - bruto.extra["value"]) < 1.0


def test_inventory_emits_effective_total_for_filtered_column(comisiones_wb):
    """A sibling KPI block must exist with `effective_total` flag and a
    value close to $1.226B (the positive-state slice of the Comisiones sum)."""
    blocks = build_inventory(comisiones_wb)
    effective = [
        b for b in blocks
        if b.kind == "kpi_candidate"
        and "effective_total" in b.quality_flags
        and "Valor Total Solicitado" in (b.provenance.columns[0] if b.provenance.columns else "")
        and b.provenance.sheet == "Comisiones- Base"
    ]
    assert effective, "expected an effective-total KPI for 'Valor Total Solicitado'"
    eff = effective[0]
    assert eff.extra["state_bucket"] == "positive"
    # ~$1.226B (CONTABILIZADO + DESEMBOLSO REALIZADO + LEGALIZADO)
    assert 1_200_000_000 < eff.extra["value"] < 1_250_000_000


def test_inventory_emits_rejected_kpi_when_material(comisiones_wb):
    """When rejected_sum is ≥5% of gross OR ≥$1M, a rejected_total KPI
    is emitted so the deck can show how much money never moved."""
    blocks = build_inventory(comisiones_wb)
    rejected = [
        b for b in blocks
        if b.kind == "kpi_candidate"
        and "rejected_total" in b.quality_flags
        and "Valor Total Solicitado" in (b.provenance.columns[0] if b.provenance.columns else "")
        and b.provenance.sheet == "Comisiones- Base"
    ]
    assert rejected, "expected a rejected-total KPI for 'Valor Total Solicitado'"
    rej = rejected[0]
    # ~$258M (RECHAZADO + RECHAZADO DESEMBOLSO)
    assert 250_000_000 < rej.extra["value"] < 270_000_000


# ───────────────────── integration: validator ─────────────────────

def test_validator_drops_misleading_total_bullet(comisiones_wb):
    """A bullet that cites the raw $1.5B without any breakdown keyword
    must be flagged as misleading."""
    inv = build_inventory(comisiones_wb)
    misleading = "VALOR TOTAL SOLICITADO 1.50B Promedio por registro: 812.3K"
    assert _bullet_misleads_on_status(misleading, inv) is True


def test_validator_accepts_total_with_disclaimer(comisiones_wb):
    """A bullet that mentions 'bruto' / 'incluye' / 'rechazado' alongside the
    $1.5B is allowed through — the reader has the context they need."""
    inv = build_inventory(comisiones_wb)
    ok_bullets = [
        "El total bruto suma 1.5B, incluyendo 274 rechazados",
        "Por estado, el total solicitado es 1.5B (258M rechazados)",
        "El total efectivo asciende a 1.23B (258M rechazados)",
        "Bogota concentra 53 registros, Medellin 23",
    ]
    for b in ok_bullets:
        assert _bullet_misleads_on_status(b, inv) is False, \
            f"should NOT be flagged: {b!r}"
