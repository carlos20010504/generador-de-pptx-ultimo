"""Regression tests for stress-test bugs (2026-05-26).

Two bugs are guarded here:

1. **dayfirst date parsing (CRITICAL data-incorrect):** time series charts
   showed ventas dispersed in invented months because pandas parsed
   `DD/MM/YYYY` as `MM/DD/YYYY` by default.

2. **silent slide drop (defensive):** `extract_for_render` had no `else`
   branch for unknown slide types — if the AI/validator let through a
   slide with `type="kpi"` or `type="anything_else"`, it disappeared
   from both `rendered` and `dropped`, leaving the audit blind.
"""
from pathlib import Path
import pandas as pd
import pytest

from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.extractor import (
    extract_for_render,
    _smart_to_datetime_series,
    _format_datetime_value,
)


def _save(tmp_path, name, **sheets) -> Path:
    p = tmp_path / f"{name}.xlsx"
    with pd.ExcelWriter(p, engine="openpyxl") as w:
        for sheet_name, df in sheets.items():
            df.to_excel(w, sheet_name=sheet_name, index=False)
    return p


class TestSmartToDatetimeSeries:
    """Direct unit tests on the helper."""

    def test_latam_ambiguous_slash_treated_as_dayfirst(self):
        # All values are DD <= 12 AND MM <= 12 — fully ambiguous. We should
        # default to LATAM convention (Spanish-speaking audience).
        s = pd.Series(["01/03/2024", "05/03/2024", "08/03/2024", "12/06/2024"])
        out = _smart_to_datetime_series(s)
        # Expected: 1 marzo, 5 marzo, 8 marzo, 12 junio
        assert list(out.dt.month) == [3, 3, 3, 6]
        assert list(out.dt.day) == [1, 5, 8, 12]

    def test_latam_definitive_slash(self):
        # "31/12" forces dayfirst — 31 is not a valid month.
        s = pd.Series(["31/12/2024", "15/01/2024", "20/02/2024"])
        out = _smart_to_datetime_series(s)
        assert list(out.dt.month) == [12, 1, 2]
        assert list(out.dt.day) == [31, 15, 20]

    def test_us_definitive_slash(self):
        # "12/31" forces month-first — 31 is not a valid month, so it must
        # be the day, and 12 must be the month.
        s = pd.Series(["12/31/2024", "01/15/2024", "02/20/2024"])
        out = _smart_to_datetime_series(s)
        assert list(out.dt.month) == [12, 1, 2]
        assert list(out.dt.day) == [31, 15, 20]

    def test_iso_dates_unaffected(self):
        s = pd.Series(["2024-03-01", "2024-06-15", "2024-12-31"])
        out = _smart_to_datetime_series(s)
        assert list(out.dt.month) == [3, 6, 12]
        assert list(out.dt.day) == [1, 15, 31]

    def test_real_datetime_passthrough(self):
        # Real datetime objects must not be re-parsed.
        s = pd.Series(pd.date_range("2024-03-01", periods=4, freq="D"))
        out = _smart_to_datetime_series(s)
        assert list(out.dt.month) == [3, 3, 3, 3]
        assert list(out.dt.day) == [1, 2, 3, 4]

    def test_mixed_garbage_coerces_to_NaT(self):
        s = pd.Series(["01/03/2024", "not a date", "05/03/2024"])
        out = _smart_to_datetime_series(s)
        assert pd.isna(out.iloc[1])
        # Valid ones still get LATAM treatment
        assert out.iloc[0].month == 3 and out.iloc[0].day == 1


class TestFormatDatetimeValueSingleString:
    """The single-value formatter also needs to do the right thing on
    a string like '01/03/2024' that arrives standalone."""

    def test_definitive_dayfirst_string(self):
        # 31 > 12 ⇒ must be DD/MM
        assert _format_datetime_value("31/12/2024") == "2024-12-31"

    def test_definitive_monthfirst_string(self):
        # 12/31 ⇒ MM/DD (US)
        assert _format_datetime_value("12/31/2024") == "2024-12-31"

    def test_ambiguous_slash_defaults_to_dayfirst(self):
        # 01/03/2024 must be 1 de marzo (LATAM default)
        assert _format_datetime_value("01/03/2024") == "2024-03-01"

    def test_iso_string_unchanged(self):
        assert _format_datetime_value("2024-03-01") == "2024-03-01"


class TestTimeSeriesChartRespectsDayfirst:
    """End-to-end: a LATAM-formatted Excel must produce a chart whose
    months reflect the real distribution, not the MM/DD interpretation."""

    def test_latam_dates_produce_correct_month_bins(self, tmp_path):
        # 11 sales: 4 in March, 3 in April, 2 in May, 2 in June 2024. All
        # written as DD/MM/YYYY strings with DD <= 12 (so fully ambiguous).
        df = pd.DataFrame({
            "fecha": [
                "01/03/2024", "05/03/2024", "08/03/2024", "12/03/2024",  # marzo
                "02/04/2024", "07/04/2024", "11/04/2024",                # abril
                "03/05/2024", "09/05/2024",                              # mayo
                "04/06/2024", "10/06/2024",                              # junio
            ],
            "ventas": [100, 200, 150, 180, 220, 175, 190, 210, 230, 250, 240],
        })
        p = _save(tmp_path, "latam_ts", Ventas=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        ts_blocks = [b for b in inv if b.kind == "time_series_candidate"]
        assert ts_blocks, "Expected a time series block for fecha+ventas"

        slides = [{
            "type": "chart", "chart_type": "line",
            "title": "Ventas", "narrative": "",
            "block_ref": ts_blocks[0].id,
        }]
        rendered, _ = extract_for_render(slides, inv, wb, p)
        assert rendered, "Time series chart should render with valid data"
        data = rendered[0]["data"]

        # Bug guard: must NOT span 11 different months. With span > 60 days
        # and < 365 days, aggregation is monthly; the real data covers
        # exactly 4 months (Mar/Apr/May/Jun). So labels should be 4 bins.
        months_seen = {lbl for lbl in data["labels"]}
        assert len(months_seen) == 4, (
            f"Expected exactly 4 monthly bins (Mar-Jun), got "
            f"{len(months_seen)}: {sorted(months_seen)}"
        )
        # And those bins must be 2024-03 .. 2024-06 — not Jan/Feb/Jul/etc.
        for m in months_seen:
            assert m in ("2024-03", "2024-04", "2024-05", "2024-06"), (
                f"Unexpected bin {m!r} — months were misparsed"
            )

        # Sanity on totals: March should aggregate 100+200+150+180 = 630
        bins = dict(zip(data["labels"], data["values"]))
        assert bins["2024-03"] == 630.0
        assert bins["2024-04"] == 220 + 175 + 190
        assert bins["2024-05"] == 210 + 230
        assert bins["2024-06"] == 250 + 240


class TestUnsupportedSlideTypeDoesNotDisappear:
    """Bug #2: a slide with an unknown `type` used to vanish silently —
    not in `rendered`, not in `dropped`, no warning. Audit + intent_report
    were blind to it. After fix, it must show up in `dropped` with reason
    `unsupported_slide_type`."""

    def test_unknown_type_appears_in_dropped(self, tmp_path):
        df = pd.DataFrame({
            "ciudad": ["A", "B", "C"] * 5,
            "monto": [100 + i * 10 for i in range(15)],
        })
        p = _save(tmp_path, "src", Hoja1=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        # Pick any block to reference (so the lookup succeeds and we hit
        # the type-dispatch branch).
        any_block = next((b for b in inv if b.kind != "table"), inv[0])
        slides = [
            # Singular `kpi` is the most common AI hallucination — not in our
            # type vocab (we use `kpi_row`).
            {"type": "kpi", "title": "Bogus KPI", "block_ref": any_block.id},
            {"type": "summary", "title": "Fake summary", "block_ref": any_block.id},
            {"type": "carousel", "title": "?", "block_ref": any_block.id},
        ]
        rendered, dropped = extract_for_render(slides, inv, wb, p)
        # None of these should sneak into rendered.
        assert all(s.get("type") in {"title", "kpi_row", "chart", "table",
                                       "text_bullets"} for s in rendered)
        # All three should be visible in dropped with the right reason.
        reasons = [d["reason"] for d in dropped if d.get("type") in
                   ("kpi", "summary", "carousel")]
        assert reasons.count("unsupported_slide_type") == 3, (
            f"Expected 3 unsupported_slide_type drops, got {reasons}"
        )
