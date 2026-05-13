"""Heavy stress tests — performance, pathological Excels, concurrency.

Each test simulates a realistic abuse case the user might throw at the
pipeline. Failures here mean the program will crash or freeze on that
class of file in production.
"""
import math
import time
import threading
from pathlib import Path

import pandas as pd
import pytest

from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.extractor import extract_for_render, auto_complete_slides


def _save(tmp_path, name, **sheets) -> Path:
    p = tmp_path / f"{name}.xlsx"
    with pd.ExcelWriter(p, engine="openpyxl") as w:
        for sheet_name, df in sheets.items():
            df.to_excel(w, sheet_name=sheet_name, index=False)
    return p


# ─────────────────────────────────────────────────────────────────────
# Performance: large workbooks
# ─────────────────────────────────────────────────────────────────────

class TestPerformance:
    def test_10k_rows_under_10s(self, tmp_path):
        """A 10K-row Excel must parse + inventory in under 10 seconds."""
        N = 10_000
        cities = ["Bogotá", "Medellín", "Cali"]
        df = pd.DataFrame({
            "id":        list(range(N)),
            "categoria": [f"Cat{i % 12}" for i in range(N)],
            "ciudad":    [cities[i % 3] for i in range(N)],
            "fecha":     pd.date_range("2020-01-01", periods=N, freq="h"),
            "monto":     [100 + i * 1.5 for i in range(N)],
        })
        p = _save(tmp_path, "big", Hoja1=df)

        start = time.time()
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        elapsed = time.time() - start

        assert elapsed < 10.0, f"Parse+inventory took {elapsed:.2f}s (>10s)"
        assert wb.sheets[0].shape[0] == N
        # KPI for monto, distribution for ciudad+categoria, time series
        assert any(b.kind == "kpi_candidate" for b in inv)
        assert any(b.kind == "categorical_distribution" for b in inv)

    def test_many_sheets_does_not_crash(self, tmp_path):
        """20 sheets in one workbook — must not crash, must dedupe redundant."""
        sheets = {}
        for i in range(20):
            df = pd.DataFrame({
                "name":  [f"item{j}" for j in range(50)],
                "value": [j * 10 for j in range(50)],
            })
            sheets[f"Sheet{i}"] = df
        p = _save(tmp_path, "many_sheets", **sheets)
        wb = parse_workbook(p)
        inv = build_inventory(wb)

        # All 20 sheets parsed
        assert len(wb.sheets) == 20
        # Most should be flagged redundant (same schema)
        redundant = sum(1 for b in inv if "redundant_sheet" in b.quality_flags)
        assert redundant >= 10, f"Expected most sheets flagged redundant, got {redundant}"


# ─────────────────────────────────────────────────────────────────────
# Pathological structure
# ─────────────────────────────────────────────────────────────────────

class TestPathologicalStructure:
    def test_duplicate_column_names(self, tmp_path):
        """Excel with two columns named identically — pandas adds .1 suffix.
        Must not crash."""
        # Build raw: pandas can't write dup names directly via DataFrame init
        # so we cheat by writing rows manually
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["valor", "valor", "categoria"])
        for i in range(20):
            ws.append([i * 10, i * 5, f"cat{i % 4}"])
        p = tmp_path / "dup_cols.xlsx"
        wb.save(p)

        parsed = parse_workbook(p)
        inv = build_inventory(parsed)
        # Should not crash; dup col gets renamed to "valor.1" by pandas
        col_names = [c.name for c in parsed.sheets[0].columns]
        assert any("valor" in c for c in col_names)
        assert len(inv) > 0

    def test_headers_in_row_5(self, tmp_path):
        """Excel where rows 0-4 are titles/notes and real headers are in row 5."""
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.append(["REPORTE EJECUTIVO 2024"])
        ws.append([])
        ws.append(["Departamento de Análisis"])
        ws.append([])
        ws.append([])
        ws.append(["Producto", "Cantidad", "Precio"])  # real headers
        for i in range(15):
            ws.append([f"Prod{i}", (i + 1) * 10, (i + 1) * 100])
        p = tmp_path / "deep_header.xlsx"
        wb.save(p)
        # Must not crash even if header promotion can't perfectly recover
        parsed = parse_workbook(p)
        assert len(parsed.sheets) > 0

    def test_only_numeric_columns(self, tmp_path):
        """All columns are numeric — categorical/text branches must be skipped
        gracefully without producing empty/broken blocks."""
        df = pd.DataFrame({f"m{i}": [j * (i + 1) for j in range(40)]
                            for i in range(5)})
        p = _save(tmp_path, "all_numeric", Data=df)
        parsed = parse_workbook(p)
        inv = build_inventory(parsed)
        # Should produce KPI candidates but no categorical distributions
        kpis = [b for b in inv if b.kind == "kpi_candidate"]
        cats = [b for b in inv if b.kind == "categorical_distribution"]
        assert len(kpis) > 0
        assert len(cats) == 0  # nothing categorical

    def test_only_text_columns(self, tmp_path):
        """All columns are text/categorical — no KPIs possible, must not crash."""
        df = pd.DataFrame({
            "nombre":   [f"Name {i}" for i in range(30)],
            "rol":      ["A", "B", "C", "D"] * 7 + ["A", "B"],
            "ciudad":   ["X", "Y", "Z"] * 10,
        })
        p = _save(tmp_path, "all_text", Data=df)
        parsed = parse_workbook(p)
        inv = build_inventory(parsed)
        kpis = [b for b in inv if b.kind == "kpi_candidate"]
        # Zero KPIs possible — pipeline must handle this without crashing
        assert len(kpis) == 0
        # But categorical distributions for the low-cardinality cols
        cats = [b for b in inv if b.kind == "categorical_distribution"]
        assert len(cats) >= 1


# ─────────────────────────────────────────────────────────────────────
# Unicode and encoding edge cases
# ─────────────────────────────────────────────────────────────────────

class TestUnicode:
    def test_chinese_japanese_arabic_values(self, tmp_path):
        """Non-Latin scripts in cell values must round-trip cleanly."""
        df = pd.DataFrame({
            "lang":   ["chinese", "japanese", "arabic", "russian", "greek"],
            "saludo": ["你好", "こんにちは", "مرحبا", "Привет", "Γειά"],
            "monto":  [100, 200, 300, 400, 500],
        })
        p = _save(tmp_path, "unicode", Data=df)
        parsed = parse_workbook(p)
        # Check that values survived encoding
        sheet = parsed.sheets[0]
        saludo_samples = [c.samples for c in sheet.columns
                            if c.name == "saludo"][0]
        assert "你好" in saludo_samples
        assert "مرحبا" in saludo_samples

    def test_emoji_in_cell_values(self, tmp_path):
        df = pd.DataFrame({
            "estado": ["✓", "✗", "⚠", "✓", "✗"],
            "valor":  [100, 0, 50, 80, 30],
        })
        p = _save(tmp_path, "emoji", Data=df)
        parsed = parse_workbook(p)
        # Should not crash; emoji column should be categorical
        cols = {c.name: c.dtype for c in parsed.sheets[0].columns}
        assert cols["estado"] in ("categorical", "bool", "text")

    def test_control_chars_stripped_or_handled(self, tmp_path):
        """Cells with embedded \\t, \\n, etc. should not break the pipeline."""
        df = pd.DataFrame({
            "raw":  ["normal", "with\ttab", "with\nnewline", "with\r\ncrlf", "ok"],
            "n":    [1, 2, 3, 4, 5],
        })
        p = _save(tmp_path, "ctrl", Data=df)
        parsed = parse_workbook(p)
        # Just verify no crash and samples can be JSON-serialized
        import json
        try:
            json.dumps([c.samples for c in parsed.sheets[0].columns],
                        default=str)
        except (TypeError, ValueError) as e:
            pytest.fail(f"Samples not JSON-serializable: {e}")


# ─────────────────────────────────────────────────────────────────────
# Concurrency: parsing same file from two threads
# ─────────────────────────────────────────────────────────────────────

class TestConcurrency:
    def test_parallel_parse_no_state_sharing(self, tmp_path):
        """Parse the same workbook from 2 threads simultaneously — no shared
        mutable state should corrupt either result."""
        df = pd.DataFrame({
            "ciudad": ["A", "B", "C"] * 30,
            "monto":  [(i + 1) * 100 for i in range(90)],
        })
        p = _save(tmp_path, "shared", Hoja1=df)

        results = []
        errors = []

        def worker():
            try:
                wb = parse_workbook(p)
                inv = build_inventory(wb)
                results.append({
                    "shape": wb.sheets[0].shape,
                    "n_blocks": len(inv),
                })
            except Exception as e:
                errors.append(e)

        t1 = threading.Thread(target=worker)
        t2 = threading.Thread(target=worker)
        t1.start(); t2.start()
        t1.join(timeout=15); t2.join(timeout=15)

        assert not errors, f"Concurrent parse raised: {errors}"
        assert len(results) == 2
        # Both runs must produce identical structure (parser is pure)
        assert results[0]["shape"] == results[1]["shape"]
        assert results[0]["n_blocks"] == results[1]["n_blocks"]


# ─────────────────────────────────────────────────────────────────────
# Resilience: corrupt / unusual files
# ─────────────────────────────────────────────────────────────────────

class TestResilience:
    def test_not_a_real_excel_file(self, tmp_path):
        """A .xlsx file that's actually plain text — should raise PipelineError,
        not a generic crash."""
        p = tmp_path / "fake.xlsx"
        p.write_text("This is not actually an Excel file.")
        from socya_pipeline.errors import PipelineError
        with pytest.raises(PipelineError) as exc_info:
            parse_workbook(p)
        # Must be a clean, user-actionable error
        assert exc_info.value.code.value == "EXCEL_INVALID"

    def test_zero_byte_file(self, tmp_path):
        """A 0-byte file — should raise PipelineError cleanly."""
        p = tmp_path / "empty.xlsx"
        p.write_bytes(b"")
        from socya_pipeline.errors import PipelineError
        with pytest.raises(PipelineError):
            parse_workbook(p)

    def test_extractor_handles_block_ref_to_missing_sheet(self, tmp_path):
        """If a Block points to a sheet that doesn't exist in the file
        (e.g. cache mismatch), the extractor must drop it gracefully."""
        df = pd.DataFrame({
            "ciudad": ["A", "B", "C"] * 10,
            "monto":  list(range(30)),
        })
        p = _save(tmp_path, "src", RealSheet=df)
        wb = parse_workbook(p)
        inv = build_inventory(wb)

        # Pick a real block, then mutate its sheet name to something fake
        chart_block = next(b for b in inv if b.kind == "categorical_distribution")
        chart_block.provenance.sheet = "DoesNotExist"

        slides = [{
            "type": "chart", "chart_type": "bar",
            "block_ref": chart_block.id, "title": "Test", "narrative": "",
        }]
        # Must drop, not crash
        rendered, dropped = extract_for_render(slides, inv, wb, p)
        # Either dropped OR rendered with empty data; never an exception
        assert isinstance(rendered, list)
        assert isinstance(dropped, list)


# ─────────────────────────────────────────────────────────────────────
# Real-world regression: the actual Comisiones V1 file
# ─────────────────────────────────────────────────────────────────────

class TestRegressionRealFile:
    def test_comisiones_v1_full_pipeline(self):
        """Run the full pipeline on the user's real Excel. This is the
        ultimate regression test — if this stops working, real users break."""
        p = Path("Comisiones V1.xlsx")
        if not p.exists():
            pytest.skip("Comisiones V1.xlsx not in cwd")
        wb = parse_workbook(p)
        inv = build_inventory(wb)
        # Sanity: real file has multiple sheets and produces blocks
        assert len(wb.sheets) >= 2
        assert len(inv) >= 5
        # No KPI should have a NaN/Inf value after our fix
        for b in inv:
            if b.kind == "kpi_candidate":
                v = b.extra.get("value")
                if v is not None:
                    assert math.isfinite(float(v)), \
                        f"KPI {b.id} ({b.label}) has non-finite value: {v}"
