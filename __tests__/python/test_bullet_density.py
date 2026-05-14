"""Tests para top-up de bullets y auto-add de hallazgos por hoja."""
import pytest
from socya_pipeline.extractor import _topup_bullets, _add_hallazgos_per_sheet
from socya_pipeline.parser import WorkbookData, SheetData, ColumnData
from socya_pipeline.inventory import build_inventory


def _make_wb_with_money_and_categories(sheet_name="Comisiones"):
    return WorkbookData(filename="test.xlsx", sheets=[
        SheetData(
            name=sheet_name,
            shape=(100, 3),
            fill_ratio=0.95,
            columns=[
                ColumnData(name="Estado", dtype="categorical",
                            n_unique=4, fill_ratio=1.0,
                            samples=["A", "B", "C", "D"],
                            top_values=[("CONTABILIZADO", 60),
                                          ("RECHAZADO", 25),
                                          ("PENDIENTE", 10),
                                          ("ANULADO", 5)]),
                ColumnData(name="Monto", dtype="currency",
                            n_unique=80, fill_ratio=1.0,
                            samples=[1000.0, 2000.0, 5000.0],
                            sum=500000.0, mean=5000.0,
                            min=100.0, max=50000.0),
                ColumnData(name="Ciudad", dtype="categorical",
                            n_unique=3, fill_ratio=0.9,
                            samples=["BOG", "MED", "CAL"],
                            top_values=[("BOGOTÁ", 70),
                                          ("MEDELLÍN", 20),
                                          ("CALI", 10)]),
            ],
            first_rows=[],
        )
    ])


# ─── Top-up ─────────────────────────────────────────────────────────

def test_topup_when_ai_returned_too_few():
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")

    # AI gave only 2 bullets
    existing = ["El total es $500K.", "Hay 100 registros."]
    out = _topup_bullets(existing, table, wb)
    assert len(out) >= 5
    # Original bullets preserved
    assert out[0] == "El total es $500K."
    assert out[1] == "Hay 100 registros."


def test_topup_when_ai_already_full():
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")
    existing = [f"Bullet {i}" for i in range(6)]
    out = _topup_bullets(existing, table, wb)
    # No change when ≥5 already
    assert out == existing


def test_topup_dedup_against_existing():
    """Si AI ya dijo 'total acumulado de Monto', no agregamos otro de
    la misma columna sumando lo mismo."""
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")
    existing = ["El total acumulado de 'Monto' suma $500K."]
    out = _topup_bullets(existing, table, wb)
    # Only ONE bullet about total of Monto
    total_bullets = [b for b in out if "total" in b.lower() and "monto" in b.lower()]
    assert len(total_bullets) <= 1


def test_topup_mix_angles_not_repeats_same_generator():
    """5 bullets agregados deben venir de generators distintos."""
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")
    out = _topup_bullets([], table, wb)
    assert len(out) >= 5
    # Heuristic: bullets shouldn't all start with the same word
    starting_words = [b.split()[0].lower() for b in out if b]
    assert len(set(starting_words)) >= 3


def test_topup_caps_at_max_bullets():
    wb = _make_wb_with_money_and_categories()
    blocks = build_inventory(wb)
    table = next(b for b in blocks if b.kind == "table")
    out = _topup_bullets([], table, wb)
    assert len(out) <= 6


# ─── Auto-add hallazgos per sheet ───────────────────────────────────

def test_auto_complete_adds_hallazgos_per_sheet():
    wb = _make_wb_with_money_and_categories("HojaA")
    # Add second sheet
    wb.sheets.append(SheetData(
        name="HojaB", shape=(50, 2), fill_ratio=0.9,
        columns=[
            ColumnData(name="Categoria", dtype="categorical",
                        n_unique=3, fill_ratio=1.0,
                        samples=["X", "Y", "Z"],
                        top_values=[("X", 30), ("Y", 15), ("Z", 5)]),
            ColumnData(name="Valor", dtype="numeric",
                        n_unique=20, fill_ratio=0.9,
                        samples=[10.0, 20.0],
                        sum=1000.0, mean=20.0, min=5.0, max=50.0),
        ],
        first_rows=[],
    ))
    blocks = build_inventory(wb)
    blocks_by_id = {b.id: b for b in blocks}

    rendered = [
        {"type": "title", "data": {"title": "T"}},
        {"type": "chart", "block_ref": next(b.id for b in blocks
                                              if b.kind == "categorical_distribution"
                                              and b.provenance.sheet == "HojaA"),
         "provenance": {"sheet": "HojaA"}},
        {"type": "chart", "block_ref": next(b.id for b in blocks
                                              if b.kind == "categorical_distribution"
                                              and b.provenance.sheet == "HojaB"),
         "provenance": {"sheet": "HojaB"}},
    ]
    out = _add_hallazgos_per_sheet(rendered, blocks, wb, blocks_by_id)
    sheets_with_hallazgos = {s.get("provenance", {}).get("sheet")
                              for s in out
                              if s.get("type") == "text_bullets"}
    assert "HojaA" in sheets_with_hallazgos
    assert "HojaB" in sheets_with_hallazgos


def test_auto_complete_skips_sheet_with_existing_hallazgos():
    wb = _make_wb_with_money_and_categories("HojaA")
    blocks = build_inventory(wb)
    blocks_by_id = {b.id: b for b in blocks}

    rendered = [
        {"type": "title", "data": {"title": "T"}},
        {"type": "chart", "block_ref": next(b.id for b in blocks
                                              if b.kind == "categorical_distribution"),
         "provenance": {"sheet": "HojaA"}},
        {"type": "text_bullets",
         "block_ref": next(b.id for b in blocks if b.kind == "table"),
         "supports_block": next(b.id for b in blocks if b.kind == "table"),
         "bullets": ["Existing bullet"],
         "provenance": {"sheet": "HojaA"}},
    ]
    out = _add_hallazgos_per_sheet(rendered, blocks, wb, blocks_by_id)
    hallazgos_for_a = [s for s in out
                        if s.get("type") == "text_bullets"
                        and (s.get("provenance") or {}).get("sheet") == "HojaA"]
    assert len(hallazgos_for_a) == 1   # not duplicated


def test_auto_complete_caps_added_hallazgos():
    """Si hay 10 hojas cubiertas, no añade 10 hallazgos — cap al razonable."""
    sheets = []
    for i in range(10):
        sheets.append(SheetData(
            name=f"Hoja{i}", shape=(20, 2), fill_ratio=0.9,
            columns=[
                ColumnData(name="C", dtype="categorical",
                            n_unique=3, fill_ratio=1.0,
                            samples=["a", "b"],
                            top_values=[("a", 10), ("b", 5)]),
            ],
            first_rows=[],
        ))
    wb = WorkbookData(filename="t.xlsx", sheets=sheets)
    blocks = build_inventory(wb)
    blocks_by_id = {b.id: b for b in blocks}

    rendered = [{"type": "title", "data": {"title": "T"}}]
    for i in range(10):
        bid = next((b.id for b in blocks
                    if b.kind == "categorical_distribution"
                    and b.provenance.sheet == f"Hoja{i}"), None)
        if bid:
            rendered.append({"type": "chart", "block_ref": bid,
                              "provenance": {"sheet": f"Hoja{i}"}})

    out = _add_hallazgos_per_sheet(rendered, blocks, wb, blocks_by_id)
    added = [s for s in out if s.get("type") == "text_bullets"]
    assert len(added) <= 3   # cap configured in MAX_AUTO_HALLAZGOS
