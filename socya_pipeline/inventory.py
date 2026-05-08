"""Catalog data blocks with provenance + quality flags."""
from dataclasses import dataclass, field
from typing import List, Set
from socya_pipeline.parser import WorkbookData, SheetData, ColumnData

@dataclass
class Provenance:
    sheet: str
    columns: List[str]
    rows: tuple              # (start, end), 0-indexed inclusive

@dataclass
class Block:
    id: str
    kind: str                # "table" | "kpi_candidate" | "categorical_distribution" | "time_series_candidate"
    label: str
    provenance: Provenance
    quality_flags: Set[str] = field(default_factory=set)
    extra: dict = field(default_factory=dict)

def build_inventory(wb: WorkbookData) -> List[Block]:
    blocks: List[Block] = []
    counter = {"T": 0, "K": 0, "C": 0, "S": 0}

    for sheet in wb.sheets:
        # 1. Whole-sheet table block
        counter["T"] += 1
        bid = f"T{counter['T']}"
        flags = set()
        if sheet.fill_ratio < 0.3:
            flags.add("low_fill_ratio")
        if sheet.shape[0] < 2:
            flags.add("too_few_rows")
        blocks.append(Block(
            id=bid, kind="table",
            label=f"{sheet.name} (tabla completa)",
            provenance=Provenance(sheet=sheet.name,
                                  columns=[c.name for c in sheet.columns],
                                  rows=(0, max(0, sheet.shape[0] - 1))),
            quality_flags=flags,
            extra={"shape": list(sheet.shape), "fill_ratio": sheet.fill_ratio},
        ))

        # 2. KPI candidates (numeric/currency aggregations) — skip ID-like columns
        for col in sheet.columns:
            if col.dtype in ("numeric", "currency") and col.fill_ratio >= 0.5:
                if _is_id_like(col, sheet.shape[0]):
                    continue
                counter["K"] += 1
                bid = f"K{counter['K']}"
                flags = set()
                if col.sum is not None and col.sum == 0:
                    flags.add("all_zero")
                # Currency cols → sum aggregation. Other numerics → mean (more
                # meaningful than sum for things like "days", "ages", "scores").
                if col.dtype == "currency":
                    label = f"Total de {col.name.strip()}"
                    value = col.sum
                    agg = "sum"
                else:
                    label = f"Promedio de {col.name.strip()}"
                    value = col.mean
                    agg = "mean"
                blocks.append(Block(
                    id=bid, kind="kpi_candidate",
                    label=label,
                    provenance=Provenance(sheet=sheet.name, columns=[col.name],
                                          rows=(0, sheet.shape[0] - 1)),
                    quality_flags=flags,
                    extra={"agg": agg, "value": value,
                            "min": col.min, "max": col.max, "mean": col.mean},
                ))

        # 3. Categorical distributions
        for col in sheet.columns:
            if col.dtype == "categorical" and 2 <= col.n_unique <= 20:
                counter["C"] += 1
                bid = f"C{counter['C']}"
                flags = set()
                if col.top_values:
                    top_pct = col.top_values[0][1] / max(1, sum(c[1] for c in col.top_values))
                    if top_pct > 0.85:
                        flags.add("single_dominant_category")
                blocks.append(Block(
                    id=bid, kind="categorical_distribution",
                    label=f"Distribución por {col.name}",
                    provenance=Provenance(sheet=sheet.name, columns=[col.name],
                                          rows=(0, sheet.shape[0] - 1)),
                    quality_flags=flags,
                    extra={"top_values": col.top_values, "n_unique": col.n_unique},
                ))

        # 4. Time series candidates (date column + numeric column pair)
        date_cols = [c for c in sheet.columns if c.dtype == "date"]
        num_cols = [c for c in sheet.columns if c.dtype in ("numeric", "currency")]
        for d in date_cols[:1]:
            for n in num_cols[:2]:
                counter["S"] += 1
                bid = f"S{counter['S']}"
                blocks.append(Block(
                    id=bid, kind="time_series_candidate",
                    label=f"{n.name} a lo largo de {d.name}",
                    provenance=Provenance(sheet=sheet.name, columns=[d.name, n.name],
                                          rows=(0, sheet.shape[0] - 1)),
                    extra={"x": d.name, "y": n.name},
                ))

    return blocks


_ID_NAME_TOKENS = ("id", "código", "codigo", "número", "numero", "n°", "no.",
                    "folio", "consecutivo", "uuid", "key")

def _is_id_like(col, n_rows: int) -> bool:
    """A column is 'ID-like' (and a bad KPI candidate) when its name signals an
    identifier (id, código, número, folio, consecutivo). Currency columns are
    NEVER IDs — currency uniqueness reflects per-row value, not identifier."""
    if col.dtype == "currency":
        return False
    name = (col.name or "").strip().lower()
    for token in _ID_NAME_TOKENS:
        if (name == token
                or name.startswith(token + " ")
                or name.endswith(" " + token)
                or f"_{token}" in name
                or f" {token} " in f" {name} "):
            return True
    return False
