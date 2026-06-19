"""Catalog data blocks with provenance + quality flags.

Inventory is the *catalog* layer: it scans each sheet and emits Blocks
representing tables, KPI candidates, categorical distributions, and time
series. It also calls the `insights` layer (optional intelligence) to
mark subsumed KPIs and surface derived KPIs (ratios). If `insights` raises,
the inventory degrades silently to the basic behavior.
"""
from dataclasses import dataclass, field
from typing import List, Set
from socya_pipeline.parser import WorkbookData, SheetData, ColumnData
from socya_pipeline import insights

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

    # Multi-sheet pre-pass: flag redundant sheets so their blocks get
    # deprioritized downstream (we still emit them — never silently delete
    # user data — but with a flag the planner / completer can read).
    try:
        redundant = set(insights.redundant_sheet_ids(wb.sheets))
    except Exception:
        redundant = set()

    for sheet in wb.sheets:
        is_redundant = sheet.name in redundant
        # 1. Whole-sheet table block
        counter["T"] += 1
        bid = f"T{counter['T']}"
        flags = set()
        if sheet.fill_ratio < 0.5:
            flags.add("low_fill_ratio")
        if sheet.shape[0] < 2:
            flags.add("too_few_rows")
        if is_redundant:
            flags.add("redundant_sheet")
        blocks.append(Block(
            id=bid, kind="table",
            label=f"{sheet.name} (tabla completa)",
            provenance=Provenance(sheet=sheet.name,
                                  columns=[c.name for c in sheet.columns],
                                  rows=(0, max(0, sheet.shape[0] - 1))),
            quality_flags=flags,
            extra={"shape": list(sheet.shape), "fill_ratio": sheet.fill_ratio},
        ))

        # 2. KPI candidates (numeric / currency / percent / score aggregations).
        # Skip year, ID-like, and non-semantic column names ('%', 'Unnamed: 5').
        for col in sheet.columns:
            if col.dtype == "year":
                continue
            if col.dtype not in ("numeric", "currency", "percent", "score"):
                continue
            if col.fill_ratio < 0.5:
                continue
            if _is_id_like(col, sheet.shape[0]):
                continue
            if _is_non_semantic_column(col.name):
                continue
            # Skip KPIs whose aggregation comes out to zero or NaN — those
            # are meaningless headlines (would render as "Total = 0" cards).
            if col.dtype == "currency" and (col.sum is None or col.sum == 0):
                continue
            if col.dtype != "currency" and (col.mean is None
                                              or (isinstance(col.mean, float)
                                                  and (col.mean != col.mean))):
                continue  # NaN check via != trick
            counter["K"] += 1
            bid = f"K{counter['K']}"
            flags = set()
            if col.sum is not None and col.sum == 0:
                flags.add("all_zero")
            # KPI on a redundant sheet (e.g. 31-row sample) is misleading
            # when the same column exists on the authoritative master with
            # the real total. Tag it so the planner/extractor can deprioritize.
            if is_redundant:
                flags.add("redundant_sheet")
            # Choose aggregation by dtype.
            #   currency → sum (totals are headlines)
            #   percent  → mean (a % does not sum meaningfully)
            #   score    → mean (rating average)
            #   numeric  → mean (safer default — sum requires money semantics)
            if col.dtype == "currency":
                # If a status column (Estado / Status) was detected in the
                # same sheet AND at least one row is in a negative state
                # (RECHAZADO, ANULADO, …), the raw sum overstates "money
                # moved". Re-label the headline and flag for disclaimer.
                status_bd = (sheet.status_breakdowns or {}).get(col.name)
                has_negative = bool(status_bd
                                     and status_bd.get("negative_count", 0) > 0)
                if has_negative:
                    label = _build_kpi_label("Total bruto solicitado",
                                              col.name.strip())
                    flags.add("needs_status_disclaimer")
                else:
                    label = _build_kpi_label("Total", col.name.strip())
                value = col.sum
                agg = "sum"
                display_unit = None
            elif col.dtype == "percent":
                label = _build_kpi_label("Promedio", col.name.strip())
                value = col.mean
                agg = "mean"
                display_unit = "%"
            elif col.dtype == "score":
                label = _build_kpi_label("Calificación promedio", col.name.strip())
                value = col.mean
                agg = "mean"
                display_unit = None
            else:
                label = _build_kpi_label("Promedio", col.name.strip())
                value = col.mean
                agg = "mean"
                display_unit = None
            extra = {"agg": agg, "value": value,
                      "min": col.min, "max": col.max, "mean": col.mean}
            if display_unit:
                extra["display_unit"] = display_unit
            # Si el parser detectó una unidad explícita en el header
            # ("(USD)", "kg", etc.) la propagamos para que el extractor
            # pueda mostrarla en el KPI value sin tener que adivinar.
            col_unit = getattr(col, "unit", None)
            if col_unit and col_unit != display_unit:
                extra["unit_hint"] = col_unit
            # Attach the status breakdown so the renderer can show the
            # "incluye N rechazados ($X)" footnote and the planner prompt
            # can warn the LLM about the disclaimer requirement.
            if col.dtype == "currency":
                status_bd = (sheet.status_breakdowns or {}).get(col.name)
                if status_bd and status_bd.get("negative_count", 0) > 0:
                    extra["status_filter"] = status_bd
            blocks.append(Block(
                id=bid, kind="kpi_candidate",
                label=label,
                provenance=Provenance(sheet=sheet.name, columns=[col.name],
                                      rows=(0, sheet.shape[0] - 1)),
                quality_flags=flags,
                extra=extra,
            ))

        # 3. Categorical distributions — emitted for two cases:
        #    (a) `dtype=categorical`: low-cardinality column (2-20 unique
        #        values) → natural pie/bar chart
        #    (b) `dtype=text` with meaningful Top-N: e.g. "Ciudad Destino"
        #        with 318 uniques but where the top 5 cover a real share.
        #        Without (b), authoritative sheets (1852 rows where every
        #        city name is high-card text) silently lose their Top-N
        #        chart candidates and the AI is forced to fall back to
        #        smaller sister sheets where the same column squeaks under
        #        the 20-unique categorical threshold.
        for col in sheet.columns:
            if _is_non_semantic_column(col.name):
                continue

            if col.dtype == "categorical":
                if not (2 <= col.n_unique <= 20):
                    continue
                if col.fill_ratio < 0.3:
                    continue
                source = "categorical"
            elif col.dtype == "text":
                # Top-N candidacy: must have top_values, must have at least
                # some duplication (n_unique < rows so there's a "top"
                # to speak of), and the leader must be material (≥3% of
                # populated rows). Hard ceiling on uniqueness prevents
                # free-text fields like "Objeto" / "Agenda" from leaking
                # through — those would chart as N bars of count=1 each.
                if not col.top_values:
                    continue
                if col.fill_ratio < 0.3:
                    continue
                rows = sheet.shape[0] or 1
                if col.n_unique >= rows * 0.5:
                    continue  # too unique → no meaningful "top"
                top_count = col.top_values[0][1]
                if top_count < max(5, rows * 0.03):
                    continue  # leader too small to anchor a chart
                # No hard cap on n_unique — Top-N rendering only shows the
                # first 10 categories anyway (handled by _build_chart_data
                # via value_counts().head(15) + _consolidate_long_tail). A
                # 318-unique "Ciudad Destino" where Medellín has 252 hits is
                # exactly the chart an audit slide should show.
                source = "top_n_text"
            else:
                continue

            counter["C"] += 1
            bid = f"C{counter['C']}"
            flags = set()
            if col.top_values:
                top_pct = col.top_values[0][1] / max(1, sum(c[1] for c in col.top_values))
                if top_pct > 0.85:
                    flags.add("single_dominant_category")
            if col.fill_ratio < 0.5:
                flags.add("low_fill_ratio")
            if is_redundant:
                flags.add("redundant_sheet")
            if source == "top_n_text":
                # Distinguish from low-card categoricals so downstream
                # consumers (planner prompt, narratives) know this is a
                # Top-N slice rather than a full distribution.
                flags.add("top_n_text")
            blocks.append(Block(
                id=bid, kind="categorical_distribution",
                label=(f"Top categorías de {col.name}" if source == "top_n_text"
                        else f"Distribución por {col.name}"),
                provenance=Provenance(sheet=sheet.name, columns=[col.name],
                                      rows=(0, sheet.shape[0] - 1)),
                quality_flags=flags,
                extra={"top_values": col.top_values, "n_unique": col.n_unique},
            ))

        # 4. Time series candidates (date column + numeric column pair).
        # Use only money-meaningful numerics for Y; skip IDs, years, scores, percents.
        date_cols = [c for c in sheet.columns if c.dtype == "date"]
        num_cols = [c for c in sheet.columns
                    if c.dtype in ("numeric", "currency")
                    and not _is_id_like(c, sheet.shape[0])]
        for d in date_cols[:1]:
            for n in num_cols[:2]:
                counter["S"] += 1
                bid = f"S{counter['S']}"
                ts_flags = set()
                if is_redundant:
                    ts_flags.add("redundant_sheet")
                blocks.append(Block(
                    id=bid, kind="time_series_candidate",
                    label=f"{n.name} a lo largo de {d.name}",
                    provenance=Provenance(sheet=sheet.name, columns=[d.name, n.name],
                                          rows=(0, sheet.shape[0] - 1)),
                    quality_flags=ts_flags,
                    extra={"x": d.name, "y": n.name},
                ))

        # 5. Effective-total KPIs derived from status_breakdowns.
        # For every currency column whose sheet has a status filter with
        # any rejected/anulated row, we emit an "Efectivo" KPI (positive
        # states only) and, when rejection is material (≥5% of gross),
        # a "Rechazado" KPI so the deck can show how much money never moved.
        # These carry flags `effective_total` / `rejected_total` so
        # extractor._kpi_priority prefers them over the raw bruto.
        for cur_col_name, bd in (sheet.status_breakdowns or {}).items():
            if not bd or bd.get("negative_count", 0) == 0:
                continue
            # Skip when the parent currency column was dropped (low fill,
            # ID-like, non-semantic name) — emitting orphan KPIs is worse
            # than silence. The parent's existence is implied by any K*
            # block on the same column.
            parent_exists = any(
                b.kind == "kpi_candidate"
                and b.provenance.sheet == sheet.name
                and b.provenance.columns == [cur_col_name]
                for b in blocks
            )
            if not parent_exists:
                continue
            cur_col = next(
                (c for c in sheet.columns if c.name == cur_col_name), None,
            )
            cur_col_unit = getattr(cur_col, "unit", None) if cur_col else None
            base_extra = {
                "agg": "sum_filtered",
                "min": None, "max": None, "mean": None,
                "status_filter": bd,
                "parent_column": cur_col_name,
            }
            if cur_col_unit:
                base_extra["unit_hint"] = cur_col_unit
            gross = bd.get("gross_sum") or 0
            # 5a. Efectivo (positive states only)
            pos_sum = bd.get("positive_sum") or 0
            if pos_sum > 0:
                counter["K"] += 1
                bid = f"K{counter['K']}"
                eff_extra = dict(base_extra)
                eff_extra["value"] = pos_sum
                eff_extra["state_bucket"] = "positive"
                eff_extra["bucket_count"] = bd.get("positive_count", 0)
                eff_extra["bucket_states"] = bd.get("positive_states", [])
                eff_flags = {"effective_total"}
                if is_redundant:
                    eff_flags.add("redundant_sheet")
                blocks.append(Block(
                    id=bid, kind="kpi_candidate",
                    label=_build_kpi_label("Total efectivo", cur_col_name.strip()),
                    provenance=Provenance(sheet=sheet.name, columns=[cur_col_name],
                                          rows=(0, sheet.shape[0] - 1)),
                    quality_flags=eff_flags,
                    extra=eff_extra,
                ))
            # 5b. Rechazado (only if it is a material slice — ≥5% of gross
            # or absolute ≥ $1M — so we don't spam KPIs for trivial rejects)
            neg_sum = bd.get("negative_sum") or 0
            if neg_sum > 0 and (
                (gross > 0 and neg_sum / gross >= 0.05) or neg_sum >= 1_000_000
            ):
                counter["K"] += 1
                bid = f"K{counter['K']}"
                rej_extra = dict(base_extra)
                rej_extra["value"] = neg_sum
                rej_extra["state_bucket"] = "negative"
                rej_extra["bucket_count"] = bd.get("negative_count", 0)
                rej_extra["bucket_states"] = bd.get("negative_states", [])
                rej_flags = {"rejected_total"}
                if is_redundant:
                    rej_flags.add("redundant_sheet")
                blocks.append(Block(
                    id=bid, kind="kpi_candidate",
                    label=_build_kpi_label("Rechazado", cur_col_name.strip()),
                    provenance=Provenance(sheet=sheet.name, columns=[cur_col_name],
                                          rows=(0, sheet.shape[0] - 1)),
                    quality_flags=rej_flags,
                    extra=rej_extra,
                ))

    # Optional intelligence layer — degrade gracefully on any failure
    try:
        _mark_subsumed_kpis(blocks)
    except Exception:
        pass
    try:
        derived = _build_derived_ratio_blocks(blocks, counter)
        blocks.extend(derived)
    except Exception:
        pass

    return blocks


def _mark_subsumed_kpis(blocks: List[Block]) -> None:
    """If a KPI's value ≈ Σ of other KPIs on the same sheet, mark the others
    as 'subsumed_by_total'. The downstream auto-completer / planner can then
    deprioritize them in favor of the headline total."""
    subsumed_ids = insights.subsumed_kpi_ids(blocks)
    if not subsumed_ids:
        return
    for b in blocks:
        if b.id in subsumed_ids:
            b.quality_flags.add("subsumed_by_total")


def _build_derived_ratio_blocks(blocks: List[Block], counter) -> List[Block]:
    """Detect KPI pairs (Solicitado/Aprobado, Presupuestado/Ejecutado, …) and
    emit derived KPI blocks with the percentage value pre-computed."""
    pairs = insights.find_ratio_pairs(blocks)
    out: List[Block] = []
    for p in pairs:
        counter["K"] += 1
        bid = f"K{counter['K']}"
        num_b = p["num_block"]
        den_b = p["den_block"]
        pct = p["percent_value"]
        # Use the denominator's sheet for provenance (more stable)
        sheet = den_b.provenance.sheet
        cols = list({*num_b.provenance.columns, *den_b.provenance.columns})
        out.append(Block(
            id=bid,
            kind="kpi_candidate",
            label=p["label"],
            provenance=Provenance(sheet=sheet, columns=cols,
                                   rows=den_b.provenance.rows),
            quality_flags={"derived"},
            extra={
                "agg": "ratio",
                "value": round(pct, 1),
                "display_unit": "%",
                "numerator_block": num_b.id,
                "denominator_block": den_b.id,
            },
        ))
    return out


def _is_non_semantic_column(name: str) -> bool:
    """A column name that the reader can't interpret on its own.

    Excludes from KPIs / charts / distributions:
      - Pandas placeholders ('Unnamed: N')
      - Single-character or symbol-only names ('%', '$', '#', '*')
      - Empty strings, NaN/None placeholders

    These columns ARE valid in tables (the reader sees them in context with
    other columns) but make terrible standalone slide titles like
    'Distribución por Unnamed: 5' or 'Promedio de %'.
    """
    s = str(name or "").strip()
    if not s:
        return True
    if s.startswith("Unnamed:") or s.startswith("Unnamed "):
        return True
    if s.lower() in ("nan", "none", "null", "na", "n/a"):
        return True
    # Pure-symbol or single-char names (% / $ / # / * / ? / .)
    alphanumeric = sum(1 for c in s if c.isalnum())
    if alphanumeric <= 1:
        return True
    return False


def _build_kpi_label(prefix: str, col_name: str) -> str:
    """Avoid redundant 'Total de Total ...' labels. Normalize snake_case
    columns ('ventas_total' → 'Ventas Total'). If the column name already
    contains the prefix word anywhere, just use the column name."""
    cn = _humanize(col_name.strip())
    cn_lower = cn.lower()
    pl = prefix.lower()
    # If the prefix word is already in the column name, don't prepend it
    if (cn_lower.startswith(pl)
            or cn_lower.endswith(" " + pl)
            or f" {pl} " in f" {cn_lower} "):
        return cn
    if cn_lower.startswith("valor "):
        rest = cn[6:].strip()
        if rest.lower().startswith(pl):
            return rest
    return f"{prefix} de {cn}"


def _humanize(s: str) -> str:
    """Display-form column name (delegates to insights.humanize_header)."""
    return insights.humanize_header(s)


_ID_NAME_TOKENS = ("id", "código", "codigo", "número", "numero", "n°", "no.",
                    "folio", "consecutivo", "uuid", "key")

def _is_id_like(col, n_rows: int) -> bool:
    """A column is 'ID-like' (and a bad KPI / time-series Y candidate).

    Order:
    1. Name says ID → it IS an ID, even if the structural detector
       misclassified the column as currency (e.g. integer IDs ≥ 1000 burlan
       is_currency_by_magnitude). This is the authoritative signal.
    2. Otherwise, currency dtype is never an ID (currency cardinality reflects
       per-row $value, not identity).
    3. Otherwise, name doesn't say ID → not ID."""
    name = (col.name or "").strip().lower()
    for token in _ID_NAME_TOKENS:
        if (name == token
                or name.startswith(token + " ")
                or name.endswith(" " + token)
                or f"_{token}" in name
                or f" {token} " in f" {name} "):
            return True
    return False
