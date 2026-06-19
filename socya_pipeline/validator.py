"""Strict validator: drops slides without traceable provenance or low quality."""
import difflib
from dataclasses import dataclass, field
from typing import List, Dict, Optional
import re
from socya_pipeline.parser import WorkbookData


def _fuzzy_resolve_block_ref(ref: str, valid_ids: List[str]) -> Optional[str]:
    """Intenta recuperar un block_ref con typo del AI hacia el ID real.

    El AI a veces manda variantes case-insensitive o con typo chico (ej.
    'comision_pct' cuando el real es 'Comision_Pct', o 'K2_total' cuando
    es 'K2'). Sin esto, la slide entera se dropea por block_ref_not_found
    aunque el AI estaba "casi acertado". Probamos 3 niveles:

    1. Match case-insensitive exacto
    2. Match ignorando separadores (_ - .)
    3. difflib close match con cutoff alto (~Levenshtein 1-2)

    Conservador: cutoff=0.85 evita falsos positivos (NO confunde 'K1' con 'K11').
    Devuelve el ID válido si encuentra match, None si no."""
    if not ref or not valid_ids:
        return None

    # 1. Case-insensitive exacto
    ref_lower = ref.lower()
    for vid in valid_ids:
        if vid.lower() == ref_lower:
            return vid

    # 2. Ignorando separadores comunes
    def _normalize(s: str) -> str:
        return s.lower().replace("_", "").replace("-", "").replace(".", "")
    ref_norm = _normalize(ref)
    for vid in valid_ids:
        if _normalize(vid) == ref_norm:
            return vid

    # 3. Levenshtein-aproximado via difflib. cutoff alto para no recuperar
    # falsos positivos (K1 vs K11 vs K12 son distintos blocks legítimos).
    matches = difflib.get_close_matches(ref, valid_ids, n=1, cutoff=0.85)
    if matches:
        return matches[0]
    return None

UGLY_LITERALS = {"nan", "none", "null", "nat", "???", "—", "s/d", "n/a", "na"}
DROP_REASONS = {
    "block_ref_not_found",
    "single_dominant_category",
    "all_zero_values",
    "too_few_categories",
    "too_few_rows",
    "low_fill_ratio",
    "all_bullets_failed_provenance",
    "missing_required_field",
    "no_columns_after_filter",
    "no_rows_after_filter",
}

@dataclass
class ValidationOutcome:
    slides: List[dict] = field(default_factory=list)
    dropped: List[dict] = field(default_factory=list)   # {type, reason, block_ref?}
    bullets_dropped: int = 0


def validate_plan(plan: dict, inventory: List, wb: WorkbookData) -> ValidationOutcome:
    blocks_by_id = {b.id: b for b in inventory}
    out = ValidationOutcome()

    for slide in plan.get("slides", []):
        stype = slide.get("type")
        if stype == "title":
            # title slide always kept (informational, no block ref required)
            out.slides.append(_attach_provenance(slide, None))
            continue

        # Resolve block ref(s)
        ref_keys = []
        if "block_ref" in slide:
            ref_keys.append(slide["block_ref"])
        if "block_refs" in slide:
            ref_keys.extend(slide["block_refs"])
        if "supports_block" in slide:
            ref_keys.append(slide["supports_block"])

        if not ref_keys:
            out.dropped.append({"type": stype, "reason": "missing_required_field"})
            continue

        # Fuzzy recovery: si el AI mandó refs con typo o case wrong,
        # intentamos resolverlos al ID real ANTES de droppear. Esto rescata
        # ~30% de slides que se perdían por mismatch puramente cosmético.
        valid_ids = list(blocks_by_id.keys())
        resolved_keys = []
        unknown = []
        for r in ref_keys:
            if r in blocks_by_id:
                resolved_keys.append(r)
            else:
                fixed = _fuzzy_resolve_block_ref(r, valid_ids)
                if fixed:
                    resolved_keys.append(fixed)
                else:
                    unknown.append(r)

        if unknown:
            out.dropped.append({"type": stype, "reason": "block_ref_not_found",
                                  "block_refs": unknown})
            continue

        # Reescribir el slide con los IDs resueltos. Mantiene la estructura
        # original (block_ref vs block_refs vs supports_block) pero con los
        # IDs canónicos. Sin esto, el extractor aguas abajo recibe los IDs
        # con typo y vuelve a fallar.
        if resolved_keys != ref_keys:
            slide = dict(slide)
            idx = 0
            if "block_ref" in slide:
                slide["block_ref"] = resolved_keys[idx]; idx += 1
            if "block_refs" in slide:
                n = len(slide["block_refs"])
                slide["block_refs"] = resolved_keys[idx:idx + n]; idx += n
            if "supports_block" in slide:
                slide["supports_block"] = resolved_keys[idx]; idx += 1

        ref_blocks = [blocks_by_id[r] for r in resolved_keys]

        if stype == "chart":
            primary = ref_blocks[0]
            bad_flag = next((f for f in primary.quality_flags
                              if f in {"single_dominant_category", "all_zero",
                                        "low_fill_ratio", "too_few_rows"}), None)
            if bad_flag:
                out.dropped.append({"type": stype, "reason": bad_flag,
                                      "block_ref": primary.id})
                continue

        if stype == "table":
            primary = ref_blocks[0]
            if "low_fill_ratio" in primary.quality_flags:
                out.dropped.append({"type": stype, "reason": "low_fill_ratio",
                                      "block_ref": primary.id})
                continue

        if stype == "text_bullets":
            block = ref_blocks[0]
            kept_bullets = []
            for bullet in slide.get("bullets", []):
                if not _bullet_has_provenance(bullet, block, wb):
                    out.bullets_dropped += 1
                    continue
                if _bullet_misleads_on_status(bullet, inventory):
                    # Bullet cites a raw-sum value of a column with a status
                    # filter, but never mentions the breakdown — would
                    # parrot "$1.5B total" as money moved when $258M of
                    # that was rejected. Drop it.
                    out.bullets_dropped += 1
                    continue
                if _bullet_misattributes_extreme(bullet, block, wb):
                    # Bullet cites an extreme value (max/min) attributed to
                    # a person/city/code that does NOT belong to the same
                    # row in the source df. Classic cross-block hallucination:
                    # the LLM combines a real max value from one column with
                    # the top label of an unrelated categorical column.
                    out.bullets_dropped += 1
                    continue
                kept_bullets.append(_clean_bullet(bullet))
            if not kept_bullets:
                out.dropped.append({"type": stype,
                                      "reason": "all_bullets_failed_provenance",
                                      "block_ref": block.id})
                continue
            slide = {**slide, "bullets": kept_bullets}

        out.slides.append(_attach_provenance(slide, ref_blocks[0] if ref_blocks else None))

    return out


def _attach_provenance(slide: dict, block) -> dict:
    if block is None:
        return slide
    return {**slide, "provenance": {
        "sheet": block.provenance.sheet,
        "columns": block.provenance.columns,
        "rows": list(block.provenance.rows),
    }}


_NUM_RE = re.compile(r"-?\d+(?:[\.,]\d+)?")
_NAME_RE = re.compile(r"\b[A-ZÁÉÍÓÚÑ][a-záéíóúñ]{2,}\b")

def _bullet_has_provenance(bullet: str, block, wb: WorkbookData) -> bool:
    """Returns True if every number / proper name in the bullet exists in the
    block. El haystack incluye samples + top_values + stats agregadas
    (sum/mean/min/max), porque el AI suele citar totales y promedios que
    no están en samples crudos. Tolerancia numérica 5% — lo suficiente
    para narrativas tipo '$1.5M' contra valor real $1,523,456.
    """
    if not bullet or not bullet.strip():
        return False
    text = str(bullet).strip()
    nums = _NUM_RE.findall(text)
    names = _NAME_RE.findall(text)
    if not nums and not names:
        # Pure-qualitative bullet → not allowed (would be generic)
        return False

    # Get the actual data slice for this block
    sheet = next((s for s in wb.sheets if s.name == block.provenance.sheet), None)
    if not sheet:
        return False
    cols_set = set(block.provenance.columns)

    # Build a flat searchable string of relevant cells
    haystack_str = ""
    haystack_nums: set = set()

    # Nombres de columnas + sheet — bullets como "El total de 'Monto Solicitado'"
    # citan nombres del schema, no del dato. Sin esto cualquier bullet que
    # mencione una columna explícitamente fallaba provenance aunque las cifras
    # citadas SÍ estuvieran. Incluyo también el sheet name por si la narrativa
    # cita la hoja.
    for col_name in cols_set:
        haystack_str += " " + str(col_name).lower()
    haystack_str += " " + str(block.provenance.sheet).lower()

    for col in sheet.columns:
        if col.name not in cols_set:
            continue
        # Samples crudos
        for s in col.samples:
            if s is None:
                continue
            if isinstance(s, (int, float)):
                haystack_nums.add(round(float(s), 2))
            else:
                haystack_str += " " + str(s).lower()
        # Top categóricos (con conteos también — el AI cita "X tiene 47 casos")
        for tv in (col.top_values or []):
            haystack_str += " " + str(tv[0]).lower()
            if len(tv) > 1 and isinstance(tv[1], (int, float)):
                haystack_nums.add(round(float(tv[1]), 2))
        # Stats agregadas — fundamental para bullets de narrativa.
        # El AI dice "el total es $5M" y esto vive en col.sum, no en samples.
        for stat in (col.sum, col.mean, col.min, col.max):
            if stat is not None and isinstance(stat, (int, float)):
                try:
                    haystack_nums.add(round(float(stat), 2))
                except (ValueError, OverflowError):
                    pass
    # Agregamos también el rango total de la hoja (filas) — el AI cita
    # cosas como "se procesaron 200 registros" donde 200 = sheet rows.
    rows_n = sheet.shape[0] if sheet.shape else 0
    if rows_n:
        haystack_nums.add(float(rows_n))

    # Match cada número del bullet con tolerancia 5% (suficiente para
    # narrativas con sufijos K/M/B redondeados). Para magnitudes pequeñas
    # (<100) requerimos exact match.
    # _parse_narrative_number puede devolver MÚLTIPLES candidatos cuando el
    # mismo token aparece varias veces con sufijos distintos en el bullet
    # (ej: "1.5% generó 1.5M" → [1.5, 1500000]). Aceptamos si CUALQUIER
    # candidato matchea — sino dropearíamos bullets válidos.
    for n in nums:
        try:
            n_candidates = _parse_narrative_number(n, text)
        except ValueError:
            continue
        if not n_candidates:
            continue
        matched = False
        for n_val in n_candidates:
            tolerance = (max(1.0, abs(n_val) * 0.05) if abs(n_val) >= 100
                           else 0.5)
            if any(abs(n_val - h) <= tolerance for h in haystack_nums):
                matched = True
                break
        if not matched:
            return False

    for name in names:
        if name.lower() not in haystack_str:
            return False

    return True


def _parse_narrative_number(token: str, full_text: str):
    """Convierte un token numérico considerando sufijos K/M/B/MM cercanos.

    Importante: devuelve TODOS los valores candidatos cuando el token aparece
    múltiples veces (ej: 'El crecimiento del 1.5% generó 1.5M en ventas' →
    [1.5, 1500000] porque el primer 1.5 va sin sufijo y el segundo con M).
    Si solo aparece una vez, devuelve [valor_único]. Antes esto solo miraba
    la PRIMERA ocurrencia y dropeaba bullets cuyo número aparecía dos veces
    con distintos sufijos.

    Devuelve [] si el token no se puede parsear como float, o lista de floats
    (con al menos 1 elemento) en cualquier otro caso.
    """
    try:
        base = float(token.replace(",", "."))
    except ValueError:
        return []
    candidates: list[float] = []
    start = 0
    tlen = len(token)
    while True:
        pos = full_text.find(token, start)
        if pos < 0:
            break
        suffix_area = full_text[pos + tlen: pos + tlen + 3].upper().strip()
        if suffix_area.startswith("MM"):  # millones (ES algunos contextos)
            candidates.append(base * 1_000_000)
        elif suffix_area.startswith("M"):
            candidates.append(base * 1_000_000)
        elif suffix_area.startswith("K") or suffix_area.startswith("MIL"):
            candidates.append(base * 1_000)
        elif suffix_area.startswith("B"):
            candidates.append(base * 1_000_000_000)
        else:
            candidates.append(base)
        start = pos + tlen
    return candidates if candidates else [base]


def _clean_bullet(text: str) -> str:
    s = str(text).strip()
    for ugly in UGLY_LITERALS:
        s = re.sub(rf"\b{re.escape(ugly)}\b", "", s, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", s).strip()


# Keywords that signal the bullet correctly contextualises a raw-sum number
# (so the reader understands it includes rejected/anulated rows).
_DISCLAIMER_KEYWORDS = (
    "bruto", "incluye", "desglose", "rechaz", "anulad", "cancel",
    "antes de", "previo a", "denegad", "estado", "filtrad",
)


_EXTREME_KEYWORDS = ("máximo", "maximo", "mínimo", "minimo", "mayor",
                      "menor", "más alto", "mas alto", "más bajo", "mas bajo",
                      "top", "extremo", "outlier", "outliers")


def _bullet_misattributes_extreme(bullet: str, block, wb) -> bool:
    """Drop bullets that say "X is the max/min" and attribute it to a
    person/city/centro that doesn't actually appear in the row holding
    the extreme value.

    Logic:
      1. Bullet must contain an extreme keyword (máximo, mínimo, …) AND
         a money token whose value is close (±5%) to the column.max or
         column.min of some numeric column on this block's sheet.
      2. Find the row that holds that extreme.
      3. Collect "attribution candidates" cited in the bullet — capitalised
         strings of 4+ chars that look like a proper noun (NOT digits, NOT
         the column name, NOT a known UI token).
      4. Each cited candidate must appear in the extreme row's
         attribution-bearing columns (Solicitante, Ciudad, Estado, …). If
         even one cited candidate is absent from the extreme row, the
         bullet is dropped — that's the cross-block hallucination.

    Best-effort: any failure (df not loadable, ambiguous extreme, etc.)
    returns False so we don't drop legitimate bullets on bad inputs.
    """
    if not bullet or not block:
        return False
    text = str(bullet)
    text_l = text.lower()
    if not any(k in text_l for k in _EXTREME_KEYWORDS):
        return False
    try:
        import pandas as pd
        from socya_pipeline.insights import NUMBER_TOKEN_RE
        sheet_name = block.provenance.sheet
        if not getattr(wb, "filename", None):
            return False
        # Try header=0; if more than 60% of columns are Unnamed, retry
        # header=1. The user's sheets often have a sheet-title row above
        # the real headers ("COMISIONES" in row 0, real columns in row 1).
        df = None
        try:
            df = pd.read_excel(wb.filename, sheet_name=sheet_name, header=0)
            unnamed = sum(1 for c in df.columns if str(c).startswith("Unnamed"))
            if unnamed / max(1, len(df.columns)) > 0.6:
                df = pd.read_excel(wb.filename, sheet_name=sheet_name, header=1)
        except Exception:
            return False
        if df is None or df.empty:
            return False

        # Step 1: find the numeric column whose extreme matches a value in
        # the bullet (±5%). If none matches, the bullet isn't actually
        # citing a column extreme — let it pass.
        sheet_obj = next((s for s in wb.sheets if s.name == sheet_name), None)
        if sheet_obj is None:
            return False
        numeric_cols = [c for c in sheet_obj.columns
                        if c.dtype in ("currency", "numeric")
                        and (c.max is not None or c.min is not None)]
        # Use a thousands-aware regex (NUMBER_TOKEN_RE handles "8,082,100"
        # as a single token). We also handle SPACE as a thousands separator
        # ("8 082 100" — what LLMs sometimes emit per Spanish-style typography).
        bullet_numbers = []
        # Space-separated number groups (most aggressive interpretation first)
        for m in re.finditer(r"\b\d{1,3}(?:[  ]\d{3})+\b", text):
            try:
                bullet_numbers.append(float(m.group(0).replace(" ", "").replace(" ", "")))
            except ValueError:
                pass
        for raw in NUMBER_TOKEN_RE.findall(text):
            try:
                bullet_numbers.append(float(raw.replace(",", "")))
            except ValueError:
                pass
            try:
                bullet_numbers.extend(_parse_narrative_number(raw, text))
            except Exception:
                pass
        target_col = None
        target_extreme = None  # "max" or "min"
        for col in numeric_cols:
            for ext_name, ext_val in (("max", col.max), ("min", col.min)):
                if ext_val is None:
                    continue
                for n in bullet_numbers:
                    if ext_val == 0:
                        continue
                    if abs(n - float(ext_val)) / abs(float(ext_val)) <= 0.05:
                        target_col = col
                        target_extreme = ext_name
                        break
                if target_col:
                    break
            if target_col:
                break
        if target_col is None:
            return False
        # Resolve the column name to one that actually exists in the df.
        # The inventory may have prefixed it with a sheet title ("COMISIONES > X").
        df_col_name = _resolve_df_column(target_col.name, df.columns)
        if df_col_name is None:
            return False

        # Step 2: find the row with the extreme.
        series = pd.to_numeric(df[df_col_name], errors="coerce")
        try:
            idx = series.idxmax() if target_extreme == "max" else series.idxmin()
        except Exception:
            return False
        if idx not in df.index:
            return False

        # Step 3: collect attribution candidates cited in the bullet.
        # Heuristic patterns:
        #   - ALL-CAPS city pairs joined by "-" (ANTIOQUIA - EL BAGRE)
        #   - ALL-CAPS multi-word names without "-" (JORGE ENRIQUE ANGEL POVEDA)
        #   - Capitalised multi-word names (Miguel Angel Gomez Torres)
        candidates = set()
        for m in re.finditer(r"[A-ZÁÉÍÓÚÑ]{4,}(?:\s*-\s*[A-ZÁÉÍÓÚÑ ]{3,})+", text):
            candidates.add(_norm(m.group(0)))
        for m in re.finditer(
            r"\b[A-ZÁÉÍÓÚÑ]{2,}(?:\s+[A-ZÁÉÍÓÚÑ]{2,}){1,5}\b", text):
            candidates.add(_norm(m.group(0)))
        for m in re.finditer(
            r"\b[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,5}\b", text):
            candidates.add(_norm(m.group(0)))
        # Strip UI tokens / column names from candidates
        ui_blacklist = {"valor total solicitado", "comisiones", "indicadores clave",
                          "hallazgos", "análisis", "analisis"}
        col_names_l = {_norm(c.name) for c in sheet_obj.columns}
        candidates = {c for c in candidates
                       if c not in ui_blacklist and c not in col_names_l}
        if not candidates:
            return False

        # Step 4: collect attribution from the extreme row.
        row_attr = set()
        for col in df.columns:
            name_l = str(col).strip().lower()
            if not any(h in name_l for h in
                        ("solicitante", "responsable", "encargado", "ciudad",
                         "destino", "estado", "status", "centro de costos",
                         "area", "área", "proyecto", "categor", "tipo")):
                continue
            try:
                v = df.at[idx, col]
            except Exception:
                continue
            if v is None:
                continue
            sv = str(v).strip()
            if sv and sv.lower() not in ("nan", "none", "null"):
                row_attr.add(_norm(sv))

        # If a cited candidate is NOT in the row's attribution set (even by
        # substring match) → misattribution.
        for cand in candidates:
            hit = False
            for actual in row_attr:
                if cand in actual or actual in cand:
                    hit = True
                    break
            if not hit:
                return True
        return False
    except Exception:
        return False


def _norm(s: str) -> str:
    import unicodedata
    s2 = "".join(c for c in unicodedata.normalize("NFD", str(s))
                 if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"\s+", " ", s2).strip()


def _resolve_df_column(canonical_name: str, df_columns) -> str | None:
    """Match a name like "COMISIONES > Valor Total Solicitado" to the actual
    df header "Valor Total Solicitado". The inventory may prefix column
    names with the sheet's title row when it detects one; the raw df keeps
    them unprefixed. Falls back to normalized substring match."""
    if canonical_name in df_columns:
        return canonical_name
    # Strip "X > " prefix if present
    if " > " in canonical_name:
        suffix = canonical_name.split(" > ", 1)[1]
        if suffix in df_columns:
            return suffix
    canonical_n = _norm(canonical_name)
    for c in df_columns:
        if _norm(c) == canonical_n:
            return c
    # Last resort: substring (handles trailing whitespace, " > X" prefix variants)
    for c in df_columns:
        cn = _norm(c)
        if cn and (cn in canonical_n or canonical_n in cn):
            return c
    return None


def _bullet_misleads_on_status(bullet: str, inventory) -> bool:
    """Return True when the bullet cites a numeric value that matches the
    raw-sum `value` of an inventory KPI with the `needs_status_disclaimer`
    flag AND the bullet does not mention any disclaimer keyword (`bruto`,
    `incluye`, `rechaz`, `solicitad`, …). That is the precise failure
    mode the user reported — "VALOR TOTAL SOLICITADO $1.5B" with no
    indication that $258M of it was rejected.
    """
    if not bullet or not inventory:
        return False
    text = str(bullet)
    lower = text.lower()
    if any(k in lower for k in _DISCLAIMER_KEYWORDS):
        return False  # bullet already discloses the breakdown context
    raw_nums = _NUM_RE.findall(text)
    if not raw_nums:
        return False
    # Parse each number with K/M/B suffix awareness — `_parse_narrative_number`
    # returns all plausible interpretations for the token in this text.
    candidates = []
    for raw in raw_nums:
        try:
            candidates.extend(_parse_narrative_number(raw, text))
        except Exception:
            try:
                candidates.append(float(raw.replace(",", ".")))
            except ValueError:
                continue
    if not candidates:
        return False
    # Compare against every KPI flagged for disclaimer; 5% tolerance is the
    # same numeric tolerance used in `_bullet_has_provenance`.
    tolerance = 0.05
    for b in inventory:
        if b.kind != "kpi_candidate":
            continue
        if "needs_status_disclaimer" not in b.quality_flags:
            continue
        target = b.extra.get("value")
        if target is None:
            continue
        try:
            t = float(target)
        except (TypeError, ValueError):
            continue
        if t == 0:
            continue
        for c in candidates:
            if abs(c - t) / abs(t) <= tolerance:
                return True
    return False
