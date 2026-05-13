"""Strict validator: drops slides without traceable provenance or low quality."""
from dataclasses import dataclass, field
from typing import List, Dict
import re
from socya_pipeline.parser import WorkbookData

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

        unknown = [r for r in ref_keys if r not in blocks_by_id]
        if unknown:
            out.dropped.append({"type": stype, "reason": "block_ref_not_found",
                                  "block_refs": unknown})
            continue

        ref_blocks = [blocks_by_id[r] for r in ref_keys]

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
                if _bullet_has_provenance(bullet, block, wb):
                    kept_bullets.append(_clean_bullet(bullet))
                else:
                    out.bullets_dropped += 1
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
