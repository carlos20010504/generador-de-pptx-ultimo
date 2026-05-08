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
    """Returns True if every number / proper name in the bullet exists in the block."""
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
    haystack_nums = set()
    for col in sheet.columns:
        if col.name not in cols_set:
            continue
        for s in col.samples:
            if s is None:
                continue
            if isinstance(s, (int, float)):
                haystack_nums.add(round(float(s), 2))
            else:
                haystack_str += " " + str(s).lower()
        for tv in (col.top_values or []):
            haystack_str += " " + str(tv[0]).lower()

    for n in nums:
        try:
            n_val = float(n.replace(",", "."))
            if not any(abs(n_val - h) <= max(1.0, abs(h) * 0.005) for h in haystack_nums):
                return False
        except ValueError:
            continue

    text_lower = text.lower()
    for name in names:
        if name.lower() not in haystack_str:
            return False

    return True


def _clean_bullet(text: str) -> str:
    s = str(text).strip()
    for ugly in UGLY_LITERALS:
        s = re.sub(rf"\b{re.escape(ugly)}\b", "", s, flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", s).strip()
