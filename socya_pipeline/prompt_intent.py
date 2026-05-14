"""Parsea la intención del usuario expresada en español.

Pure functions, no I/O. Used by planner.py to:
1. Inject structured signals into the AI payload.
2. Enforce post-AI: trim/pad slides to match `requested_slide_count`,
   inject missing `required_sheets` slides.

The parsing is deliberately conservative — when in doubt, return None and
let the AI decide. False positives are worse than no-ops because they
override the user's natural-language intent silently.
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import List, Optional


_SLIDE_COUNT_PATTERNS = (
    re.compile(r"(\d{1,3})\s*(?:slides?|diapositivas?|láminas?|laminas?)",
               re.IGNORECASE),
    re.compile(r"deck\s+de\s+(\d{1,3})", re.IGNORECASE),
)


@dataclass
class SheetMatch:
    """Una hoja que el user mencionó en el prompt + el matching real."""
    requested: str
    matched: Optional[str]
    closest: Optional[str]
    ratio: float


@dataclass
class PromptIntent:
    """Resultado del parsing — todo opcional, valores `None`/`[]` ⇒ sin señal."""
    requested_slide_count: Optional[int] = None
    required_sheets: List[SheetMatch] = field(default_factory=list)
    topic_hints: List[str] = field(default_factory=list)


def extract(user_prompt: str, available_sheet_names: List[str]) -> PromptIntent:
    """Parsea `user_prompt` y devuelve un PromptIntent estructurado."""
    if not user_prompt or not user_prompt.strip():
        return PromptIntent()
    return PromptIntent(
        requested_slide_count=_extract_slide_count(user_prompt),
        required_sheets=[],   # Task 2
        topic_hints=[],       # Task 3
    )


def _extract_slide_count(prompt: str) -> Optional[int]:
    """Devuelve el primer slide-count válido (≥1, ≤999) o None."""
    for pattern in _SLIDE_COUNT_PATTERNS:
        match = pattern.search(prompt)
        if match:
            try:
                n = int(match.group(1))
            except ValueError:
                continue
            if 1 <= n <= 999:
                return n
    return None
