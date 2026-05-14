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

_FUZZY_THRESHOLD = 0.65

_STOPWORDS = frozenset({
    "todo", "todos", "toda", "todas", "el", "la", "los", "las",
    "y", "o", "u", "de", "del", "en", "con", "para", "por",
    "datos", "información", "informacion", "info", "hoja", "hojas",
    "tabla", "tablas", "slide", "slides", "diapositiva", "diapositivas",
    "lámina", "láminas", "lamina", "laminas", "deck", "presentación",
    "presentacion", "powerpoint", "ppt", "pptx",
    "más", "mas", "menos", "que", "como", "asi", "así",
})


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
        required_sheets=_extract_required_sheets(user_prompt,
                                                   available_sheet_names),
        topic_hints=[],   # Task 3
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


def _normalize(text: str) -> str:
    """lower + strip accents + collapse whitespace + strip punctuation
    excepto espacios y guiones."""
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", text)
    no_accents = "".join(c for c in nfkd if not unicodedata.combining(c))
    out = no_accents.lower().strip()
    out = re.sub(r"[_\.,;:!?\(\)\[\]\{\}]+", " ", out)
    out = re.sub(r"\s+", " ", out).strip()
    return out


def _ratio(a: str, b: str) -> float:
    """SequenceMatcher ratio sobre versiones normalizadas."""
    return SequenceMatcher(None, _normalize(a), _normalize(b)).ratio()


def _shares_token_prefix(ng: str, sheet: str, min_prefix: int = 4) -> bool:
    """Devuelve True si algún token del ngram comparte un prefijo de
    `min_prefix` caracteres con algún token de la sheet (normalizadas).

    Esto evita falsos positivos como `proyecciones` ↔ `acciones` (alto
    ratio por sufijo común) y exige solapamiento al inicio de palabra.
    """
    ng_tokens = [t for t in _normalize(ng).split() if t]
    sh_tokens = [t for t in re.split(r"[\s\-_]+", _normalize(sheet)) if t]
    for nt in ng_tokens:
        for st in sh_tokens:
            if len(nt) < min_prefix or len(st) < min_prefix:
                # tokens cortos: requiere igualdad completa
                if nt == st:
                    return True
                continue
            # comparte ≥min_prefix caracteres iniciales en alguna dirección
            if nt[:min_prefix] == st[:min_prefix]:
                return True
            if nt.startswith(st) or st.startswith(nt):
                return True
    return False


def _extract_required_sheets(prompt: str,
                              available: List[str]) -> List[SheetMatch]:
    """Encuentra menciones a hojas en el prompt usando fuzzy match.

    Estrategia conservadora:
    1. Genera n-gramas (1..4 tokens) del prompt normalizado.
    2. Para cada hoja, busca el mejor n-grama por ratio.
    3. Acepta como match si ratio ≥ _FUZZY_THRESHOLD Y comparte prefijo
       de token con la hoja (evita falsos positivos por sufijos comunes).
    4. Si no hubo NINGÚN match, devuelve a lo sumo 1 entrada `closest`
       (el mejor candidato razonable, ratio ≥ 0.45 y con prefijo).
    """
    if not available:
        return []

    norm_prompt = _normalize(prompt)
    if not norm_prompt:
        return []

    tokens = norm_prompt.split()
    ngrams: List[str] = []
    for n in (4, 3, 2, 1):
        for i in range(len(tokens) - n + 1):
            ng = " ".join(tokens[i:i+n])
            if all(t in _STOPWORDS for t in tokens[i:i+n]):
                continue
            if n == 1 and (len(ng) < 4 or ng in _STOPWORDS):
                continue
            ngrams.append(ng)

    if not ngrams:
        return []

    matches: List[SheetMatch] = []
    matched_sheets: set = set()
    for sheet in available:
        best_ratio = 0.0
        best_ng = ""
        for ng in ngrams:
            r = _ratio(ng, sheet)
            if r > best_ratio:
                best_ratio = r
                best_ng = ng
        if best_ratio >= _FUZZY_THRESHOLD and _shares_token_prefix(best_ng, sheet):
            matches.append(SheetMatch(
                requested=best_ng,
                matched=sheet,
                closest=None,
                ratio=round(best_ratio, 3),
            ))
            matched_sheets.add(sheet)

    # Closest-fallback solo si NO hubo matches: a lo sumo 1 entrada.
    # Heurística: el usuario escribió una palabra "nombre-ish" (que aparece
    # capitalizada o en MAYÚSCULAS en el prompt original) — interpretamos
    # eso como intención de referirse a una hoja, aunque no exista.
    if not matches:
        # Tokens del prompt original que están capitalizados (probables
        # nombres propios de hoja). Normalizamos para comparar.
        capitalized_norm_tokens: set = set()
        for raw_tok in re.findall(r"[\wáéíóúñÁÉÍÓÚÑ]+", prompt):
            if (raw_tok[:1].isupper() or raw_tok.isupper()) and len(raw_tok) >= 4:
                capitalized_norm_tokens.add(_normalize(raw_tok))

        def _has_capitalized_token(ng: str) -> bool:
            return any(t in capitalized_norm_tokens for t in ng.split())

        candidate_ngrams = [
            ng for ng in ngrams
            if ((" " in ng) or (len(ng) >= 6)) and _has_capitalized_token(ng)
        ]

        # Pares (ng, sheet) con su ratio, descartando los falsos positivos
        # (ratio ≥ threshold pero sin solapamiento de prefijo de token).
        scored: List[tuple] = []
        for ng in candidate_ngrams:
            for sheet in available:
                r = _ratio(ng, sheet)
                if r >= _FUZZY_THRESHOLD and not _shares_token_prefix(ng, sheet):
                    # Falso positivo por sufijo común — descartar este par.
                    continue
                scored.append((r, ng, sheet))

        # Mejor candidato bajo el threshold
        scored = [s for s in scored if s[0] < _FUZZY_THRESHOLD]
        if scored:
            scored.sort(key=lambda x: x[0], reverse=True)
            best_overall_ratio, best_overall_ng, best_overall_sheet = scored[0]
            matches.append(SheetMatch(
                requested=best_overall_ng,
                matched=None,
                closest=best_overall_sheet,
                ratio=round(best_overall_ratio, 3),
            ))

    seen_sheets: dict = {}
    for m in matches:
        key = m.matched or f"NONE_{m.requested}"
        if key not in seen_sheets or m.ratio > seen_sheets[key].ratio:
            seen_sheets[key] = m
    return list(seen_sheets.values())
