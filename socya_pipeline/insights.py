"""Statistical / semantic insight helpers — pure functions, no I/O.

This module is the "brain layer" on top of the existing parser/inventory:
nothing else depends on it for correctness. If a function here raises or
returns None, the rest of the pipeline keeps working with its existing
defaults. Each helper is independently testable.

Capabilities
- Outlier detection (z-score on numeric series)
- Pareto concentration (top-N share of total)
- Seasonality hints for monthly / quarterly time series
- Semantic column detection (geography / person / id-coded / boolean-disguised)
- PII detection + masking (cédulas / emails / phones)
- Smart KPI subsumption (Total General ≈ sum of sub-totals)
- KPI ratio pairing (Solicitado / Aprobado → % aprobación)
- Auto-prompt inference from workbook structure
- Shared formatters (compact K/M/B numbers, humanized headers, JSON loose parse)
"""
from __future__ import annotations

import json
import math
import re
from typing import Iterable, List, Optional, Tuple

# ─────────────────────────────────────────────────────────────────────
# Shared constants (avoid duplicating these across modules)
# ─────────────────────────────────────────────────────────────────────

# "Empty / placeholder" cell values that should be treated as null when
# rendering tables, validating bullets, or formatting outputs.
UGLY_LITERALS: frozenset = frozenset({
    "nan", "none", "null", "nat", "???", "—", "s/d", "n/a", "na",
})

# Token used to find numbers in narrative / bullet text for anti-hallucination
# checks. Matches both plain integers and thousands-separated forms.
NUMBER_TOKEN_RE = re.compile(r"\d{1,3}(?:[\.,]\d{3})*(?:[\.,]\d+)?|\d+")

# Tokens whose presence in a column NAME signals an identifier (not a value).
ID_NAME_TOKENS: tuple = (
    "id", "código", "codigo", "número", "numero", "n°", "no.",
    "folio", "consecutivo", "uuid", "key",
)

# Status-column detection. Used to spot columns like "Estado" / "Status"
# that filter a money column into ejecutado vs rechazado. Without this,
# a KPI like "Total Valor Solicitado = $1.5B" misleads the reader because
# it includes rows that were rejected and never moved money.
STATUS_COLUMN_NAME_TOKENS: tuple = (
    "estado", "status", "situaci", "resultado", "decisi",
    "etapa", "fase", "tramite", "trámite", "condicion", "condición",
)

# Substrings (uppercased) that indicate a row's value was NOT executed.
# Match is case-insensitive substring (matches "RECHAZADO" and
# "RECHAZADO DESEMBOLSO"). Keep the stem short to catch conjugations.
NEGATIVE_STATUS_TOKENS: tuple = (
    "RECHAZAD", "ANULAD", "CANCELAD", "DENEGAD", "INVALID",
    "ELIMINAD", "BORRAD", "DESCARTAD", "NO APROB", "DESAPROB",
    "DEVUELT", "REVERSAD", "INADMI",
)

# Substrings that indicate a row IS executed (real money moved / committed).
POSITIVE_STATUS_TOKENS: tuple = (
    "CONTABILIZAD", "DESEMBOLSAD", "EJECUTAD", "LEGALIZAD",
    "PAGAD", "REALIZAD", "FINALIZAD", "COMPLETAD", "FACTURAD",
    "CERRAD", "LIQUIDAD",
)

# Substrings that indicate a row is in-flight (not rejected, not executed).
PENDING_STATUS_TOKENS: tuple = (
    "PENDIENT", "EN ESPER", "EN PROCES", "EN AJUST", "EN CONTABIL",
    "EN REVIS", "APROBADO LIDER", "APROBAD", "SOLICITAD",
    "EN TRAMITE", "EN TRÁMITE", "EN CURSO", "ABIERT",
)


# ─────────────────────────────────────────────────────────────────────
# Shared formatters
# ─────────────────────────────────────────────────────────────────────

def format_compact(value, *, prefix: str = "", big_decimals: int = 0,
                     mid_decimals: int = 1, fallback: str = "—") -> str:
    """Format a number as a compact, human-readable string with optional unit
    prefix. Single source of truth for K/M formatting across the pipeline.

    We deliberately **do NOT use "B"** — in Spanish "billón" is 10^12, while
    the American "B" is 10^9. A KPI label "1.5B pesos" reads as
    "$1,500,000,000,000" to a Spanish reader (a thousand times larger). All
    values ≥ 1B are expressed in millions with a thousands separator
    (e.g. "$1,504M"), matching LatAm financial-press convention.

    Examples:
      format_compact(1500)              → "1.5K"
      format_compact(1_500_000)         → "1.5M"
      format_compact(1_500_000_000)     → "1,500M"
      format_compact(1_504_462_283)     → "1,504M"
      format_compact(1500, prefix="$")  → "$1.5K"
      format_compact(None)              → "—"
      format_compact(float("nan"))      → "—"
    """
    if value is None:
        return fallback
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if not math.isfinite(f):
        return fallback
    sign = "-" if f < 0 else ""
    a = abs(f)
    if a >= 1_000_000:
        millions = a / 1_000_000
        # ≥ 1,000M: drop the decimal point — "1,504M" is cleaner than
        # "1504.5M" and avoids the ambiguous "B" suffix in Spanish.
        if millions >= 1000:
            return f"{sign}{prefix}{millions:,.{big_decimals}f}M"
        if millions == int(millions):
            return f"{sign}{prefix}{int(millions)}M"
        return f"{sign}{prefix}{millions:.{mid_decimals}f}M"
    if a >= 1_000:
        k = a / 1_000
        if k == int(k):
            return f"{sign}{prefix}{int(k)}K"
        return f"{sign}{prefix}{k:.{mid_decimals}f}K"
    if f.is_integer():
        return f"{sign}{prefix}{int(a)}"
    return f"{sign}{prefix}{a:.{mid_decimals}f}"


def humanize_header(value) -> str:
    """Convert raw column headers into reader-friendly form.

    - Strips Unnamed:/nan/none placeholders → returns ""
    - snake_case → "Snake Case"
    - camelCase → "Camel Case"
    - Preserves all-caps acronyms (≥3 chars)
    """
    s = str(value or "").strip()
    if not s:
        return ""
    if s.startswith("Unnamed:") or s.lower() in ("nan", "none", "null"):
        return ""
    # snake / kebab → spaces
    s = s.replace("_", " ").replace("-", " ")
    # camelCase split: insert space before each uppercase that follows lowercase
    s = re.sub(r"(?<=[a-záéíóúñ])(?=[A-ZÁÉÍÓÚÑ])", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    parts = []
    for w in s.split(" "):
        if w.isupper() and len(w) >= 3:
            parts.append(w)
        else:
            parts.append(w.capitalize() if w else w)
    return " ".join(parts)


def parse_loose_json(text: str) -> Optional[dict]:
    """Extract a JSON object from a model response that may include markdown
    fences or surrounding prose. Returns None if no valid object found."""
    raw = (text or "").strip()
    if not raw:
        return None
    raw = raw.replace("```json", "").replace("```", "").strip()
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if m:
        raw = m.group(0)
    try:
        result = json.loads(raw)
        return result if isinstance(result, dict) else None
    except (json.JSONDecodeError, ValueError):
        return None


def is_id_like_name(name: str) -> bool:
    """Return True when the column NAME signals an identifier (id, código,
    folio, etc.). Does not consider data — see callers for cardinality checks."""
    n = (name or "").strip().lower()
    if not n:
        return False
    for token in ID_NAME_TOKENS:
        if (n == token
                or n.startswith(token + " ")
                or n.endswith(" " + token)
                or f"_{token}" in n
                or f" {token} " in f" {n} "):
            return True
    return False


# ─────────────────────────────────────────────────────────────────────
# Outliers & Pareto
# ─────────────────────────────────────────────────────────────────────

def detect_outliers(values: Iterable[float], z_threshold: float = 2.5
                     ) -> List[Tuple[int, float, float]]:
    """Return [(index, value, z_score)] for values whose |z-score| ≥ threshold.
    Uses sample std; needs ≥4 points to be meaningful."""
    nums = [float(v) for v in values if v is not None and _isfinite(v)]
    if len(nums) < 4:
        return []
    mean = sum(nums) / len(nums)
    var = sum((v - mean) ** 2 for v in nums) / (len(nums) - 1)
    sd = math.sqrt(var)
    if sd == 0:
        return []
    out = []
    for i, v in enumerate(nums):
        z = (v - mean) / sd
        if abs(z) >= z_threshold:
            out.append((i, v, z))
    return out


def pareto_share(values: Iterable[float], top_n: int = 3) -> Optional[float]:
    """Return the share of total represented by the top N values (0.0 – 1.0).
    Returns None if there's no data or total is 0."""
    nums = sorted([float(v) for v in values if v is not None and _isfinite(v)],
                  reverse=True)
    total = sum(nums)
    if not nums or total <= 0:
        return None
    return sum(nums[:top_n]) / total


def iqr_outliers(values: Iterable[float], k: float = 1.5
                  ) -> List[Tuple[int, float, str]]:
    """IQR-based outlier detection. More robust than z-score for skewed data
    (z-score assumes normal distribution; financial / categorical data is
    rarely normal). Returns [(index, value, side)] where side is 'high' or
    'low'. Needs ≥5 points for a meaningful Q1/Q3."""
    nums = [(i, float(v)) for i, v in enumerate(values)
              if v is not None and _isfinite(v)]
    if len(nums) < 5:
        return []
    sorted_vals = sorted(v for _, v in nums)
    n = len(sorted_vals)
    # Linear interpolation quartiles (Tukey's method)
    def _quantile(arr, q):
        pos = (len(arr) - 1) * q
        lo = int(pos)
        hi = min(lo + 1, len(arr) - 1)
        frac = pos - lo
        return arr[lo] * (1 - frac) + arr[hi] * frac
    q1 = _quantile(sorted_vals, 0.25)
    q3 = _quantile(sorted_vals, 0.75)
    iqr = q3 - q1
    if iqr <= 0:
        return []
    low_fence = q1 - k * iqr
    high_fence = q3 + k * iqr
    out = []
    for i, v in nums:
        if v > high_fence:
            out.append((i, v, "high"))
        elif v < low_fence:
            out.append((i, v, "low"))
    return out


def trend_acceleration(values: List[float]) -> Optional[str]:
    """For a time series with ≥4 points, return 'accelerating',
    'decelerating', or None if change is roughly linear / noisy.
    Compares the slope of the second half vs the first half."""
    nums = [float(v) for v in values if v is not None and _isfinite(v)]
    if len(nums) < 4:
        return None
    mid = len(nums) // 2
    first_half = nums[:mid]
    second_half = nums[mid:]
    if not first_half or not second_half:
        return None
    fh_slope = (first_half[-1] - first_half[0]) / max(1, len(first_half) - 1)
    sh_slope = (second_half[-1] - second_half[0]) / max(1, len(second_half) - 1)
    # Need a meaningful baseline to compare slopes
    baseline = max(abs(fh_slope), abs(sh_slope), 1.0)
    diff = sh_slope - fh_slope
    if abs(diff) < 0.25 * baseline:
        return None  # roughly linear
    return "accelerating" if diff > 0 else "decelerating"


def volatility_score(values: List[float]) -> Optional[float]:
    """Coefficient of variation (stddev / mean) for a numeric series.
    Returns None for series with mean ≤ 0 or fewer than 4 points.
    Useful to flag noisy / volatile time-series."""
    nums = [float(v) for v in values if v is not None and _isfinite(v)]
    if len(nums) < 4:
        return None
    mean = sum(nums) / len(nums)
    if mean <= 0:
        return None
    var = sum((v - mean) ** 2 for v in nums) / max(1, (len(nums) - 1))
    sd = math.sqrt(var)
    return sd / mean


# ─────────────────────────────────────────────────────────────────────
# Seasonality (monthly / quarterly time series)
# ─────────────────────────────────────────────────────────────────────

_QUARTER_END_MONTHS = {3, 6, 9, 12}


# ─────────────────────────────────────────────────────────────────────
# Internationalized month / weekday vocabularies
# ─────────────────────────────────────────────────────────────────────

_MONTHS_BY_LANG = {
    "es": ["enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "setiembre", "octubre",
            "noviembre", "diciembre"],
    "en": ["january", "february", "march", "april", "may", "june",
            "july", "august", "september", "october", "november", "december"],
    "pt": ["janeiro", "fevereiro", "março", "marco", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"],
}
_MONTHS_SHORT = {
    "es": ["ene", "feb", "mar", "abr", "may", "jun",
            "jul", "ago", "sep", "sept", "oct", "nov", "dic"],
    "en": ["jan", "feb", "mar", "apr", "may", "jun",
            "jul", "aug", "sep", "sept", "oct", "nov", "dec"],
    "pt": ["jan", "fev", "mar", "abr", "mai", "jun",
            "jul", "ago", "set", "out", "nov", "dez"],
}
_WEEKDAYS_BY_LANG = {
    "es": ["lunes", "martes", "miércoles", "miercoles", "jueves",
            "viernes", "sábado", "sabado", "domingo"],
    "en": ["monday", "tuesday", "wednesday", "thursday", "friday",
            "saturday", "sunday"],
    "pt": ["segunda", "terça", "terca", "quarta", "quinta", "sexta",
            "sábado", "sabado", "domingo"],
}


_MONTH_ALIASES = {
    # Map non-canonical short forms to their canonical 0-based month index
    "sept": 8, "setiembre": 8, "marco": 2, "março": 2,
}


def month_index(label: str) -> Optional[int]:
    """Return 0..11 if `label` is a month name (any supported language /
    short form), else None."""
    s = str(label).strip().lower()
    if not s:
        return None
    if s in _MONTH_ALIASES:
        return _MONTH_ALIASES[s]
    for lang, months in _MONTHS_BY_LANG.items():
        if s in months:
            idx = months.index(s)
            # ES list has a duplicate ("setiembre" right after "septiembre")
            return min(idx, 11)
    for lang, shorts in _MONTHS_SHORT.items():
        if s in shorts:
            return min(shorts.index(s), 11)
    return None


def weekday_index(label: str) -> Optional[int]:
    s = str(label).strip().lower()
    for lang, days in _WEEKDAYS_BY_LANG.items():
        if s in days:
            return days.index(s) % 7
    return None


def seasonality_hint(labels: List[str], values: List[float]) -> Optional[str]:
    """If labels look like Q1/Q2/Q3/Q4 codes (or YYYY-Qn), return a short hint
    describing peak/valley patterns. Returns None when no pattern is detectable."""
    if len(labels) < 4 or len(labels) != len(values):
        return None
    nums = [float(v) for v in values]
    quarters = [_extract_quarter(l) for l in labels]
    if any(q is None for q in quarters):
        return None
    # Average per quarter (1..4) across years
    by_q: dict = {1: [], 2: [], 3: [], 4: []}
    for q, v in zip(quarters, nums):
        by_q[q].append(v)
    avgs = {q: (sum(vs) / len(vs)) for q, vs in by_q.items() if vs}
    if len(avgs) < 4:
        return None
    peak_q = max(avgs, key=lambda k: avgs[k])
    valley_q = min(avgs, key=lambda k: avgs[k])
    if avgs[peak_q] <= avgs[valley_q] * 1.15:
        return None  # too flat
    quarter_names = {1: "Q1", 2: "Q2", 3: "Q3", 4: "Q4"}
    return (f"Patrón estacional: el {quarter_names[peak_q]} concentra los "
            f"valores más altos y el {quarter_names[valley_q]} los más bajos.")


def _extract_quarter(label: str) -> Optional[int]:
    s = str(label).upper()
    m = re.search(r"Q([1-4])", s)
    if m:
        return int(m.group(1))
    # YYYY-MM format
    m = re.match(r"\d{4}-(\d{2})", s)
    if m:
        month = int(m.group(1))
        if 1 <= month <= 12:
            return (month - 1) // 3 + 1
    return None


# ─────────────────────────────────────────────────────────────────────
# Semantic column detection
# ─────────────────────────────────────────────────────────────────────

_GEO_TOKENS = (
    # ES
    "ciudad", "departamento", "municipio", "región", "region", "país", "pais",
    "estado", "destino", "ubicación", "ubicacion", "sede", "zona",
    "provincia", "localidad", "barrio", "comuna",
    # EN
    "city", "country", "state", "region", "province", "location", "site",
    "area", "district", "address", "zip", "postcode",
    # PT
    "cidade", "país", "pais", "estado", "região", "regiao", "município",
    "municipio", "localização", "localizacao",
)
_PERSON_TOKENS = (
    # ES
    "nombre", "solicitante", "responsable", "cliente", "beneficiario",
    "empleado", "asesor", "vendedor", "contacto", "usuario", "titular",
    "gerente", "supervisor", "operario", "trabajador", "funcionario",
    # EN
    "name", "customer", "client", "user", "employee", "manager", "owner",
    "contact", "person", "buyer", "seller", "vendor", "rep", "agent",
    "requester", "approver",
    # PT
    "nome", "cliente", "funcionário", "funcionario", "vendedor", "gerente",
    "responsável", "responsavel", "usuário", "usuario",
)
_DATE_STRING_TOKENS = (
    # ES
    "fecha", "día", "dia", "mes", "año", "ano", "periodo", "período",
    "trimestre", "semestre", "fechado",
    # EN
    "date", "day", "month", "year", "period", "quarter", "week", "datetime",
    "timestamp",
    # PT
    "data", "dia", "mês", "mes", "ano", "trimestre", "período", "periodo",
)


def column_semantic_kind(name: str) -> Optional[str]:
    """Return a semantic label for a column based on its name. None if generic."""
    n = str(name).strip().lower()
    if any(t in n for t in _GEO_TOKENS):
        return "geo"
    if any(t in n for t in _PERSON_TOKENS):
        return "person"
    if any(t in n for t in _DATE_STRING_TOKENS):
        return "date_like"
    return None


# ─────────────────────────────────────────────────────────────────────
# Structural type detection (works regardless of column name / language)
# ─────────────────────────────────────────────────────────────────────

def is_year_column(samples: List, name: str = "",
                     threshold: float = 0.9) -> bool:
    """True when ≥`threshold` of non-null samples are integers in [1900, 2100].

    Two distinct years required by default — but when the column NAME signals
    'year' explicitly (año / year / fiscal year / etc.), even a constant
    column counts (e.g. a 'Año' = 2024 dimension column on a single-year
    Excel). This prevents the constant from being misclassified as numeric
    and turning into a 'Promedio de Año = 2024' KPI."""
    nums = []
    for s in samples:
        if s is None:
            continue
        try:
            f = float(s)
            if not _isfinite(f):
                continue
            nums.append(f)
        except (TypeError, ValueError):
            continue
    if len(nums) < 3:
        return False
    in_range = [n for n in nums if 1900 <= n <= 2100 and float(n).is_integer()]
    if len(in_range) / len(nums) < threshold:
        return False
    if len(set(in_range)) >= 2:
        return True
    # Constant-year case: only accept when the column name signals "year"
    name_lower = str(name).lower()
    return any(t in name_lower for t in ("año", "ano ", " ano",
                                            "year", "fiscal", "ejercicio"))


def is_percentage_column(samples: List, name: str = "") -> bool:
    """True when the values look like percentages.
    Allowed forms:
      - decimal in [-1, 1] with at least one non-trivial value (e.g. growth
        rates may be negative: -0.22 = -22%)
      - integer/float in [-100, 100] when the column name hints at %."""
    nums = []
    for s in samples:
        if s is None:
            continue
        try:
            f = float(s)
            if _isfinite(f):
                nums.append(f)
        except (TypeError, ValueError):
            continue
    if len(nums) < 3:
        return False
    name_l = str(name).lower()
    has_pct_hint = ("%" in name_l or "porcentaje" in name_l or "percent" in name_l
                     or "rate" in name_l or "tasa" in name_l
                     or "ratio" in name_l or "growth" in name_l
                     or "crecimiento" in name_l)
    # Decimal form (with negatives allowed for deltas): -1..1
    if all(-1 <= v <= 1 for v in nums) and any(v not in (0, 1, -1) for v in nums):
        return True
    # Integer/float -100..100 — only when name hints
    if has_pct_hint and all(-100 <= v <= 100 for v in nums):
        return True
    return False


def is_score_column(samples: List, name: str = "") -> bool:
    """True when values look like a small-range rating (e.g., 1..5, 1..10).
    Heuristic: all integer, distinct count ≤ 10, range ≤ 10."""
    nums = []
    for s in samples:
        if s is None:
            continue
        try:
            f = float(s)
            if _isfinite(f) and f.is_integer():
                nums.append(int(f))
        except (TypeError, ValueError):
            continue
    if len(nums) < 5:
        return False
    distinct = set(nums)
    if len(distinct) > 10:
        return False
    span = max(distinct) - min(distinct)
    if span > 10 or span < 2:
        return False
    return min(distinct) >= 0


def format_to_dtype(number_format: str) -> Optional[str]:
    """Map an Excel cell number_format string to one of our dtypes.

    Excel formats are like:
      "$#,##0.00"            → currency
      "#,##0.00 €"           → currency
      "0.00%"                → percent
      "yyyy-mm-dd"           → date
      "h:mm:ss"              → date (datetime)
      "0"                    → numeric (no formatting hints)
      "General"              → None (no info)

    Returns one of {"currency","percent","date","numeric"} or None
    when the format is unknown/General. None means: defer to our structural
    detectors (the format gives no useful signal)."""
    if not number_format:
        return None
    fmt = str(number_format).strip()
    if not fmt or fmt.lower() == "general":
        return None
    # Currency markers — language/locale agnostic
    if any(sym in fmt for sym in ("$", "€", "£", "¥", "₩", "₹", "₽")):
        return "currency"
    if "[$" in fmt:  # Excel locale-aware currency: [$USD-409]#,##0.00
        return "currency"
    if "%" in fmt:
        return "percent"
    # Date / datetime — Excel uses y/m/d/h tokens
    fmt_low = fmt.lower()
    has_date_token = any(t in fmt_low for t in ("yyyy", "yy", "mmm", "mmmm",
                                                  "dd", "ddd", "dddd"))
    has_time_token = any(t in fmt_low for t in ("h:", "hh", "ss", "am/pm",
                                                  "a/p"))
    # `mm` and `m` are ambiguous (month vs minute) — only count if combined
    if has_date_token or (has_time_token and ("m" in fmt_low or "y" in fmt_low)):
        return "date"
    # A plain numeric format with thousands separator and no special markers
    if any(c in fmt for c in ("0", "#")):
        return "numeric"
    return None


def looks_money_by_name(name: str) -> bool:
    """True when the column name is in the i18n money-hint vocabulary."""
    n = str(name).lower()
    return any(t in n for t in _MONEY_HINTS)


def is_currency_by_magnitude(samples: List) -> bool:
    """Currency by structure (no name): all values >= 0, magnitude is large
    (≥1000 typical), at least 70% are 'round-ish' (no >2 decimal places),
    and the values are NOT a sequential ID-like progression."""
    nums = []
    for s in samples:
        if s is None:
            continue
        try:
            f = float(s)
            if _isfinite(f):
                nums.append(f)
        except (TypeError, ValueError):
            continue
    if len(nums) < 5:
        return False
    if any(n < 0 for n in nums):
        return False
    if max(nums) == min(nums):
        return False
    big = [n for n in nums if n >= 1000]
    if len(big) / len(nums) < 0.5:
        return False
    rounded = [round(n, 2) == round(n, 6) for n in nums]
    if sum(rounded) / len(nums) < 0.7:
        return False
    # Reject ID-like sequences: all distinct integers, sample is approximately
    # a contiguous range. Currency rarely satisfies this.
    if _looks_like_sequential_ids(nums):
        return False
    return True


def is_total_row(row_values: List, prev_rows_numeric: List[List[float]]) -> bool:
    """Heuristic: a row is a 'total/subtotal' if it has at least one text cell
    containing 'total/subtotal/suma/sum/grand total' AND its numeric cells are
    suspiciously large compared to the median of preceding rows.

    `row_values` is the row's cell values (in column order).
    `prev_rows_numeric` is a list of prior rows' numeric values (parallel
    columns) — used to assess whether the row's numbers look like aggregates.
    """
    if not row_values:
        return False
    has_total_word = False
    for v in row_values:
        if v is None:
            continue
        s = str(v).strip().lower()
        if s in ("total", "subtotal", "suma", "sum", "grand total",
                  "totals", "totales", "total general"):
            has_total_word = True
            break
        if any(s.startswith(t) for t in ("total ", "subtotal ", "suma ",
                                           "sum ", "totales ")):
            has_total_word = True
            break
    return has_total_word


def _looks_like_sequential_ids(nums: List[float]) -> bool:
    """Heuristic: all integers, all distinct, and samples span a contiguous
    range (max - min ≈ count - 1). True for IDs 1..N or any near-contiguous
    sequence; almost never true for monetary amounts."""
    if not all(float(n).is_integer() for n in nums):
        return False
    ints = [int(n) for n in nums]
    if len(set(ints)) != len(ints):
        return False  # has duplicates → not IDs
    span = max(ints) - min(ints)
    # Allow up to 2x slack: a sampled monotone integer column can have small
    # gaps but ID columns will be very close to span = N - 1.
    return span <= max(20, len(ints) * 2)


def is_constant_column(samples: List) -> bool:
    """All non-null values are the same — the column carries no information
    and should be skipped from KPIs/charts."""
    distinct = set()
    seen = 0
    for s in samples:
        if s is None or (isinstance(s, str) and not s.strip()):
            continue
        seen += 1
        distinct.add(s)
        if len(distinct) > 1:
            return False
    return seen >= 3 and len(distinct) == 1


def looks_boolean_disguised(samples: List) -> bool:
    """True when distinct non-null values ⊂ {sí/no, yes/no, true/false,
    activo/inactivo, 0/1, etc.} and there are exactly two distinct values."""
    seen = set()
    for s in samples:
        if s is None:
            continue
        v = str(s).strip().lower()
        if not v:
            continue
        seen.add(v)
        if len(seen) > 2:
            return False
    if len(seen) != 2:
        return False
    bool_pairs = (
        {"sí", "no"}, {"si", "no"}, {"yes", "no"},
        {"true", "false"}, {"verdadero", "falso"}, {"verdadeiro", "falso"},
        {"activo", "inactivo"}, {"active", "inactive"}, {"ativo", "inativo"},
        {"1", "0"}, {"1.0", "0.0"},
        {"aprobado", "rechazado"}, {"approved", "rejected"},
        {"on", "off"}, {"sí", "no "}, {"open", "closed"},
    )
    return any(seen == p for p in bool_pairs)


def detect_computed_column(target_values: List[float],
                            other_columns: List[Tuple[str, List[float]]],
                            tolerance: float = 0.01) -> Optional[str]:
    """Detect whether `target_values` ≈ sum of two other numeric columns.
    Returns 'A + B' (label form) when found. Conservative: requires ≥10 rows
    and ≥80% of rows to match within `tolerance` relative error."""
    if len(target_values) < 10 or len(other_columns) < 2:
        return None
    t = [float(v) if v is not None else None for v in target_values]
    n = len(t)
    for i in range(len(other_columns)):
        a_name, a_vals = other_columns[i]
        a = list(a_vals)
        if len(a) != n:
            continue
        for j in range(i + 1, len(other_columns)):
            b_name, b_vals = other_columns[j]
            b = list(b_vals)
            if len(b) != n:
                continue
            matches = 0
            valid = 0
            for ti, ai, bi in zip(t, a, b):
                if ti is None or ai is None or bi is None:
                    continue
                expected = ai + bi
                if abs(expected) < 1e-9:
                    if abs(ti) < 1e-9:
                        matches += 1
                    valid += 1
                    continue
                if abs(ti - expected) / max(abs(expected), 1e-9) <= tolerance:
                    matches += 1
                valid += 1
            if valid >= 10 and matches / valid >= 0.8:
                return f"{a_name} + {b_name}"
    return None


# ─────────────────────────────────────────────────────────────────────
# PII detection + masking
# ─────────────────────────────────────────────────────────────────────

_RE_EMAIL = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
# Phone: REQUIERE al menos un indicador real de teléfono (+, paréntesis,
# guión, o espacio interno). Pure-digit strings y floats con '.' NO son
# phones — eran la fuente principal de falsos positivos en columnas de
# dinero (ej. "1500000.50" matcheaba antes con `[\d\s().-]`).
_RE_PHONE = re.compile(
    r"^"
    r"(?:\+\d[\d\s().-]{7,}\d"          # comienza con + (internacional), o
    r"|\(\d{2,4}\)[\s\d.-]{6,}\d"      # tiene paréntesis explícitos, o
    r"|\d[\d ]{2,}-[\d -]{4,}\d"       # tiene guiones internos
    r"|\d{2,4}\s\d{3,4}\s\d{4,})"      # tiene espacios separando bloques
    r"$"
)
_RE_DOC_ID = re.compile(r"^\d{7,11}$")


def detect_pii_kind(value) -> Optional[str]:
    """Identify the PII kind of a single cell value: 'email' | 'phone' | 'doc_id'."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    if _RE_EMAIL.match(s):
        return "email"
    if _RE_DOC_ID.match(s):
        return "doc_id"
    if _RE_PHONE.match(s):
        return "phone"
    return None


def mask_pii(value, kind: str) -> str:
    """Mask `value` based on its PII kind."""
    s = str(value).strip()
    if kind == "email":
        local, _, domain = s.partition("@")
        return (f"{local[:1]}***@***" if local
                else "***@***")
    if kind == "doc_id":
        return ("***" + s[-4:]) if len(s) >= 4 else "***"
    if kind == "phone":
        digits = re.sub(r"\D", "", s)
        return ("***" + digits[-4:]) if len(digits) >= 4 else "***"
    return s


_PII_COL_HINTS = {
    "doc_id": (
        # ES / LatAm
        "cédula", "cedula", "documento", "doc", "dni", "rfc", "rut", "nit",
        "curp", "ssn", "nif", "passport", "pasaporte", "identificación",
        "identificacion", "ced ",
        # EN
        "id number", "ssn", "tax id", "national id", "license", "license_no",
        # PT
        "cpf", "cnpj", "rg",
    ),
    "email": ("email", "correo", "e-mail", "mail", "e_mail", "correio"),
    "phone": (
        "teléfono", "telefono", "celular", "móvil", "movil", "phone",
        "fono", "contacto telefónico", "mobile", "tel", "telephone",
        "telefone", "celular",
    ),
}
# Column names that should NEVER be enmascared as doc_id (numeric magnitudes)
_MONEY_HINTS = (
    # ES
    "valor", "monto", "precio", "costo", "coste", "total", "ingreso",
    "salario", "importe", "venta", "comprado", "pago", "subtotal", "deuda",
    "saldo", "balance", "facturado",
    # EN
    "amount", "price", "cost", "revenue", "sales", "income", "salary",
    "payment", "balance", "billed", "invoice", "value",
    # PT
    "valor", "preço", "preco", "custo", "receita", "venda", "pagamento",
    "salário", "salario",
)


def is_pii_column(samples: List, col_name: str = "",
                    threshold: float = 0.7) -> Optional[str]:
    """Given a list of sample cells AND the column name, return the dominant
    PII kind if at least `threshold` of values match AND (for doc_id) the
    column name does NOT look like a money/quantity column.

    `col_name` is highly recommended — without it, we cannot tell '1234567'
    (a 7-digit ID) from '1234567' (an amount in COP). When in doubt, return
    None (safer to leave un-masked than to mask a financial column)."""
    valid = [s for s in samples if s is not None and str(s).strip()]
    if not valid:
        return None
    kinds: dict = {}
    for v in valid:
        k = detect_pii_kind(v)
        if k:
            kinds[k] = kinds.get(k, 0) + 1
    if not kinds:
        return None
    best_kind, best_count = max(kinds.items(), key=lambda kv: kv[1])
    if best_count / len(valid) < threshold:
        return None

    name_lower = str(col_name or "").lower()
    is_money_col = any(h in name_lower for h in _MONEY_HINTS)

    # Money columns are never doc_ids ni phones, no importa cómo se vean
    # los dígitos. Antes sólo cubríamos doc_id; ahora también phone porque
    # floats como '1500000.50' pueden parecer phones tras endurecer la regex,
    # pero la columna tipo 'Subtotal'/'Monto' nos dice que NO lo son.
    if best_kind in ("doc_id", "phone") and is_money_col:
        return None

    if best_kind == "doc_id":
        # Require explicit hint in the column name to enmascarar — silent
        # masking of opaque numeric columns is risky
        if not any(h in name_lower for h in _PII_COL_HINTS["doc_id"]):
            return None

    # Email / phone: if the regex matched the cell content itself, that's
    # already very specific (foo@bar.com). Allow even without column hint.
    return best_kind


# ─────────────────────────────────────────────────────────────────────
# Smart KPI subsumption (Total General ≈ Σ sub-totales)
# ─────────────────────────────────────────────────────────────────────

def subsumed_kpi_ids(kpi_blocks: List, tolerance: float = 0.05
                      ) -> List[str]:
    """Return KPI block IDs that are likely 'subsumed' by another KPI whose
    value approximately equals the sum of the others.

    Heuristic: for each currency-sum KPI A on sheet S, look at OTHER currency-
    sum KPIs on the SAME sheet whose values sum to within `tolerance` of A's
    value. If found, the others are 'subsumed' by A — A is the meaningful
    headline KPI; the rest are breakdown lines.
    """
    if not kpi_blocks:
        return []
    # Group by (sheet, state_bucket) so an "effective total" headline can
    # subsume its sub-components without being incorrectly compared against
    # the rejected-bucket KPIs (which would never sum to it).
    by_group: dict = {}
    for b in kpi_blocks:
        if b.kind != "kpi_candidate":
            continue
        if b.extra.get("agg") not in ("sum", "sum_filtered"):
            continue
        v = b.extra.get("value")
        if v is None:
            continue
        sheet = b.provenance.sheet
        bucket = b.extra.get("state_bucket")  # None for raw sums
        by_group.setdefault((sheet, bucket), []).append(b)

    subsumed: set = set()
    for _key, blocks in by_group.items():
        for big in blocks:
            big_v = float(big.extra.get("value") or 0)
            if big_v <= 0:
                continue
            others = [b for b in blocks if b.id != big.id and b.id not in subsumed]
            if len(others) < 2:
                continue
            other_sum = sum(float(b.extra.get("value") or 0) for b in others)
            if other_sum <= 0:
                continue
            # If others sum to within tolerance of big, big is the General Total
            if abs(other_sum - big_v) / big_v <= tolerance:
                for b in others:
                    subsumed.add(b.id)
    return list(subsumed)


# ─────────────────────────────────────────────────────────────────────
# KPI ratio pairing (Solicitado / Aprobado, etc.)
# ─────────────────────────────────────────────────────────────────────

# (numerator_token, denominator_token, label) — pct = numerator / denominator * 100
# So for "% Aprobación" the numerator must be 'aprobado' (qty approved) and the
# denominator 'solicitado' (qty requested). Order matters!
_PAIR_TOKENS = [
    # ES
    ("aprobado", "solicitado", "% Aprobación"),
    ("aprobada", "solicitada", "% Aprobación"),
    ("ejecutado", "aprobado", "% Ejecución"),
    ("ejecutado", "presupuestado", "% Ejecución presupuestal"),
    ("ejecutado", "presupuesto", "% Ejecución presupuestal"),
    ("final", "inicial", "Variación final/inicial"),
    ("real", "planeado", "% Cumplimiento"),
    ("real", "estimado", "% Precisión estimación"),
    ("facturado", "vendido", "% Facturación"),
    ("ejecutado", "solicitado", "% Ejecución"),
    ("real", "meta", "% Cumplimiento meta"),
    ("real", "objetivo", "% Cumplimiento objetivo"),
    # EN
    ("approved", "requested", "% Approval"),
    ("actual", "budgeted", "% Budget execution"),
    ("actual", "budget", "% Budget execution"),
    ("actual", "planned", "% Plan compliance"),
    ("actual", "forecast", "% Forecast accuracy"),
    ("actual", "target", "% Target attainment"),
    ("actual", "expected", "% Expected accuracy"),
    ("closed", "opened", "% Closure rate"),
    # PT
    ("aprovado", "solicitado", "% Aprovação"),
    ("realizado", "planejado", "% Cumprimento"),
    ("realizado", "orçado", "% Execução orçamentária"),
    ("realizado", "orcado", "% Execução orçamentária"),
    ("realizado", "meta", "% Cumprimento meta"),
]


def find_ratio_pairs(kpi_blocks: List) -> List[dict]:
    """Find pairs of KPIs whose names match a known (numerator, denominator)
    pattern. Returns [{label, num_block, den_block, percent_value}]."""
    by_norm: dict = {}
    for b in kpi_blocks:
        if b.kind != "kpi_candidate":
            continue
        v = b.extra.get("value")
        if v is None:
            continue
        norm = _norm_name(b.label)
        by_norm[norm] = b

    pairs = []
    seen_keys: set = set()
    for num_tok, den_tok, label in _PAIR_TOKENS:
        for norm_a, block_a in by_norm.items():
            if num_tok not in norm_a:
                continue
            base_a = norm_a.replace(num_tok, "").strip()
            for norm_b, block_b in by_norm.items():
                if block_a.id == block_b.id:
                    continue
                if den_tok not in norm_b:
                    continue
                base_b = norm_b.replace(den_tok, "").strip()
                # Bases must be similar (same underlying concept)
                if not _names_close(base_a, base_b):
                    continue
                key = tuple(sorted([block_a.id, block_b.id]))
                if key in seen_keys:
                    continue
                seen_keys.add(key)
                num_v = float(block_a.extra.get("value") or 0)
                den_v = float(block_b.extra.get("value") or 0)
                if den_v <= 0:
                    continue
                pct = num_v / den_v * 100
                pairs.append({
                    "label": label,
                    "num_block": block_a,
                    "den_block": block_b,
                    "percent_value": pct,
                })
    return pairs


def _norm_name(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())


def _names_close(a: str, b: str) -> bool:
    """Two base names are 'close' if their word sets overlap by ≥50%."""
    aw = set(re.findall(r"\w+", a))
    bw = set(re.findall(r"\w+", b))
    if not aw or not bw:
        return True  # both empty bases — consider match
    common = aw & bw
    return len(common) / max(len(aw), len(bw)) >= 0.5


# ─────────────────────────────────────────────────────────────────────
# Multi-sheet awareness
# ─────────────────────────────────────────────────────────────────────

_MERGED_HEADER_PREFIX_RE = re.compile(r"^[^>]+\s*>\s*")


def _normalize_col_for_overlap(name) -> str:
    """Strip merged-header prefix ("COMISIONES > Ciudad Destino" → "ciudad
    destino") + 'Unnamed: N' artifacts + whitespace + case. Without this,
    a sheet whose first physical row was a merged title gets prefixed
    column names by the parser and never matches its smaller-sample twin,
    so the smaller sample wins chart slots even though it's a subset of
    the authoritative sheet."""
    s = str(name or "").strip()
    # Iteratively strip merged-header prefixes (handles nested "A > B > col")
    while True:
        new = _MERGED_HEADER_PREFIX_RE.sub("", s)
        if new == s:
            break
        s = new
    s = s.strip().lower()
    # 'Unnamed: 0'-style headers carry no signal — drop them to a sentinel
    # that won't match across sheets.
    if s.startswith("unnamed:") or s in ("nan", "none", ""):
        return ""
    return s


def redundant_sheet_ids(sheets) -> List[str]:
    """Return sheet names that are likely 'redundant' duplicates of another
    sheet in the same workbook.

    A sheet is redundant when it shares ≥70% of (normalized) column names
    with another sheet AND it has fewer rows than that other sheet (it's
    a sample / subset), OR identical row count but appears later in the
    workbook (templates copy-pasted as 'Sheet1', 'Sheet2' → keep first).

    Column-name normalization strips merged-header prefixes so that
    "COMISIONES > Ciudad Destino" matches "Ciudad Destino" across the
    sample/master split that's typical in audit workbooks. Without that,
    the parser's prefix handling silently breaks the comparison and
    smaller-sample sheets (Hoja1, 31 rows) stop being recognized as
    subsets of authoritative ones (Comisiones- Base, 1852 rows).

    The earliest / largest sheet is kept; later/smaller copies are flagged.
    """
    if not sheets or len(sheets) < 2:
        return []
    redundant: set = set()
    for i, a in enumerate(sheets):
        if a.name in redundant:
            continue
        cols_a = {_normalize_col_for_overlap(c.name) for c in a.columns}
        cols_a.discard("")
        if not cols_a:
            continue
        for j, b in enumerate(sheets):
            if i == j or b.name in redundant:
                continue
            cols_b = {_normalize_col_for_overlap(c.name) for c in b.columns}
            cols_b.discard("")
            if not cols_b:
                continue
            overlap = len(cols_a & cols_b) / max(len(cols_a), len(cols_b))
            if overlap < 0.7:
                continue
            rows_a = a.shape[0]
            rows_b = b.shape[0]
            # Smaller sheet is the redundant one. We deliberately do NOT
            # require a.fill_ratio <= b.fill_ratio: a 31-row scratchpad
            # almost always has denser fill than a 1852-row production
            # table (because optional columns are filled in the small
            # curated sample but not in the big real one), and that
            # historically blocked the redundancy flag.
            if rows_a < rows_b:
                redundant.add(a.name)
                break
            # Identical row count → keep the FIRST one in the workbook
            if rows_a == rows_b and j < i:
                redundant.add(a.name)
                break
    return list(redundant)


# ─────────────────────────────────────────────────────────────────────
# Auto-prompt inference (when the user provides no prompt)
# ─────────────────────────────────────────────────────────────────────

def infer_prompt(wb_filename: str, sheet_names: List[str],
                  column_names: List[str]) -> str:
    """Infer a short prompt describing what's likely in this workbook.
    Used only when the user did not supply one. Always returns a string."""
    base = re.sub(r"[._-]+", " ", str(wb_filename or "")).strip()
    base = re.sub(r"\.xlsx?$", "", base, flags=re.IGNORECASE).strip()
    if base:
        topic = base
    elif sheet_names:
        topic = sheet_names[0]
    else:
        topic = "los datos"

    hints = []
    norm_cols = " ".join(c.lower() for c in column_names)
    if any(t in norm_cols for t in ("ciudad", "destino", "ubicación")):
        hints.append("distribución geográfica")
    if any(t in norm_cols for t in ("fecha", "mes", "año", "periodo")):
        hints.append("evolución temporal")
    if any(t in norm_cols for t in ("valor", "monto", "total", "precio",
                                     "costo", "ingreso", "salario")):
        hints.append("indicadores financieros")
    if any(t in norm_cols for t in ("estado", "status", "categoría",
                                     "categoria", "tipo")):
        hints.append("composición por categorías")
    if hints:
        return f"Análisis ejecutivo de {topic}: {', '.join(hints)}."
    return f"Análisis ejecutivo de {topic}."


# ─────────────────────────────────────────────────────────────────────
# Status-column awareness (filter detection)
# ─────────────────────────────────────────────────────────────────────

def classify_status_value(value) -> str:
    """Classify a single status cell as 'positive', 'negative', 'pending', or
    'unknown'. Case- and accent-insensitive substring match against curated
    token lists. 'unknown' is returned when no token matches — we never guess.
    """
    if value is None:
        return "unknown"
    s = str(value).strip().upper()
    if not s:
        return "unknown"
    for tok in NEGATIVE_STATUS_TOKENS:
        if tok in s:
            return "negative"
    for tok in POSITIVE_STATUS_TOKENS:
        if tok in s:
            return "positive"
    for tok in PENDING_STATUS_TOKENS:
        if tok in s:
            return "pending"
    return "unknown"


def is_status_filter_column(col_name: str, top_values) -> bool:
    """True iff a column looks like a workflow status filter (its name hints at
    'Estado'/'Status' AND at least one of its frequent values is a negative
    token like RECHAZADO/ANULADO). Conservative: requires BOTH gates to fire
    so a column named 'Tipo' or a status column without any rejected rows
    does not trigger the disclaimer pipeline.
    """
    name_l = (col_name or "").strip().lower()
    if not name_l:
        return False
    if not any(tok in name_l for tok in STATUS_COLUMN_NAME_TOKENS):
        return False
    if not top_values:
        return False
    for entry in top_values:
        # top_values is List[[value, count]]
        try:
            v = entry[0]
        except (TypeError, IndexError):
            continue
        if classify_status_value(v) == "negative":
            return True
    return False


def compute_status_breakdown(df, currency_col: str, status_col: str) -> dict:
    """Group `currency_col` by `status_col` and bucket the per-state sums into
    positive / negative / pending / unknown. Returns a dict ready to attach
    to a Block's `extra` field:

        {
          "status_column": "Estado",
          "per_state": {"CONTABILIZADO": 1186_741_209.0, ...},
          "positive_sum": float, "positive_count": int, "positive_states": [...],
          "negative_sum": float, "negative_count": int, "negative_states": [...],
          "pending_sum":  float, "pending_count":  int, "pending_states":  [...],
          "unknown_sum":  float, "unknown_count":  int, "unknown_states":  [...],
          "gross_sum":    float, "gross_count":    int,
        }

    Returns {} on any failure — caller treats absence as "no breakdown".
    """
    try:
        import pandas as pd  # local import keeps the insights module dep-light
    except Exception:
        return {}
    if df is None or currency_col not in df.columns or status_col not in df.columns:
        return {}
    try:
        nums = pd.to_numeric(df[currency_col], errors="coerce")
        nums = nums.replace([float("inf"), float("-inf")], pd.NA)
        sub = df[[status_col]].assign(__val=nums).dropna(subset=["__val"])
        if sub.empty:
            return {}
        grouped = sub.groupby(status_col, dropna=False)["__val"]
        per_state_sum = grouped.sum().to_dict()
        per_state_cnt = grouped.count().to_dict()
    except Exception:
        return {}

    buckets = {
        "positive": {"sum": 0.0, "count": 0, "states": []},
        "negative": {"sum": 0.0, "count": 0, "states": []},
        "pending":  {"sum": 0.0, "count": 0, "states": []},
        "unknown":  {"sum": 0.0, "count": 0, "states": []},
    }
    per_state_clean: dict = {}
    for state, total in per_state_sum.items():
        try:
            t = float(total)
        except (TypeError, ValueError):
            continue
        if not math.isfinite(t):
            continue
        per_state_clean[str(state) if state is not None else ""] = t
        bucket = classify_status_value(state)
        buckets[bucket]["sum"] += t
        buckets[bucket]["count"] += int(per_state_cnt.get(state, 0))
        buckets[bucket]["states"].append(str(state) if state is not None else "")

    gross_sum = sum(b["sum"] for b in buckets.values())
    gross_count = sum(b["count"] for b in buckets.values())
    out = {
        "status_column": status_col,
        "per_state": per_state_clean,
        "gross_sum": gross_sum,
        "gross_count": gross_count,
    }
    for k, v in buckets.items():
        out[f"{k}_sum"] = v["sum"]
        out[f"{k}_count"] = v["count"]
        out[f"{k}_states"] = v["states"]
    return out


# ─────────────────────────────────────────────────────────────────────
# Internal utils
# ─────────────────────────────────────────────────────────────────────

def _isfinite(v) -> bool:
    try:
        f = float(v)
        return not (math.isnan(f) or math.isinf(f))
    except (TypeError, ValueError):
        return False
