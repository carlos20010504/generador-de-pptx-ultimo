"""Pure-data Excel parser. No semantic inference."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Any
import pandas as pd
from socya_pipeline.errors import PipelineError, ErrorCode

@dataclass
class ColumnData:
    name: str
    dtype: str            # "numeric" | "currency" | "date" | "categorical" | "text" | "bool"
    n_unique: int
    fill_ratio: float
    samples: List[Any] = field(default_factory=list)
    min: Any = None
    max: Any = None
    mean: Any = None
    sum: Any = None
    top_values: List[List[Any]] = field(default_factory=list)  # [[value, count], ...]

@dataclass
class SheetData:
    name: str
    shape: tuple              # (rows, cols)
    fill_ratio: float
    columns: List[ColumnData]
    first_rows: List[List[Any]]   # up to 8 rows in compact form

@dataclass
class WorkbookData:
    filename: str
    sheets: List[SheetData]

def parse_workbook(path) -> WorkbookData:
    p = Path(path)
    try:
        xls = pd.ExcelFile(p, engine=None)  # auto-detect openpyxl/xlrd
    except Exception as e:
        raise PipelineError(
            ErrorCode.EXCEL_INVALID,
            "El archivo no se pudo leer como Excel.",
            details=str(e)[:200],
            user_action="upload_again",
        )

    sheets: List[SheetData] = []
    for sheet_name in xls.sheet_names:
        try:
            df = xls.parse(sheet_name)
        except Exception:
            continue
        df = _promote_real_headers(xls, sheet_name, df)
        sheets.append(_summarize_sheet(sheet_name, df))

    if not sheets:
        raise PipelineError(
            ErrorCode.EXCEL_EMPTY,
            "El Excel se abrió pero no tiene hojas legibles.",
            user_action="upload_again",
        )

    return WorkbookData(filename=p.name, sheets=sheets)


def promote_real_headers(xls, sheet_name: str, df: pd.DataFrame) -> pd.DataFrame:
    """Public wrapper — see _promote_real_headers."""
    return _promote_real_headers(xls, sheet_name, df)


def _promote_real_headers(xls, sheet_name: str, df: pd.DataFrame) -> pd.DataFrame:
    """If pandas inferred 'Unnamed: N' headers (common when row 0 is a title or the
    sheet has merged-cell title rows), scan the first 10 rows for the first row
    where >=50% cells are non-numeric strings — that is the real header row.
    Re-read with header=that row.
    """
    cols = [str(c) for c in df.columns]
    unnamed_ratio = sum(1 for c in cols if c.startswith("Unnamed:")) / max(1, len(cols))
    if unnamed_ratio < 0.3:
        return df  # headers look fine

    # Scan first 10 rows for a likely header row
    scan_rows = min(10, len(df))
    for i in range(scan_rows):
        row = df.iloc[i]
        non_null = [v for v in row if pd.notna(v)]
        if len(non_null) < max(2, len(cols) * 0.5):
            continue
        # A header row has mostly strings, not numbers
        string_cells = sum(1 for v in non_null
                           if isinstance(v, str) and not _looks_numeric(v))
        if string_cells / len(non_null) >= 0.7:
            try:
                # df.iloc[i] is row (i+1) in the source file because pandas
                # consumed row 0 as the default header.
                return xls.parse(sheet_name, header=i + 1)
            except Exception:
                return df
    return df


def _looks_numeric(s: str) -> bool:
    s = s.strip().replace(",", ".").replace("$", "").replace("%", "")
    try:
        float(s)
        return True
    except (ValueError, AttributeError):
        return False


def _summarize_sheet(name: str, df: pd.DataFrame) -> SheetData:
    rows, cols = df.shape
    if rows == 0 or cols == 0:
        return SheetData(name=name, shape=(rows, cols), fill_ratio=0.0,
                         columns=[], first_rows=[])
    fill_ratio = float(df.notna().sum().sum()) / max(1, rows * cols)
    columns = [_summarize_column(c, df[c]) for c in df.columns]
    first_rows = df.head(8).where(df.head(8).notna(), None).values.tolist()
    return SheetData(name=name, shape=(rows, cols), fill_ratio=fill_ratio,
                     columns=columns, first_rows=first_rows)


def _summarize_column(name: Any, series: pd.Series) -> ColumnData:
    name_str = str(name)
    fill = float(series.notna().sum()) / max(1, len(series))
    n_unique = int(series.nunique(dropna=True))
    dtype = _infer_dtype(name_str, series)
    samples = series.dropna().head(8).tolist()

    col = ColumnData(name=name_str, dtype=dtype, n_unique=n_unique,
                     fill_ratio=fill, samples=_jsonify(samples))

    if dtype in ("numeric", "currency"):
        nums = pd.to_numeric(series, errors="coerce").dropna()
        if not nums.empty:
            col.min = float(nums.min())
            col.max = float(nums.max())
            col.mean = float(nums.mean())
            col.sum = float(nums.sum())
    elif dtype in ("categorical", "text"):
        vc = series.dropna().astype(str).value_counts().head(8)
        col.top_values = [[v, int(c)] for v, c in vc.items()]

    return col


def _infer_dtype(name: str, series: pd.Series) -> str:
    if pd.api.types.is_bool_dtype(series):
        return "bool"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "date"
    if pd.api.types.is_numeric_dtype(series):
        lower = name.lower()
        if any(t in lower for t in ("total", "valor", "precio", "costo",
                                      "salario", "monto", "ingreso")):
            return "currency"
        return "numeric"
    n_unique = series.nunique(dropna=True)
    if n_unique > 0 and n_unique <= max(20, len(series) * 0.05):
        return "categorical"
    return "text"


def _jsonify(values):
    out = []
    for v in values:
        if pd.isna(v):
            out.append(None)
        elif hasattr(v, "isoformat"):
            out.append(v.isoformat())
        elif isinstance(v, (int, float, str, bool)):
            out.append(v)
        else:
            out.append(str(v))
    return out
