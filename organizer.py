import pandas as pd
import numpy as np
import sys
import json
import os
import re
import unicodedata

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

# ═══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ═══════════════════════════════════════════════════════════════════════════════
MAX_SLIDES = 25
ROWS_PER_TABLE_SLIDE = 12
MIN_FILL_RATIO = 0.25  # Al menos 25% de celdas con datos para ser válida
MAX_KPIS = 6
MAX_CHART_CATEGORIES = 6
MAX_PIE_CATEGORIES = 4
MAX_BAR_CATEGORIES = 5
MAX_AUTO_CHARTS = 3
MAX_TABLE_COLS = 7
MAX_CONCLUSIONES = 10
MAX_INSIGHTS_AVANZADOS = 8

PLACEHOLDER_VALS = {'???', '—', 'n/a', 'na', 'nan', 'none', '', '0', '-',
                    'null', 'sin datos', 'sin información', 'sin dato',
                    'no aplica', 'no disponible', 'nd', 's/d'}

SHEET_FAMILY_LABELS = {
    'auditoria': 'auditoria',
    'checklist': 'checklist',
    'coso': 'coso',
    'hallazgos': 'hallazgos',
    'oportunidades': 'oportunidades',
    'matriz_riesgos': 'matriz_riesgos',
    'evidencias': 'evidencias',
    'arqueo': 'arqueo',
    'procedimiento': 'procedimiento',
    'cuestionario': 'cuestionario',
    'distribucion': 'distribucion',
    'general': 'general',
}

PRIMARY_SHEET_FAMILY_SCORES = {
    'auditoria': 520,
    'checklist': 260,
    'matriz_riesgos': 180,
    'hallazgos': 120,
    'oportunidades': 80,
    'cuestionario': 40,
    'coso': 20,
    'distribucion': -80,
    'procedimiento': -140,
    'evidencias': -220,
    'arqueo': -260,
    'general': 0,
}

# ═══════════════════════════════════════════════════════════════════════════════
# UTILIDADES DE LIMPIEZA
# ═══════════════════════════════════════════════════════════════════════════════

def normalizar_columnas_unicas(columns):
    usadas = {}
    resultado = []
    for idx, col in enumerate(columns):
        nombre = str(col).strip() if pd.notna(col) and str(col).strip() and not str(col).startswith('Unnamed') else f"Col_{idx}"
        if nombre in usadas:
            usadas[nombre] += 1
            nombre = f"{nombre}_{usadas[nombre]}"
        else:
            usadas[nombre] = 0
        resultado.append(nombre)
    return resultado


def extract_real_sheet(df):
    """Encuentra la fila de encabezados reales dentro de un DataFrame."""
    try:
        raw_data = [df.columns.tolist()] + df.values.tolist()
        best_row_idx = 0
        max_valid = 0
        for i, row in enumerate(raw_data[:20]):
            valid_cols = [str(x) for x in row if pd.notna(x) and str(x).strip() 
                         and not str(x).startswith('Unnamed') 
                         and 'TITLE:' not in str(x) 
                         and 'TYPE:' not in str(x) 
                         and 'SUBTITLE:' not in str(x)]
            if len(valid_cols) > max_valid:
                max_valid = len(valid_cols)
                best_row_idx = i
                
        if max_valid >= 2:
            new_header = raw_data[best_row_idx]
            new_data = raw_data[best_row_idx+1:]
            df_new = pd.DataFrame(new_data, columns=new_header)
            df_new.columns = normalizar_columnas_unicas(df_new.columns)
            df_new = df_new.dropna(axis=1, how='all').dropna(axis=0, how='all')
            return df_new
    except: pass
    return df


def limpiar_df(df):
    df = df.copy()
    df.columns = normalizar_columnas_unicas(df.columns)
    for idx, col in enumerate(df.columns):
        serie = df.iloc[:, idx]
        col_str = str(col).lower()
        if 'fecha' in col_str or 'date' in col_str or 'mes' in col_str:
            try:
                if serie.dtype == 'float64' or serie.dtype == 'int64':
                    df.iloc[:, idx] = pd.to_datetime(serie, unit='D', origin='1899-12-30').dt.strftime('%d/%m/%Y')
            except:
                pass
        if serie.dtype == object or serie.dtype == 'string':
            df.iloc[:, idx] = serie.fillna('—')
        else:
            df.iloc[:, idx] = serie.fillna(0)
    return df


def remover_filas_basura(df):
    df = df.copy()
    df.columns = normalizar_columnas_unicas(df.columns)
    palabras_basura = ['diapositiva', 'power point', 'powerpoint', 
                       'agrega', 'información', 'columna', 'imagen adjunta', 'placeholder']
    mask = pd.Series([True] * len(df), index=df.index)
    for idx, col in enumerate(df.columns):
        serie = df.iloc[:, idx]
        if serie.dtype == object:
            for palabra in palabras_basura:
                mask &= ~serie.astype(str).str.lower().str.contains(palabra, na=False)
    df = df[mask]
    c_importantes = ['Id Comisión', 'Solicitante', 'Valor Total Solicitado']
    cols_check = [c for c in c_importantes if c in df.columns]
    if cols_check:
        df = df[df[cols_check].notna().any(axis=1)]
        if 'Id Comisión' in df.columns:
            df = df[df['Id Comisión'].notna()]
    return df


def es_columna_generica(col):
    """Retorna True si el nombre de columna es genérico (Col_0, Unnamed, etc.)."""
    s = str(col).strip().lower()
    return s.startswith('col_') or s.startswith('unnamed') or not s


def es_valor_fantasma(val):
    """Retorna True si el valor es un placeholder / dato fantasma."""
    if pd.isna(val):
        return True
    s = str(val).strip().lower()
    return s in PLACEHOLDER_VALS


def validar_tabla(headers, filas):
    """Valida que una tabla tenga datos reales — no fantasma."""
    if not headers or not filas:
        return False
    valid_headers = [h for h in headers if not es_columna_generica(h)]
    if len(valid_headers) < 1:
        return False
    filas_validas = 0
    for fila in filas:
        non_empty = [v for v in fila if not es_valor_fantasma(v)]
        if len(non_empty) >= 1:
            filas_validas += 1
    return filas_validas >= 1


def validar_grafica(labels, valores):
    """Valida que una gráfica tenga datos reales."""
    if not labels or not valores:
        return False
    if len(labels) != len(valores):
        return False
    # Filtrar labels fantasma
    pares_validos = []
    for label, value in zip(labels, valores):
        numeric_value = parse_numeric_value(value)
        if es_valor_fantasma(label) or numeric_value is None:
            continue
        pares_validos.append((label, numeric_value))
    positivos = [value for _, value in pares_validos if value > 0]
    if len(positivos) >= 2:
        return True
    return len(pares_validos) >= 2 and any(value != 0 for _, value in pares_validos)


def limpiar_serie_categorica(serie):
    """Limpia una serie categórica eliminando placeholders."""
    serie_clean = serie.astype(str).str.strip()
    return serie_clean[~serie_clean.str.lower().isin(PLACEHOLDER_VALS)]


def _normalize_numeric_text(text):
    text = str(text).strip()
    if not text:
        return None, 1.0

    lowered = text.lower().replace('\xa0', ' ')
    if lowered in PLACEHOLDER_VALS:
        return None, 1.0

    negative = False
    if lowered.startswith('(') and lowered.endswith(')'):
        negative = True
        lowered = lowered[1:-1].strip()
    lowered = lowered.replace('−', '-').replace('–', '-').replace('—', '-')
    if lowered.startswith('-'):
        negative = True
        lowered = lowered[1:].strip()

    multiplier = 1.0
    if re.search(r'\b(mil\s*mm|mil\s*millones|bn|billones?)\b', lowered):
        multiplier = 1_000_000_000.0
    elif re.search(r'\b(mm|millones?)\b', lowered) or re.search(r'\d\s*m\b', lowered):
        multiplier = 1_000_000.0
    elif re.search(r'\b(k|mil)\b', lowered):
        multiplier = 1_000.0

    lowered = re.sub(r'(?i)\b(cop|usd|eur|pesos?|dolares?|moneda|aprox|aproximado|estimado)\b', '', lowered)
    cleaned = re.sub(r'[^0-9,.\-]', '', lowered)
    if not cleaned or not re.search(r'\d', cleaned):
        return None, multiplier

    cleaned = cleaned.lstrip('+')
    cleaned = re.sub(r'(?<!^)-', '', cleaned)

    if ',' in cleaned and '.' in cleaned:
        if cleaned.rfind(',') > cleaned.rfind('.'):
            cleaned = cleaned.replace('.', '').replace(',', '.')
        else:
            cleaned = cleaned.replace(',', '')
    elif ',' in cleaned:
        parts = cleaned.split(',')
        if len(parts) == 2:
            left, right = parts
            if len(right) <= 2:
                cleaned = f'{left}.{right}'
            elif len(right) == 3:
                cleaned = f'{left}{right}'
            else:
                cleaned = ''.join(parts)
        else:
            cleaned = ''.join(parts)
    elif '.' in cleaned:
        parts = cleaned.split('.')
        if len(parts) == 2:
            left, right = parts
            if len(right) <= 2:
                cleaned = f'{left}.{right}'
            elif len(right) == 3:
                cleaned = f'{left}{right}'
        elif len(parts) > 2:
            if len(parts[-1]) <= 2:
                cleaned = ''.join(parts[:-1]) + '.' + parts[-1]
            else:
                cleaned = ''.join(parts)

    if negative and cleaned and not cleaned.startswith('-'):
        cleaned = f'-{cleaned}'
    return cleaned, multiplier


def parse_numeric_value(value, kind_hint=None):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)):
        if pd.isna(value):
            return None
        return float(value)

    normalized, multiplier = _normalize_numeric_text(value)
    if normalized is None:
        return None
    try:
        numeric_value = float(normalized) * multiplier
    except Exception:
        return None

    return numeric_value


def normalize_numeric_series(series, kind_hint=None):
    return series.apply(lambda value: parse_numeric_value(value, kind_hint))


def normalize_semantic_text(value):
    text = str(value or '').replace('\xa0', ' ').strip().lower()
    if not text:
        return ''
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r'[^a-z0-9]+', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()


def unique_non_empty_texts(values, limit=None):
    seen = set()
    result = []
    for value in values or []:
        text = str(value or '').strip()
        if not text:
            continue
        key = normalize_semantic_text(text)
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(text)
        if limit is not None and len(result) >= limit:
            break
    return result


def build_sheet_semantic_signature(name, df=None):
    parts = [normalize_semantic_text(name)]
    if df is None or getattr(df, 'empty', True):
        return " ".join([part for part in parts if part])

    for col in list(df.columns)[:10]:
        normalized = normalize_semantic_text(col)
        if normalized:
            parts.append(normalized)

    sampled_rows = df.head(4).fillna('').values.tolist()
    for row in sampled_rows:
        row_tokens = []
        for value in row[:4]:
            normalized = normalize_semantic_text(value)
            if normalized:
                row_tokens.append(normalized)
        if row_tokens:
            parts.append(" ".join(row_tokens))

    return " ".join([part for part in parts if part])


def classify_sheet_family(name, df=None):
    signature = build_sheet_semantic_signature(name, df)
    if not signature:
        return 'general'

    if 'coso' in signature or ('componente' in signature and 'accion recomendada' in signature):
        return 'coso'
    if 'hallazgo' in signature:
        return 'hallazgos'
    if 'oportunidad de mejora' in signature or 'oportunidades de mejora' in signature or 'mejora' in signature:
        return 'oportunidades'
    if 'matriz de riesgos' in signature or ('riesgo' in signature and 'causa' in signature and 'consecuencia' in signature):
        return 'matriz_riesgos'
    if 'check list' in signature or ('prueba de auditoria' in signature and 'cumple' in signature):
        return 'checklist'
    if 'fto de arqueo' in signature or 'acta de arqueo' in signature or 'formato de arqueo' in signature or 'arqueo de caja' in signature:
        return 'arqueo'
    if 'soportes evidencias' in signature or ('soporte' in signature and 'evidencia' in signature):
        return 'evidencias'
    if ('auditoria' in signature and any(token in signature for token in ['pregunta', 'criterio', 'verificacion', 'revision'])) or (
        'controles existentes' in signature and ('preguntas' in signature or 'revision verificacion' in signature)
    ):
        return 'auditoria'
    if 'procedimiento' in signature or 'politica formal' in signature or 'finalidad del fondo' in signature:
        return 'procedimiento'
    if 'preguntas' in signature or 'cuestionario' in signature:
        return 'cuestionario'
    if 'distribucion' in signature or re.search(r'\btd\b', signature):
        return 'distribucion'
    return 'general'


def build_workbook_profile(sheets):
    families = {}
    for name, df in (sheets or {}).items():
        families[name] = classify_sheet_family(name, df)

    audit_families = {
        'auditoria',
        'checklist',
        'coso',
        'hallazgos',
        'oportunidades',
        'matriz_riesgos',
        'evidencias',
        'arqueo',
        'procedimiento',
        'cuestionario',
    }
    detected_audit = [family for family in families.values() if family in audit_families]
    workbook_type = 'auditoria_control' if len(detected_audit) >= 3 else 'general'
    family_labels = [
        SHEET_FAMILY_LABELS.get(family, family)
        for family in unique_non_empty_texts(families.values())
        if family != 'general'
    ]

    conclusions = []
    insights = []
    if workbook_type == 'auditoria_control':
        conclusions.append(
            f"El archivo corresponde a una revision de auditoria y control con {len(sheets)} hojas funcionales."
        )
        if family_labels:
            conclusions.append(
                f"Se identifican frentes de {', '.join(family_labels[:6])}, por lo que la lectura debe priorizar hallazgos, riesgos y acciones."
            )
        insights.append(
            "El libro se comporta como expediente de auditoria: conviene resumir controles, brechas y recomendaciones antes que forzar metricas financieras."
        )
        if 'matriz_riesgos' in families.values():
            insights.append("La matriz de riesgos debe usarse como fuente de exposicion y mitigacion para la narrativa ejecutiva.")
        if 'checklist' in families.values():
            insights.append("El checklist aporta cobertura de pruebas y cumplimiento; es mejor sintetizar brechas que mostrar toda la tabla cruda.")
        if 'coso' in families.values():
            insights.append("La evaluacion COSO debe escalar componentes pendientes o no evaluados, incluso si el estado viene incompleto.")

    return {
        'tipo_libro': workbook_type,
        'familias_por_hoja': families,
        'familias_detectadas': family_labels,
        'conclusiones': unique_non_empty_texts(conclusions, limit=4),
        'insights': unique_non_empty_texts(insights, limit=6),
    }


def preferred_keywords_for_sheet_family(sheet_family):
    mapping = {
        'auditoria': ['control', 'criterio', 'pregunta', 'revision', 'verificacion', 'observacion'],
        'checklist': ['prueba', 'cumple', 'no cumple', 'observacion'],
        'hallazgos': ['hallazgo', 'riesgo', 'plan', 'accion', 'estado', 'evidencia'],
        'oportunidades': ['oportunidad', 'estado', 'observacion', 'control', 'riesgo'],
        'matriz_riesgos': ['riesgo', 'causa', 'consecuencia', 'control', 'recomendacion'],
        'evidencias': ['evidencia', 'soporte', 'documento', 'observacion'],
        'arqueo': ['fecha', 'responsable', 'auditor', 'valor', 'saldo'],
        'procedimiento': ['pregunta', 'respuesta', 'control', 'observacion'],
        'cuestionario': ['pregunta', 'respuesta', 'observacion', 'estado'],
        'coso': ['componente', 'item', 'estado', 'accion'],
    }
    return mapping.get(sheet_family, [])


def select_semantic_columns(df, sheet_family='general', max_cols=MAX_TABLE_COLS):
    valid_cols = [c for c in df.columns if not es_columna_generica(c)]
    if not valid_cols:
        valid_cols = df.columns.tolist()
    preferred = [normalize_semantic_text(keyword) for keyword in preferred_keywords_for_sheet_family(sheet_family)]
    selected = []
    for col in valid_cols:
        normalized = normalize_semantic_text(col)
        if preferred and any(keyword in normalized for keyword in preferred):
            selected.append(col)
    for col in valid_cols:
        if col not in selected:
            selected.append(col)
    return selected[:max_cols]


def build_table_from_dataframe(df, sheet_name, sheet_family='general', max_cols=MAX_TABLE_COLS, max_rows=30, text_limit=150, min_meaningful_cells=1):
    if df is None or df.empty:
        return None

    selected_cols = select_semantic_columns(df, sheet_family=sheet_family, max_cols=max_cols)
    if not selected_cols:
        return None

    df_res = df[selected_cols].head(max_rows).copy()
    for col in df_res.columns:
        if pd.api.types.is_string_dtype(df_res[col]) or df_res[col].dtype == object:
            df_res[col] = df_res[col].astype(str).str[:text_limit]

    mask = df_res.apply(
        lambda row: sum(1 for value in row if not es_valor_fantasma(value)) >= min_meaningful_cells,
        axis=1
    )
    df_res = df_res[mask]
    if df_res.empty:
        return None

    tabla_info = {
        'encabezados': [str(col) for col in selected_cols],
        'filas': df_res.values.tolist(),
        'hoja_origen': sheet_name,
        'sheet_family': sheet_family,
    }
    if validar_tabla(tabla_info['encabezados'], tabla_info['filas']):
        return tabla_info
    return None


def is_probable_numeric_identifier(values):
    numeric_values = [float(value) for value in values if value is not None]
    if len(numeric_values) < 4:
        return False
    integers = [value for value in numeric_values if float(value).is_integer()]
    if len(integers) < max(4, int(len(numeric_values) * 0.8)):
        return False
    ordered = integers[: min(10, len(integers))]
    deltas = [ordered[i + 1] - ordered[i] for i in range(len(ordered) - 1)]
    if not deltas:
        return False
    return len({round(delta, 6) for delta in deltas}) <= 2 and max(abs(delta) for delta in deltas) <= 10


def compactar_categorias(labels, valores, max_items=MAX_BAR_CATEGORIES, otros_label='Otros'):
    """Compacta categorías menores para evitar gráficas sobrecargadas."""
    pares = []
    for label, value in zip(labels, valores):
        numeric_value = parse_numeric_value(value)
        if numeric_value is None:
            continue
        label_text = str(label).strip()
        if not label_text or numeric_value <= 0:
            continue
        pares.append((label_text[:30], numeric_value))

    if len(pares) <= max_items:
        return [label for label, _ in pares], [value for _, value in pares]

    principales = pares[:max_items - 1]
    restantes = pares[max_items - 1:]
    total_otros = sum(value for _, value in restantes)
    if total_otros > 0:
        principales.append((otros_label, total_otros))

    return [label for label, _ in principales], [value for _, value in principales]


# ═══════════════════════════════════════════════════════════════════════════════
# SCORING Y PRIORIZACIÓN
# ═══════════════════════════════════════════════════════════════════════════════

def score_sheet_for_primary(name, df):
    name_l = str(name).lower()
    rows, cols = df.shape
    score = rows * cols
    if cols < 2 or rows < 2:
        return -1
    family = classify_sheet_family(name, df)
    score += PRIMARY_SHEET_FAMILY_SCORES.get(family, 0)
    if any(k in name_l for k in ['muestra total', 'base de comisiones', 'consolidado total']):
        score += 500
    if any(k in name_l for k in ['ventas', 'inventario', 'stock', 'productos', 'resumen', 'dashboard', 'datos']):
        score += 180
    if any(k in name_l for k in ['hallazgo', 'oportunidad', 'mejora', 'coso', 'td', 'distribucion']):
        score -= 200
    headers = [str(c).strip().lower() for c in df.columns if pd.notna(c)]
    if any(h in headers for h in ['solicitante', 'valor total solicitado', 'ventas_totales', 'precio_venta', 'stock']):
        score += 120
    return score


# ═══════════════════════════════════════════════════════════════════════════════
# ANÁLISIS ESTADÍSTICO INTELIGENTE
# ═══════════════════════════════════════════════════════════════════════════════

def analizar_columna_numerica(df, col):
    """Analiza una columna numérica y retorna estadísticas reales."""
    profile = build_numeric_analysis_profile(df, col)
    serie = profile['series_valid']
    if len(serie) < 2:
        return None
    return {
        'columna': str(col),
        'kind': profile.get('kind', 'number'),
        'total': float(serie.sum()),
        'promedio': float(serie.mean()),
        'mediana': float(serie.median()),
        'minimo': float(serie.min()),
        'maximo': float(serie.max()),
        'desv_std': float(serie.std()),
        'conteo': int(len(serie)),
        'sin_datos': int(len(df) - len(serie)),
        'q1': float(serie.quantile(0.25)),
        'q3': float(serie.quantile(0.75)),
        'resolved_ratio': round(float(profile.get('resolved_ratio') or 0), 3),
        'mixed_currency': bool(profile.get('mixed_currency')),
        'currencies_detected': profile.get('currencies_detected') or ['COP'],
        'unresolved_rows': int(profile.get('unresolved_rows') or 0),
    }


def es_columna_identificador(col):
    nombre = str(col).lower()
    return any(token in nombre for token in ['id', 'codigo', 'consecutivo', 'numero', 'nro', 'folio', 'radicado'])


def es_columna_persona(col):
    nombre = str(col).lower()
    return any(token in nombre for token in ['nombre', 'solicitante', 'responsable', 'cliente', 'proveedor', 'empleado', 'usuario', 'colaborador'])


def es_dimension_ejecutiva(col):
    nombre = str(col).lower()
    return any(token in nombre for token in ['estado', 'status', 'tipo', 'categoria', 'mes', 'ciudad', 'destino', 'centro', 'control', 'riesgo', 'hallazgo', 'proceso'])


def infer_numeric_kind(label=None):
    normalized = str(label or '').strip().lower()
    if any(token in normalized for token in ['porcentaje', 'avance', 'cumplimiento', '%', 'ratio', 'participacion', 'share', 'margen']):
        return 'percent'
    if any(token in normalized for token in ['valor', 'monto', 'total', 'costo', 'precio', 'ingreso', 'venta', 'gasto', 'importe', 'cop', 'peso', 'pesos', 'moneda', 'tarifa']):
        return 'currency'
    return 'number'


def normalize_currency_code(value):
    text = str(value or '').strip().upper()
    if not text:
        return None
    if any(token in text for token in ['COP', 'PESO', 'PESOS', 'COL$', 'CO$']):
        return 'COP'
    if any(token in text for token in ['USD', 'US$', 'DOLAR', 'DÓLAR']):
        return 'USD'
    if any(token in text for token in ['EUR', 'EURO']):
        return 'EUR'
    if 'GBP' in text or 'LIBRA' in text:
        return 'GBP'
    return None


def is_currency_header(name):
    normalized = str(name or '').strip().lower()
    return any(token in normalized for token in ['moneda', 'divisa', 'currency', 'tipo moneda'])


def is_exchange_rate_header(name):
    normalized = str(name or '').strip().lower()
    return any(token in normalized for token in ['trm', 'tasa', 'tipo cambio', 'exchange rate', 'fx', 'conversion'])


def find_currency_support_columns(df):
    currency_col = None
    rate_col = None
    for col in df.columns:
        if currency_col is None and is_currency_header(col):
            currency_col = col
        if rate_col is None and is_exchange_rate_header(col):
            rate_col = col
    return currency_col, rate_col


def build_financial_series(df, value_col):
    series = normalize_numeric_series(df[value_col], value_col)
    currency_col, rate_col = find_currency_support_columns(df)
    currency_codes = []
    converted = []
    converted_rows = 0
    unresolved_rows = 0

    if currency_col is None:
        return {
            'series_raw': series,
            'series': series.fillna(0),
            'series_valid': series.dropna(),
            'currency_column': None,
            'rate_column': None,
            'currencies_detected': ['COP'],
            'mixed_currency': False,
            'conversion_applied': False,
            'resolved_ratio': 1.0 if len(series) else 0.0,
            'unresolved_rows': 0,
        }

    currency_series = df[currency_col].apply(normalize_currency_code)
    rate_series = normalize_numeric_series(df[rate_col], rate_col) if rate_col is not None else pd.Series([None] * len(df), index=df.index)

    for index, amount in series.items():
        if amount is None or pd.isna(amount):
            converted.append(np.nan)
            continue
        currency_code = currency_series.get(index)
        if currency_code:
            currency_codes.append(currency_code)
        if currency_code in (None, 'COP'):
            converted.append(float(amount))
            converted_rows += 1
            continue
        rate_value = rate_series.get(index) if rate_col is not None else None
        if rate_value is not None and not pd.isna(rate_value) and rate_value > 0:
            converted.append(float(amount) * float(rate_value))
            converted_rows += 1
        else:
            converted.append(np.nan)
            unresolved_rows += 1

    converted_series = pd.Series(converted, index=series.index, dtype='float64')
    unique_currencies = sorted({code for code in currency_codes if code})
    return {
        'series_raw': converted_series,
        'series': converted_series.fillna(0),
        'series_valid': converted_series.dropna(),
        'currency_column': currency_col,
        'rate_column': rate_col,
        'currencies_detected': unique_currencies or ['COP'],
        'mixed_currency': len(unique_currencies) > 1,
        'conversion_applied': rate_col is not None and converted_rows > 0,
        'resolved_ratio': converted_rows / max(1, int(series.notna().sum())),
        'unresolved_rows': unresolved_rows,
    }


def build_numeric_analysis_profile(df, col):
    empty_series = pd.Series(np.nan, index=df.index, dtype='float64')
    if col not in df.columns:
        return {
            'kind': infer_numeric_kind(col),
            'series_raw': empty_series,
            'series_valid': empty_series.dropna(),
            'mixed_currency': False,
            'conversion_applied': False,
            'resolved_ratio': 0.0,
            'unresolved_rows': 0,
            'currencies_detected': ['COP'],
            'currency_column': None,
            'rate_column': None,
        }

    metric_kind = infer_numeric_kind(col)
    if metric_kind == 'currency':
        financial = build_financial_series(df, col)
        series_raw = pd.to_numeric(financial.get('series_raw'), errors='coerce')
        return {
            'kind': metric_kind,
            'series_raw': series_raw,
            'series_valid': series_raw.dropna(),
            'mixed_currency': bool(financial.get('mixed_currency')),
            'conversion_applied': bool(financial.get('conversion_applied')),
            'resolved_ratio': float(financial.get('resolved_ratio') or 0),
            'unresolved_rows': int(financial.get('unresolved_rows') or 0),
            'currencies_detected': financial.get('currencies_detected') or ['COP'],
            'currency_column': financial.get('currency_column'),
            'rate_column': financial.get('rate_column'),
        }

    series = pd.to_numeric(normalize_numeric_series(df[col], col), errors='coerce')
    total_numeric = int(series.notna().sum())
    return {
        'kind': metric_kind,
        'series_raw': series,
        'series_valid': series.dropna(),
        'mixed_currency': False,
        'conversion_applied': False,
        'resolved_ratio': 1.0 if total_numeric else 0.0,
        'unresolved_rows': 0,
        'currencies_detected': ['COP'],
        'currency_column': None,
        'rate_column': None,
    }


def should_trust_numeric_profile(profile, min_resolved_ratio=0.7):
    if not profile:
        return False
    if profile.get('kind') != 'currency':
        return True
    if not profile.get('mixed_currency'):
        return True
    return float(profile.get('resolved_ratio') or 0) >= min_resolved_ratio


def build_financial_context(df, cols_info):
    money_columns = [c['nombre'] for c in cols_info if c.get('tipo') == 'numerico' and infer_numeric_kind(c.get('nombre')) == 'currency' and c['nombre'] in df.columns]
    details = []
    currencies_detected = set()
    has_unresolved = False
    for col_name in money_columns[:6]:
        profile = build_financial_series(df, col_name)
        currencies_detected.update(profile.get('currencies_detected') or [])
        has_unresolved = has_unresolved or profile.get('unresolved_rows', 0) > 0
        details.append({
            'columna': col_name,
            'currency_column': profile.get('currency_column'),
            'rate_column': profile.get('rate_column'),
            'currencies_detected': profile.get('currencies_detected') or ['COP'],
            'mixed_currency': bool(profile.get('mixed_currency')),
            'conversion_applied': bool(profile.get('conversion_applied')),
            'resolved_ratio': round(float(profile.get('resolved_ratio') or 0), 3),
            'unresolved_rows': int(profile.get('unresolved_rows') or 0),
        })
    return {
        'currency_columns': details,
        'currencies_detected': sorted(currencies_detected) or ['COP'],
        'has_mixed_currency': len(currencies_detected) > 1,
        'has_unresolved_conversion': has_unresolved,
    }


def build_data_quality_profile(df, cols_info):
    total_rows = int(len(df))
    total_cells = int(df.shape[0] * df.shape[1])
    placeholder_cells = 0
    for col in df.columns:
        serie = df[col]
        if serie.dtype == object or str(serie.dtype) == 'string':
            placeholder_cells += int(serie.astype(str).str.strip().str.lower().isin(PLACEHOLDER_VALS).sum())

    duplicate_rows = int(df.astype(str).duplicated().sum()) if total_rows else 0
    sparse_columns = []
    for col_info in cols_info:
        stats = col_info.get('stats') or {}
        if col_info.get('tipo') == 'numerico' and stats:
            missing_pct = stats.get('sin_datos', 0) / max(1, total_rows)
            if missing_pct >= 0.2:
                sparse_columns.append({
                    'columna': col_info['nombre'],
                    'missing_ratio': round(float(missing_pct), 3),
                })

    currency_alerts = []
    for col_info in cols_info:
        if col_info.get('tipo') != 'numerico' or infer_numeric_kind(col_info.get('nombre')) != 'currency':
            continue
        profile = build_numeric_analysis_profile(df, col_info['nombre'])
        if profile.get('mixed_currency') or profile.get('unresolved_rows'):
            currency_alerts.append({
                'columna': col_info['nombre'],
                'currencies_detected': profile.get('currencies_detected') or ['COP'],
                'resolved_ratio': round(float(profile.get('resolved_ratio') or 0), 3),
                'unresolved_rows': int(profile.get('unresolved_rows') or 0),
            })

    placeholder_ratio = placeholder_cells / max(1, total_cells)
    duplicate_ratio = duplicate_rows / max(1, total_rows)
    quality_score = max(0.0, 1.0 - (placeholder_ratio * 0.55) - (duplicate_ratio * 0.45))
    return {
        'total_rows': total_rows,
        'total_cells': total_cells,
        'placeholder_cells': int(placeholder_cells),
        'placeholder_ratio': round(float(placeholder_ratio), 3),
        'duplicate_rows': duplicate_rows,
        'duplicate_ratio': round(float(duplicate_ratio), 3),
        'sparse_numeric_columns': sparse_columns[:6],
        'currency_alerts': currency_alerts[:6],
        'quality_score': round(float(quality_score), 3),
    }


def detectar_columnas_importantes(df):
    """Detecta cuáles columnas son las más importantes en el DataFrame."""
    cols_info = []
    for col in df.columns:
        if es_columna_generica(col):
            continue
        
        info = {'nombre': str(col), 'tipo': 'texto', 'importancia': 0}
        serie = df[col].dropna()
        if len(serie) == 0:
            continue
        
        col_lower = str(col).lower()
        
        # Detectar tipo
        serie_num = normalize_numeric_series(serie, col)
        ratio_num = serie_num.notna().sum() / max(1, len(serie))
        
        if ratio_num >= 0.6:
            info['tipo'] = 'numerico'
            info['importancia'] += 20
            stats = analizar_columna_numerica(df, col)
            if stats:
                info['stats'] = stats
                if is_probable_numeric_identifier(serie_num.dropna().tolist()) and not es_dimension_ejecutiva(col):
                    info['tipo'] = 'identificador'
                    info['importancia'] -= 18
                if stats['total'] > 1000000:
                    info['importancia'] += 30
                elif stats['total'] > 10000:
                    info['importancia'] += 15
        else:
            unique_vals = serie.astype(str).nunique()
            total_vals = len(serie)
            ratio_unique = unique_vals / max(1, total_vals)
            
            if ratio_unique <= 0.3 and unique_vals <= 20:
                info['tipo'] = 'categorica'
                info['importancia'] += 15
                info['valores_unicos'] = unique_vals
            elif ratio_unique > 0.8:
                info['tipo'] = 'identificador'
                info['importancia'] += 5
            else:
                info['tipo'] = 'texto'
                info['importancia'] += 3
        
        # Bonus por keywords en nombre de columna
        if any(k in col_lower for k in ['total', 'valor', 'costo', 'precio', 'monto', 'gasto', 'ingreso', 'venta']):
            info['importancia'] += 25
        if any(k in col_lower for k in ['estado', 'status', 'tipo', 'categoria']):
            info['importancia'] += 20
        if any(k in col_lower for k in ['nombre', 'solicitante', 'responsable', 'cliente', 'proveedor']):
            info['importancia'] -= 10
        if any(k in col_lower for k in ['fecha', 'date', 'periodo', 'mes', 'año']):
            info['importancia'] += 10
        if any(k in col_lower for k in ['id', 'codigo', 'folio', 'numero']):
            info['importancia'] -= 12
        if any(k in col_lower for k in ['porcentaje', 'avance', '%', 'cumplimiento']):
            info['importancia'] += 18
        if es_dimension_ejecutiva(col):
            info['importancia'] += 12
        if es_columna_identificador(col):
            info['importancia'] -= 10
        if es_columna_persona(col) and info['tipo'] in ('texto', 'identificador'):
            info['importancia'] -= 8
            
        cols_info.append(info)
    
    return sorted(cols_info, key=lambda x: x['importancia'], reverse=True)


# ═══════════════════════════════════════════════════════════════════════════════
# ANÁLISIS AVANZADO: OUTLIERS, PARETO, CORRELACIONES
# ═══════════════════════════════════════════════════════════════════════════════

def detectar_outliers(df, col):
    """Detecta outliers usando el método IQR (Rango Intercuartílico)."""
    profile = build_numeric_analysis_profile(df, col)
    if not should_trust_numeric_profile(profile):
        return None
    serie = profile['series_valid']
    if len(serie) < 8:
        return None
    Q1 = float(serie.quantile(0.25))
    Q3 = float(serie.quantile(0.75))
    IQR = Q3 - Q1
    if IQR <= 0:
        return None
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    outliers = serie[(serie < lower) | (serie > upper)]
    if len(outliers) == 0:
        return None
    return {
        'columna': str(col),
        'total_outliers': int(len(outliers)),
        'pct_outliers': round(len(outliers) / len(serie) * 100, 1),
        'rango_normal': [round(lower, 2), round(upper, 2)],
        'valor_min_outlier': float(outliers.min()),
        'valor_max_outlier': float(outliers.max()),
        'ejemplos': [float(v) for v in outliers.nlargest(3).tolist()]
    }


def analisis_pareto(df, col_cat, col_num=None):
    """Análisis de concentración Pareto (80/20) sobre una columna categórica."""
    if col_cat not in df.columns:
        return None
    
    serie_cat = limpiar_serie_categorica(df[col_cat])
    if len(serie_cat) < 3:
        return None
    
    if col_num and col_num in df.columns:
        df_temp = df.loc[serie_cat.index].copy()
        metric_profile = build_numeric_analysis_profile(df_temp, col_num)
        if not should_trust_numeric_profile(metric_profile):
            return None
        df_temp[col_num] = metric_profile['series_raw']
        df_temp = df_temp.dropna(subset=[col_num])
        grouped = df_temp.groupby(col_cat)[col_num].sum().sort_values(ascending=False)
        grouped = grouped[grouped > 0]
    else:
        grouped = serie_cat.value_counts()
    
    if len(grouped) < 3:
        return None
    
    total = grouped.sum()
    if total <= 0:
        return None
    
    cumsum = grouped.cumsum()
    cumsum_pct = (cumsum / total * 100).round(1)
    
    n_80 = int((cumsum_pct <= 80).sum()) + 1
    n_80 = min(n_80, len(grouped))
    pct_cat_80 = round(n_80 / len(grouped) * 100, 0)
    
    top_items = []
    for i, (cat, val) in enumerate(grouped.head(5).items()):
        top_items.append({
            'categoria': str(cat)[:40],
            'valor': float(val),
            'pct': round(val / total * 100, 1),
            'pct_acumulado': float(cumsum_pct.iloc[i]) if i < len(cumsum_pct) else 100.0
        })
    
    concentracion = 'alta' if pct_cat_80 <= 25 else 'moderada' if pct_cat_80 <= 50 else 'dispersa'
    
    return {
        'columna_categoria': str(col_cat),
        'columna_valor': str(col_num) if col_num else 'conteo',
        'total_categorias': int(len(grouped)),
        'categorias_80_pct': n_80,
        'pct_categorias_para_80': pct_cat_80,
        'top_items': top_items,
        'concentracion': concentracion,
        'lider': str(grouped.index[0])[:40] if len(grouped) > 0 else '',
        'lider_pct': round(float(grouped.iloc[0]) / total * 100, 1) if len(grouped) > 0 else 0
    }


def detectar_correlaciones(df, cols_info):
    """Detecta correlaciones significativas entre columnas numéricas."""
    # Excluir columnas de fecha (fecha x fecha siempre da ~1.0, no es útil)
    fecha_keywords = ['fecha', 'date', 'periodo', 'mes', 'año', 'dia', 'day', 'month', 'year']
    cols_num = [c['nombre'] for c in cols_info 
                if c['tipo'] == 'numerico' and c['nombre'] in df.columns 
                and 'stats' in c and c['stats']['conteo'] >= 5
                and not any(k in str(c['nombre']).lower() for k in fecha_keywords)]
    if len(cols_num) < 2:
        return []
    
    correlaciones = []
    profile_cache = {}
    for i in range(len(cols_num)):
        for j in range(i+1, len(cols_num)):
            col_a, col_b = cols_num[i], cols_num[j]
            if col_a not in profile_cache:
                profile_cache[col_a] = build_numeric_analysis_profile(df, col_a)
            if col_b not in profile_cache:
                profile_cache[col_b] = build_numeric_analysis_profile(df, col_b)
            if not should_trust_numeric_profile(profile_cache[col_a]) or not should_trust_numeric_profile(profile_cache[col_b]):
                continue
            sa = profile_cache[col_a]['series_raw']
            sb = profile_cache[col_b]['series_raw']
            valid = sa.notna() & sb.notna()
            n = int(valid.sum())
            if n < 5:
                continue
            corr = float(sa[valid].corr(sb[valid]))
            if pd.isna(corr) or abs(corr) < 0.5:
                continue
            tipo = ('positiva fuerte' if corr >= 0.8 else 
                    'positiva moderada' if corr >= 0.5 else 
                    'negativa fuerte' if corr <= -0.8 else 'negativa moderada')
            correlaciones.append({
                'col_a': str(col_a),
                'col_b': str(col_b),
                'correlacion': round(corr, 3),
                'tipo': tipo,
                'n_observaciones': n
            })
    
    return sorted(correlaciones, key=lambda x: abs(x['correlacion']), reverse=True)[:5]


def detectar_tendencia_temporal(df, cols_info):
    """Detecta si hay una columna de fecha y analiza tendencias temporales."""
    col_fecha = None
    for c in cols_info:
        if c['nombre'] in df.columns:
            col_lower = str(c['nombre']).lower()
            if any(k in col_lower for k in ['fecha', 'date', 'periodo', 'mes']):
                col_fecha = c['nombre']
                break
    
    if not col_fecha:
        return None
    
    cols_num = [c['nombre'] for c in cols_info 
                if c['tipo'] == 'numerico' and c['nombre'] in df.columns 
                and 'stats' in c and c['stats']['conteo'] >= 3]
    if not cols_num:
        return None
    
    col_val = cols_num[0]
    
    try:
        df_temp = df[[col_fecha, col_val]].copy()
        metric_profile = build_numeric_analysis_profile(df_temp, col_val)
        if not should_trust_numeric_profile(metric_profile):
            return None
        df_temp[col_val] = metric_profile['series_raw']
        df_temp = df_temp.dropna()
        if len(df_temp) < 3:
            return None
        
        # Try to parse dates
        try:
            df_temp['_fecha'] = pd.to_datetime(df_temp[col_fecha], errors='coerce')
        except:
            return None
        
        df_temp = df_temp.dropna(subset=['_fecha'])
        if len(df_temp) < 3:
            return None
        
        df_temp = df_temp.sort_values('_fecha')
        vals = df_temp[col_val].values
        
        # Build monthly series for chart
        df_temp['_mes'] = df_temp['_fecha'].dt.to_period('M')
        mensual = df_temp.groupby('_mes')[col_val].sum().sort_index()
        
        if len(mensual) < 2:
            return None
            
        vals_mensual = mensual.values
        n_meses = len(vals_mensual)
        
        # Calculate trend using the monthly series (more stable than raw records)
        mitad = max(1, n_meses // 2)
        avg_inicio = float(np.mean(vals_mensual[:mitad]))
        avg_fin = float(np.mean(vals_mensual[-mitad:]))
        
        if avg_inicio == 0:
            cambio_pct = 0
        else:
            cambio_pct = round((avg_fin - avg_inicio) / abs(avg_inicio) * 100, 1)
        
        # Evaluate stability and trend using the monthly variation
        tendencia = 'creciente' if cambio_pct > 10 else 'decreciente' if cambio_pct < -10 else 'estable'
        
        serie_labels = [str(p) for p in mensual.index[-MAX_CHART_CATEGORIES:]]
        serie_valores = [float(v) for v in mensual.values[-MAX_CHART_CATEGORIES:]]
        
        return {
            'columna_fecha': str(col_fecha),
            'columna_valor': str(col_val),
            'tendencia': tendencia,
            'cambio_pct': cambio_pct,
            'promedio_inicio': round(avg_inicio, 2),
            'promedio_fin': round(avg_fin, 2),
            'serie_temporal': {
                'labels': serie_labels,
                'valores': serie_valores
            } if len(serie_labels) >= 2 else None
        }
    except:
        return None


def generar_insights_avanzados(df, cols_info, paretos, outliers_list, correlaciones, tendencia):
    """Genera insights de alto nivel combinando todos los análisis avanzados."""
    insights = []
    total_filas = len(df)
    
    # 1. Insights de concentración (Pareto)
    for p in (paretos or []):
        if p and p['concentracion'] in ('alta', 'moderada') and p.get('top_items'):
            top = p['top_items'][0]
            insights.append({
                'tipo': 'concentracion',
                'importancia': 95 if p['concentracion'] == 'alta' else 80,
                'texto': (f"Alta concentración: '{top['categoria']}' representa el {top['pct']}% "
                         f"del total de {p['columna_valor']}. Solo {p['categorias_80_pct']} de "
                         f"{p['total_categorias']} categorías acumulan el 80% del valor."),
                'accion': (f"Focalizar control sobre las {p['categorias_80_pct']} categorías "
                          f"dominantes de '{p['columna_categoria']}'.")
            })
    
    # 2. Insights de outliers (anomalías)
    for o in (outliers_list or []):
        if o and o['total_outliers'] >= 1:
            insights.append({
                'tipo': 'anomalia',
                'importancia': 88,
                'texto': (f"Se detectaron {o['total_outliers']} valores atípicos en "
                         f"'{o['columna']}' ({o['pct_outliers']}% de los datos). "
                         f"El rango normal va hasta {format_number(o['rango_normal'][1])}, "
                         f"pero hay valores hasta {format_number(o['valor_max_outlier'])}."),
                'accion': (f"Investigar los {o['total_outliers']} registros atípicos de "
                          f"'{o['columna']}' para determinar si son errores o casos excepcionales.")
            })
    
    # 3. Insights de correlaciones
    for c in (correlaciones or [])[:2]:
        verbo = 'también crece' if c['correlacion'] > 0 else 'decrece'
        insights.append({
            'tipo': 'correlacion',
            'importancia': 75,
            'texto': (f"Correlación {c['tipo']} ({c['correlacion']:.2f}) entre '{c['col_a']}' "
                     f"y '{c['col_b']}': cuando una sube, la otra {verbo}."),
            'accion': (f"Considerar '{c['col_a']}' y '{c['col_b']}' como variables "
                      f"vinculadas en la toma de decisiones.")
        })
    
    # 4. Insight de tendencia temporal
    if tendencia and tendencia.get('tendencia') != 'estable':
        emoji = '📈' if tendencia['tendencia'] == 'creciente' else '📉'
        insights.append({
            'tipo': 'tendencia',
            'importancia': 85,
            'texto': (f"Tendencia {tendencia['tendencia']} en '{tendencia['columna_valor']}': "
                     f"cambio del {tendencia['cambio_pct']:+.1f}% entre inicio y fin del periodo."),
            'accion': (f"Monitorear la tendencia de '{tendencia['columna_valor']}' y "
                      f"proyectar impacto si continúa.")
        })
    
    # 5. Insight de distribución desbalanceada
    for col_info in cols_info:
        if col_info['tipo'] != 'categorica' or col_info['nombre'] not in df.columns:
            continue
        serie = limpiar_serie_categorica(df[col_info['nombre']])
        if len(serie) < 5:
            continue
        dist = serie.value_counts()
        if len(dist) >= 2:
            top_pct = dist.iloc[0] / len(serie) * 100
            if top_pct > 55:
                insights.append({
                    'tipo': 'desbalance',
                    'importancia': 72,
                    'texto': (f"En '{col_info['nombre']}', '{dist.index[0]}' domina con el "
                             f"{top_pct:.0f}% de los registros ({dist.iloc[0]:,} de {len(serie):,})."),
                    'accion': (f"Evaluar si la concentración del {top_pct:.0f}% en "
                              f"'{dist.index[0]}' refleja un patrón esperado o una anomalía.")
                })
                break
    
    # 6. Insight de calidad de datos
    cols_faltantes = []
    for c in cols_info:
        if c['tipo'] == 'numerico' and 'stats' in c:
            pct_miss = c['stats']['sin_datos'] / max(1, total_filas) * 100
            if pct_miss > 15:
                cols_faltantes.append((c['nombre'], round(pct_miss, 0)))
    if cols_faltantes:
        cols_str = ", ".join([f"'{n}' ({p:.0f}%)" for n, p in cols_faltantes[:3]])
        insights.append({
            'tipo': 'calidad_datos',
            'importancia': 60,
            'texto': f"Datos faltantes significativos en: {cols_str}.",
            'accion': "Verificar la completitud de datos antes de tomar decisiones sobre estas columnas."
        })
    
    insights.sort(key=lambda x: x['importancia'], reverse=True)
    return insights[:MAX_INSIGHTS_AVANZADOS]


# ═══════════════════════════════════════════════════════════════════════════════
# GENERACIÓN DE KPIs INTELIGENTES
# ═══════════════════════════════════════════════════════════════════════════════

def generar_kpis_automaticos(df, cols_info):
    """Genera KPIs automáticos con contexto analítico desde los datos reales."""
    kpis = []
    total_filas = len(df)
    
    # 1. Total de registros (siempre)
    kpis.append({
        'label': 'Total Registros',
        'value': f'{total_filas:,}',
        'importancia': 50,
        'contexto': f'Base de análisis sobre {total_filas:,} filas de datos'
    })
    
    # 2. KPIs de columnas numéricas importantes (con contexto)
    for col_info in cols_info:
        if col_info['tipo'] != 'numerico' or 'stats' not in col_info:
            continue
        stats = col_info['stats']
        col_name = col_info['nombre']
        col_lower = col_name.lower()
        
        if any(k in col_lower for k in ['total', 'valor', 'costo', 'precio', 'monto', 'ingreso', 'venta', 'gasto']):
            metric_kind = infer_numeric_kind(col_name)
            context_suffix = ""
            if metric_kind == 'currency' and stats.get('mixed_currency'):
                currencies = "/".join(stats.get('currencies_detected', ['COP'])[:3])
                resolved_pct = int(round((stats.get('resolved_ratio') or 0) * 100))
                context_suffix = f" | Monedas: {currencies} ({resolved_pct}% resuelto)"
            # KPI principal: Total del valor
            kpis.append({
                'label': f'Total {col_name[:25]}',
                'value': format_number(stats['total'], kind=metric_kind),
                'importancia': col_info['importancia'] + 15,
                'contexto': f"Promedio: {format_number(stats['promedio'], kind=metric_kind)} | Máximo: {format_number(stats['maximo'], kind=metric_kind)}{context_suffix}"
            })
            # KPI secundario: Promedio con referencia
            if stats['maximo'] > 0 and stats['total'] > 0:
                ratio_max = stats['maximo'] / stats['total'] * 100
                kpis.append({
                    'label': f'Promedio {col_name[:20]}',
                    'value': format_number(stats['promedio'], kind=metric_kind),
                    'importancia': col_info['importancia'] + 5,
                    'contexto': f"El máximo ({format_number(stats['maximo'], kind=metric_kind)}) es {ratio_max:.0f}% del total"
                })
        elif any(k in col_lower for k in ['porcentaje', 'avance', 'cumplimiento', '%']):
            kpis.append({
                'label': f'Promedio {col_name[:22]}',
                'value': f'{stats["promedio"]:.1f}%' if stats['promedio'] <= 1 else f'{stats["promedio"]:.1f}%',
                'importancia': col_info['importancia'] + 12,
                'contexto': f"Mínimo: {stats['minimo']:.1f}% | Máximo: {stats['maximo']:.1f}%"
            })
        elif stats['total'] > 0:
            kpis.append({
                'label': f'Suma {col_name[:25]}',
                'value': format_number(stats['total']),
                'importancia': col_info['importancia'],
                'contexto': f"Sobre {stats['conteo']:,} registros con datos"
            })
    
    # 3. KPIs de columnas categóricas (con distribución)
    for col_info in cols_info:
        if col_info['tipo'] != 'categorica':
            continue
        col_name = col_info['nombre']
        unique = col_info.get('valores_unicos', 0)
        if unique > 0 and col_name in df.columns:
            serie = limpiar_serie_categorica(df[col_name])
            if len(serie) > 0:
                top_val = serie.value_counts().index[0] if len(serie.value_counts()) > 0 else '—'
                top_pct = serie.value_counts().iloc[0] / len(serie) * 100 if len(serie) > 0 else 0
                kpis.append({
                    'label': f'{col_name[:25]}',
                    'value': f'{unique} tipos',
                    'importancia': col_info['importancia'],
                    'contexto': f"Líder: '{str(top_val)[:20]}' ({top_pct:.0f}%)"
                })
    
    kpis.sort(key=lambda x: x['importancia'], reverse=True)
    return kpis[:MAX_KPIS]


# ═══════════════════════════════════════════════════════════════════════════════
# GENERACIÓN DE GRÁFICAS INTELIGENTES
# ═══════════════════════════════════════════════════════════════════════════════

def generar_graficas_automaticas(df, cols_info, tendencia=None):
    """Genera datos de gráficas automáticamente — tipo inteligente según datos."""
    graficas = []
    
    cols_cat = [c for c in cols_info if c['tipo'] == 'categorica']
    cols_num = [c for c in cols_info if c['tipo'] == 'numerico' and 'stats' in c]
    
    # 1. Distribución por columna categórica más importante
    for cat_col in cols_cat[:3]:
        col_name = cat_col['nombre']
        if es_columna_persona(col_name) or es_columna_identificador(col_name):
            continue
        if col_name not in df.columns:
            continue
        serie = limpiar_serie_categorica(df[col_name])
        if len(serie) < 3:
            continue
        dist = serie.value_counts().head(MAX_CHART_CATEGORIES)
        labels = [str(l)[:30] for l in dist.index.tolist()]
        valores = [int(v) for v in dist.values.tolist()]
        
        if not validar_grafica(labels, valores):
            continue
        
        # Elegir tipo de gráfica inteligentemente
        n_cats = len(labels)
        total = sum(valores)
        top_pct = valores[0] / total * 100 if total > 0 else 0
        
        if n_cats <= MAX_PIE_CATEGORIES and top_pct < 70:
            tipo_grafica = 'pie'
        elif n_cats <= MAX_PIE_CATEGORIES:
            tipo_grafica = 'doughnut'
        else:
            tipo_grafica = 'bar'

        chart_limit = MAX_PIE_CATEGORIES if tipo_grafica in ('pie', 'doughnut') else MAX_BAR_CATEGORIES
        labels, valores = compactar_categorias(labels, valores, max_items=chart_limit)
        n_cats = len(labels)
        if n_cats < 2:
            continue
        
        graficas.append({
            'tipo': tipo_grafica,
            'titulo': f'Distribución por {col_name}',
            'labels': labels,
            'valores': valores,
            'dimension_label': col_name,
            'metric_label': 'Registros',
            'aggregation': 'conteo',
            'hoja_origen': getattr(df, 'attrs', {}).get('sheet_name'),
            'importancia': cat_col['importancia'] + 8,
            'insight_auto': (f"'{labels[0]}' lidera con {valores[0]:,} registros "
                           f"({top_pct:.0f}% del total).")
        })
    
    # 2. Valor numérico por categoría (gráficas de valor)
    for cat_col in cols_cat[:2]:
        if es_columna_persona(cat_col['nombre']) or es_columna_identificador(cat_col['nombre']):
            continue
        for num_col in cols_num[:3]:
            cat_name = cat_col['nombre']
            num_name = num_col['nombre']
            if es_columna_identificador(num_name):
                continue
            if cat_name not in df.columns or num_name not in df.columns:
                continue
            
            df_temp = df.copy()
            df_temp[cat_name] = limpiar_serie_categorica(df_temp[cat_name])
            df_temp = df_temp[df_temp[cat_name].notna() & (df_temp[cat_name] != '')]
            metric_profile = build_numeric_analysis_profile(df_temp, num_name)
            if not should_trust_numeric_profile(metric_profile):
                continue
            df_temp[num_name] = metric_profile['series_raw']
            df_temp = df_temp.dropna(subset=[num_name])
            
            grouped = df_temp.groupby(cat_name)[num_name].sum().sort_values(ascending=False)
            grouped = grouped[grouped > 0].head(MAX_CHART_CATEGORIES)
            
            if len(grouped) < 2:
                continue
            
            labels = [str(l)[:30] for l in grouped.index.tolist()]
            valores = [float(v) for v in grouped.values.tolist()]
            labels, valores = compactar_categorias(labels, valores, max_items=MAX_BAR_CATEGORIES)
            
            if not validar_grafica(labels, valores):
                continue
            
            total = sum(valores)
            top_pct = valores[0] / total * 100 if total > 0 else 0
            
            graficas.append({
                'tipo': 'bar',
                'titulo': f'{num_name} por {cat_name}',
                'labels': labels,
                'valores': valores,
                'dimension_label': cat_name,
                'metric_label': num_name,
                'aggregation': 'suma',
                'hoja_origen': getattr(df, 'attrs', {}).get('sheet_name'),
                'importancia': cat_col['importancia'] + num_col['importancia'],
                'insight_auto': (f"'{labels[0]}' concentra {format_number(valores[0], kind=infer_numeric_kind(num_name))} "
                               f"({top_pct:.0f}%) de {num_name}.")
            })
    
    # 3. Serie temporal / tendencia (gráfica de línea)
    if tendencia and tendencia.get('serie_temporal'):
        st = tendencia['serie_temporal']
        if len(st['labels']) >= 3 and validar_grafica(st['labels'], st['valores']):
            graficas.append({
                'tipo': 'line',
                'titulo': f"Tendencia de {tendencia['columna_valor']} en el tiempo",
                'labels': st['labels'],
                'valores': st['valores'],
                'dimension_label': tendencia.get('columna_fecha', 'Periodo'),
                'metric_label': tendencia.get('columna_valor', 'Valor'),
                'aggregation': 'tendencia',
                'hoja_origen': getattr(df, 'attrs', {}).get('sheet_name'),
                'importancia': 85,
                'insight_auto': (f"Tendencia {tendencia['tendencia']} con cambio "
                               f"del {tendencia['cambio_pct']:+.1f}% en el periodo.")
            })
    
    # 4. Top-N comparativo (si hay columna identificadora + numérica)
    for num_col in cols_num[:1]:
        if es_columna_identificador(num_col['nombre']):
            continue
        # Buscar columna de nombres/identificadores
        col_nombre = None
        for c in cols_info:
            if c['tipo'] in ('texto', 'identificador') and c['nombre'] in df.columns:
                cl = str(c['nombre']).lower()
                if 'producto' in cl:
                    col_nombre = c['nombre']
                    break
        if not col_nombre:
            continue
        
        df_temp = df[[col_nombre, num_col['nombre']]].copy()
        metric_profile = build_numeric_analysis_profile(df_temp, num_col['nombre'])
        if not should_trust_numeric_profile(metric_profile):
            continue
        df_temp[num_col['nombre']] = metric_profile['series_raw']
        df_temp = df_temp.dropna(subset=[num_col['nombre']])
        top = df_temp.groupby(col_nombre)[num_col['nombre']].sum().sort_values(ascending=False).head(MAX_CHART_CATEGORIES)
        top = top[top > 0]
        
        if len(top) < 3:
            continue
        
        labels = [str(l)[:28] for l in top.index.tolist()]
        valores = [float(v) for v in top.values.tolist()]
        labels, valores = compactar_categorias(labels, valores, max_items=MAX_BAR_CATEGORIES)
        
        if validar_grafica(labels, valores):
            graficas.append({
                'tipo': 'bar',
                'titulo': f'Top {len(labels)} por {num_col["nombre"]}',
                'labels': labels,
                'valores': valores,
                'dimension_label': col_nombre,
                'metric_label': num_col['nombre'],
                'aggregation': 'suma',
                'hoja_origen': getattr(df, 'attrs', {}).get('sheet_name'),
                'importancia': num_col['importancia'] + 15,
                'insight_auto': f"'{labels[0]}' lidera con {format_number(valores[0], kind=infer_numeric_kind(num_col['nombre']))}."
            })
    
    graficas.sort(key=lambda x: x['importancia'], reverse=True)
    return graficas[:MAX_AUTO_CHARTS]


# ═══════════════════════════════════════════════════════════════════════════════
# GENERACIÓN DE CONCLUSIONES INTELIGENTES
# ═══════════════════════════════════════════════════════════════════════════════

def generar_conclusiones(df, cols_info, kpis, es_comisiones, 
                         paretos=None, outliers_list=None, correlaciones=None, tendencia=None):
    """Genera conclusiones lógicas REALES y profundas basadas en datos analizados."""
    conclusiones = []
    total_filas = len(df)
    
    # 1. Conclusión sobre tamaño y completitud del dataset
    cols_con_datos = sum(1 for c in cols_info if c.get('stats', {}).get('conteo', 0) > 0 or c.get('valores_unicos', 0) > 0)
    conclusiones.append(f"Base de datos: {total_filas:,} registros y {cols_con_datos} columnas con información.")
    
    # 2. Conclusiones de concentración Pareto (PODEROSAS)
    for p in (paretos or []):
        if not p:
            continue
        if p['concentracion'] == 'alta':
            top = p['top_items'][0] if p['top_items'] else None
            if top:
                conclusiones.append(
                    f"CONCENTRACIÓN CRÍTICA: En '{p['columna_categoria']}', solo "
                    f"{p['categorias_80_pct']} de {p['total_categorias']} categorías "
                    f"({p['pct_categorias_para_80']:.0f}%) acumulan el 80% del valor. "
                    f"'{top['categoria']}' lidera con el {top['pct']}%.")
        elif p['concentracion'] == 'moderada':
            top = p['top_items'][0] if p['top_items'] else None
            if top:
                conclusiones.append(
                    f"Concentración moderada en '{p['columna_categoria']}': "
                    f"'{top['categoria']}' encabeza con el {top['pct']}% del total.")
    
    # 3. Conclusiones de anomalías (outliers)
    for o in (outliers_list or []):
        if not o:
            continue
        if o['total_outliers'] >= 2:
            conclusiones.append(
                f"Se identificaron {o['total_outliers']} valores atípicos en '{o['columna']}' "
                f"que superan el rango normal ({format_number(o['rango_normal'][0], kind=infer_numeric_kind(o['columna']))} a "
                f"{format_number(o['rango_normal'][1], kind=infer_numeric_kind(o['columna']))}). El valor más alto alcanza "
                f"{format_number(o['valor_max_outlier'], kind=infer_numeric_kind(o['columna']))}.")
        elif o['total_outliers'] == 1:
            conclusiones.append(
                f"Un valor atípico de {format_number(o['valor_max_outlier'], kind=infer_numeric_kind(o['columna']))} en '{o['columna']}' "
                f"supera significativamente el rango esperado.")
    
    # 4. Conclusiones de correlación
    for c in (correlaciones or [])[:2]:
        if c['correlacion'] > 0:
            conclusiones.append(
                f"Correlación {c['tipo']} positiva (r={c['correlacion']:.2f}) entre "
                f"'{c['col_a']}' y '{c['col_b']}'.")
        else:
            conclusiones.append(
                f"Correlación {c['tipo']} negativa (r={c['correlacion']:.2f}) entre "
                f"'{c['col_a']}' y '{c['col_b']}'.")
    
    # 5. Conclusiones de tendencia temporal
    if tendencia:
        if tendencia.get('tendencia') == 'creciente':
            conclusiones.append(
                f"Tendencia CRECIENTE en '{tendencia['columna_valor']}': "
                f"aumento del {tendencia['cambio_pct']:+.1f}% entre el inicio y fin del periodo. "
                f"(Promedio inicio: {format_number(tendencia['promedio_inicio'], kind=infer_numeric_kind(tendencia['columna_valor']))}, "
                f"promedio fin: {format_number(tendencia['promedio_fin'], kind=infer_numeric_kind(tendencia['columna_valor']))}.)")
        elif tendencia.get('tendencia') == 'decreciente':
            conclusiones.append(
                f"Tendencia DECRECIENTE en '{tendencia['columna_valor']}': "
                f"caída del {tendencia['cambio_pct']:+.1f}% en el periodo analizado. "
                f"(Promedio inicio: {format_number(tendencia['promedio_inicio'], kind=infer_numeric_kind(tendencia['columna_valor']))}, "
                f"promedio fin: {format_number(tendencia['promedio_fin'], kind=infer_numeric_kind(tendencia['columna_valor']))}.)")
        elif tendencia.get('tendencia') == 'estable':
            conclusiones.append(
                f"Tendencia ESTABLE en '{tendencia['columna_valor']}': "
                f"variación mínima ({tendencia['cambio_pct']:+.1f}%), manteniéndose en un promedio cercano a "
                f"{format_number(tendencia['promedio_fin'], kind=infer_numeric_kind(tendencia['columna_valor']))} al cierre del periodo.")
    
    # 6. Análisis profundo de columnas numéricas
    for col_info in cols_info:
        if col_info['tipo'] != 'numerico' or 'stats' not in col_info:
            continue
        stats = col_info['stats']
        col_name = col_info['nombre']
        
        # Variabilidad (coeficiente de variación)
        if stats['promedio'] > 0:
            cv = stats['desv_std'] / stats['promedio']
            if cv > 1.0:
                conclusiones.append(
                    f"Alta variabilidad en '{col_name}' (CV={cv:.1f}): "
                    f"Rango: {format_number(stats['minimo'], kind=infer_numeric_kind(col_name))} - "
                    f"{format_number(stats['maximo'], kind=infer_numeric_kind(col_name))}. Promedio: {format_number(stats['promedio'], kind=infer_numeric_kind(col_name))}.")
            elif cv > 0.5:
                conclusiones.append(
                    f"Variabilidad moderada en '{col_name}' (CV={cv:.1f}). "
                    f"Promedio: {format_number(stats['promedio'], kind=infer_numeric_kind(col_name))}, "
                    f"Máximo: {format_number(stats['maximo'], kind=infer_numeric_kind(col_name))}.")
        
        # Concentración en máximo
        if stats['maximo'] > 0 and stats['total'] > 0:
            ratio_max = stats['maximo'] / stats['total']
            if ratio_max > 0.25:
                conclusiones.append(
                    f"El valor máximo de '{col_name}' ({format_number(stats['maximo'], kind=infer_numeric_kind(col_name))}) "
                    f"representa el {ratio_max*100:.0f}% del total, indicando alta concentración "
                    f"en pocos registros.")
    
    # 7. Distribución categórica (top dominante)
    for col_info in cols_info:
        if col_info['tipo'] != 'categorica' or col_info['nombre'] not in df.columns:
            continue
        col_name = col_info['nombre']
        serie = limpiar_serie_categorica(df[col_name])
        if len(serie) < 5:
            continue
        dist = serie.value_counts()
        if len(dist) >= 2:
            top_val = dist.index[0]
            top_count = dist.iloc[0]
            top_pct = top_count / len(serie) * 100
            if top_pct > 50:
                segundo = dist.index[1]
                segundo_pct = dist.iloc[1] / len(serie) * 100
                conclusiones.append(
                    f"'{col_name}' concentrado en '{top_val}' ({top_pct:.0f}%, {top_count:,} registros), "
                    f"seguido de '{segundo}' ({segundo_pct:.0f}%).")
            elif len(dist) <= 5:
                top3 = ", ".join([f"'{v}' ({c:,})" for v, c in dist.head(3).items()])
                conclusiones.append(f"Distribución de '{col_name}': {top3}.")
    
    # 8. Datos faltantes relevantes
    for col_info in cols_info:
        if col_info['tipo'] == 'numerico' and 'stats' in col_info:
            pct_missing = col_info['stats']['sin_datos'] / max(1, total_filas) * 100
            if pct_missing > 20:
                conclusiones.append(
                    f"Columna '{col_info['nombre']}': {col_info['stats']['sin_datos']:,} "
                    f"registros en blanco ({pct_missing:.0f}%).")
    
    # Eliminar duplicados cercanos y limitar
    seen = set()
    unique_conclusiones = []
    for c in conclusiones:
        key = c[:60]
        if key not in seen:
            seen.add(key)
            unique_conclusiones.append(c)
    
    return unique_conclusiones[:MAX_CONCLUSIONES]


def format_number(val, kind='number', compact=True):
    """Formatea un número para presentación con convenciones consistentes."""
    val = float(val)
    abs_val = abs(val)

    if kind == 'percent':
        percent_val = val * 100 if abs_val <= 1.2 else val
        decimals = 0 if abs(percent_val) >= 10 else 1
        return f'{percent_val:.{decimals}f}%'.replace('.', ',')

    if compact and abs_val >= 1_000_000_000:
        scaled, suffix = val / 1_000_000_000, ' mil MM'
    elif compact and abs_val >= 1_000_000:
        scaled, suffix = val / 1_000_000, ' M'
    elif compact and abs_val >= 1_000:
        scaled, suffix = val / 1_000, ' mil'
    else:
        scaled, suffix = val, ''

    if compact:
        decimals = 0 if abs(scaled) >= 100 else 1
    elif abs_val >= 1000:
        decimals = 0
    elif float(val).is_integer():
        decimals = 0
    else:
        decimals = 2 if abs_val < 1 else 1

    formatted = f'{scaled:,.{decimals}f}'.replace(',', '_').replace('.', ',').replace('_', '.')
    if kind == 'currency':
        return f'COP {formatted}{suffix}'
    return f'{formatted}{suffix}'


# ═══════════════════════════════════════════════════════════════════════════════
# LECTORES ESPECIALIZADOS
# ═══════════════════════════════════════════════════════════════════════════════

def leer_coso(excel_path):
    try:
        xl = pd.ExcelFile(excel_path)
        sheet_name = None
        for sn in xl.sheet_names:
            if 'coso' in sn.lower():
                sheet_name = sn
                break
        if not sheet_name: return None
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, skiprows=3)
        header_row_idx = 0
        for i in range(min(10, len(df))):
            fila_valores = df.iloc[i].tolist()
            fila_str = " ".join([str(v).lower() for v in fila_valores])
            if any(k in fila_str for k in ['componente', 'item', 'evalua', 'punto de control', 'control']):
                header_row_idx = i
                break
        header_values = [
            str(value).strip() if pd.notna(value) and str(value).strip() else f'Col_{index}'
            for index, value in enumerate(df.iloc[header_row_idx].tolist())
        ]
        df = df.iloc[header_row_idx + 1:].copy().dropna(how='all')
        while df.shape[1] > 0 and df.iloc[:, 0].isna().all():
            df = df.iloc[:, 1:]
            header_values = header_values[1:]
        if df.shape[1] >= 3:
            df.columns = normalizar_columnas_unicas(header_values[:df.shape[1]])
            componente_col = df.columns[0]
            item_col = df.columns[1]
            estado_col = df.columns[2]
            accion_col = df.columns[3] if df.shape[1] >= 4 else None

            df[componente_col] = df[componente_col].ffill()
            df = df[df[item_col].notna()]
            df[item_col] = df[item_col].astype(str).str.strip()
            df = df[df[item_col].str.len() > 3]
            df = df[~df[item_col].str.lower().str.contains('item|evalua|punto de control', na=False)]
            df = df[~df[componente_col].astype(str).str.lower().str.contains('componente|evalua', na=False)]
            df[estado_col] = df[estado_col].apply(
                lambda value: 'Pendiente de evaluacion' if es_valor_fantasma(value) else str(value).strip()
            )

            export_cols = [componente_col, item_col, estado_col]
            export_headers = ['Componente', 'Ítems Evaluados', 'Estado']
            if accion_col:
                accion_values = df[accion_col].astype(str).str.strip()
                if accion_values.str.len().gt(0).any():
                    export_cols.append(accion_col)
                    export_headers.append('Acción Recomendada')
                    df[accion_col] = accion_values.str[:120]

            df[item_col] = df[item_col].str[:120]
            df[componente_col] = df[componente_col].astype(str).str[:80]
            df[estado_col] = df[estado_col].astype(str).str[:60]
            tabla = {
                'encabezados': export_headers,
                'filas': df[export_cols].values.tolist(),
                'hoja_origen': sheet_name,
                'sheet_family': 'coso',
            }
            if validar_tabla(tabla['encabezados'], tabla['filas']):
                return tabla
        return None
    except Exception as e: 
        print(f"Error leer_coso: {e}", file=sys.stderr)
        return None


def leer_distribucion_mes(excel_path):
    try:
        xl = pd.ExcelFile(excel_path)
        sheet_name = None
        for sn in xl.sheet_names:
            if 'td' == sn.lower().strip() or 'distribucion' in sn.lower():
                sheet_name = sn
                break
        if not sheet_name: return None
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        df = df.dropna(how='all').iloc[1:]
        if df.shape[1] > 0:
            df.columns = ['Centro de Costos'] + [f'C{i}' for i in range(df.shape[1]-1)]
            df_slide = df[['Centro de Costos']].dropna().head(12)
            tabla = {
                'encabezados': ['Centro de Costos'],
                'filas': df_slide.values.tolist(),
                'hoja_origen': sheet_name,
            }
            if validar_tabla(tabla['encabezados'], tabla['filas']):
                return tabla
        return None
    except: return None


# ═══════════════════════════════════════════════════════════════════════════════
# PRESUPUESTO DE SLIDES INTELIGENTE
# ═══════════════════════════════════════════════════════════════════════════════

def calcular_presupuesto_slides(resultado, es_comisiones):
    """
    Calcula cuántas slides asignar a cada sección para no exceder MAX_SLIDES.
    Prioriza gráficas con datos ricos y limita tablas paginadas.
    """
    presupuesto = {
        'portada': 1,
        'estructura': 0,
        'resumen_kpis': 0,
        'desglose_financiero': 0,
        'graficas': 0,
        'tabla_principal': 0,
        'hallazgos': 0,
        'coso': 0,
        'genericas': 0,
        'conclusiones': 0,
        'cierre': 1
    }
    
    slots_disponibles = MAX_SLIDES - 2  # -2 por portada y cierre
    
    # Prioridad 1: Resumen ejecutivo / KPIs (siempre)
    if resultado.get('resumen_ejecutivo') or resultado.get('resumen_generico') or resultado.get('kpis_automaticos'):
        presupuesto['resumen_kpis'] = 1
        slots_disponibles -= 1
    
    # Prioridad 2: Gráficas (máximo 4 — priorizamos datos ricos)
    num_graficas = len(resultado.get('graficas_automaticas', []))
    if es_comisiones:
        num_graficas += (1 if resultado.get('grafica_estados') else 0)
        num_graficas += (1 if resultado.get('grafica_ciudades') else 0)
        num_graficas += (1 if resultado.get('grafica_valores') else 0)
        num_graficas += (1 if resultado.get('centros_costos') else 0)
    graficas_slots = min(num_graficas, MAX_AUTO_CHARTS, slots_disponibles)
    presupuesto['graficas'] = graficas_slots
    slots_disponibles -= graficas_slots
    
    # Prioridad 3: Desglose financiero (solo si comisiones)
    if es_comisiones and resultado.get('grafica_valores') and slots_disponibles > 0:
        presupuesto['desglose_financiero'] = 1
        slots_disponibles -= 1
    
    # Prioridad 4: Top solicitantes (si comisiones)
    if es_comisiones and resultado.get('top_solicitantes') and slots_disponibles > 0:
        presupuesto['top_solicitantes'] = 1
        slots_disponibles -= 1
    
    # Prioridad 5: Tabla principal (máximo 2 páginas, ESTRICTO)
    num_filas = len(resultado.get('muestra_tabla', {}).get('filas', []))
    if num_filas > 0 and slots_disponibles > 0:
        tabla_pages = min(2, max(1, num_filas // ROWS_PER_TABLE_SLIDE + (1 if num_filas % ROWS_PER_TABLE_SLIDE else 0)))
        tabla_pages = min(tabla_pages, slots_disponibles)
        presupuesto['tabla_principal'] = tabla_pages
        slots_disponibles -= tabla_pages
    
    # Prioridad 6: Hallazgos y oportunidades (máximo 3)
    num_otras = len(resultado.get('otras_tablas', {}))
    if num_otras > 0 and slots_disponibles > 0:
        hallazgos_slots = min(num_otras, 3, slots_disponibles)
        presupuesto['hallazgos'] = hallazgos_slots
        slots_disponibles -= hallazgos_slots
    
    # Prioridad 7: COSO
    if resultado.get('coso') and slots_disponibles > 0:
        presupuesto['coso'] = 1
        slots_disponibles -= 1
    
    # Prioridad 8: Conclusiones (siempre reservamos espacio)
    if (resultado.get('conclusiones') or resultado.get('analisis_avanzado')) and slots_disponibles > 0:
        presupuesto['conclusiones'] = 1
        slots_disponibles -= 1
    
    # Prioridad 9: Estructura del archivo
    if resultado.get('metadatos', {}).get('hojas_encontradas') and slots_disponibles > 0:
        presupuesto['estructura'] = 1
        slots_disponibles -= 1
    
    # Prioridad 10: Hojas genéricas adicionales (máximo 3)
    num_genericas = len(resultado.get('genericas', {}))
    if num_genericas > 0 and slots_disponibles > 0:
        gen_slots = min(num_genericas, 3, slots_disponibles)
        presupuesto['genericas'] = gen_slots
        slots_disponibles -= gen_slots
    
    presupuesto['slots_restantes'] = slots_disponibles
    presupuesto['total_estimado'] = MAX_SLIDES - slots_disponibles
    
    return presupuesto


# ═══════════════════════════════════════════════════════════════════════════════
# PIPELINE PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

def preparar_datos_para_slides(excel_path):
    try:
        sheets = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e:
        return {"error": str(e)}

    resultado = {}
    resultado['metadatos'] = {
        'hojas_encontradas': list(sheets.keys()),
        'archivo': os.path.basename(excel_path)
    }
    
    for key in sheets.keys():
        sheets[key] = extract_real_sheet(sheets[key])

    workbook_profile = build_workbook_profile(sheets)
    resultado['metadatos']['tipo_libro'] = workbook_profile.get('tipo_libro', 'general')
    resultado['metadatos']['clasificacion_hojas'] = workbook_profile.get('familias_por_hoja', {})
    resultado['metadatos']['familias_detectadas'] = workbook_profile.get('familias_detectadas', [])
    resultado['perfil_libro'] = workbook_profile

    # === BUSCAR HOJA PRINCIPAL ===
    target_sheet = None
    if sheets:
        best_sheet = None
        best_score = -1
        for name, df in sheets.items():
            current_score = score_sheet_for_primary(name, df)
            if current_score > best_score:
                best_score = current_score
                best_sheet = name
        target_sheet = best_sheet

    processed_sheets = set()
    es_comisiones = False
    
    if target_sheet:
        resultado['metadatos']['hoja_principal'] = target_sheet
        processed_sheets.add(target_sheet)
        df = sheets[target_sheet]
        df = remover_filas_basura(df)
        df = limpiar_df(df)
        df.attrs['sheet_name'] = target_sheet
        es_comisiones = all(col in df.columns for col in ['Solicitante', 'Valor Total Solicitado'])
        
        # ══════════════════════════════════════════════════════════════
        # ANÁLISIS INTELIGENTE UNIVERSAL
        # ══════════════════════════════════════════════════════════════
        
        cols_info = detectar_columnas_importantes(df)
        resultado['_columnas_analizadas'] = len(cols_info)
        resultado['contexto_financiero'] = build_financial_context(df, cols_info)
        resultado['calidad_datos'] = build_data_quality_profile(df, cols_info)
        
        # ── ANÁLISIS AVANZADO ────────────────────────────────────────
        # Outliers
        outliers_results = []
        for c in cols_info:
            if c['tipo'] == 'numerico' and 'stats' in c and c['stats']['conteo'] >= 8:
                o = detectar_outliers(df, c['nombre'])
                if o:
                    outliers_results.append(o)
        
        # Pareto (sobre las 3 cols categóricas más importantes)
        pareto_results = []
        cols_cat = [c for c in cols_info if c['tipo'] == 'categorica']
        cols_num = [c for c in cols_info if c['tipo'] == 'numerico' and 'stats' in c]
        pareto_seen_cats = set()
        
        for cat_c in cols_cat[:3]:
            cat_name = cat_c['nombre']
            # Pareto por valor numérico principal (más informativo)
            if cols_num and cat_name not in pareto_seen_cats:
                p_val = analisis_pareto(df, cat_name, cols_num[0]['nombre'])
                if p_val:
                    pareto_results.append(p_val)
                    pareto_seen_cats.add(cat_name)
                    continue  # Skip count-based if value-based worked
            # Pareto por conteo (fallback)
            if cat_name not in pareto_seen_cats:
                p = analisis_pareto(df, cat_name)
                if p:
                    pareto_results.append(p)
                    pareto_seen_cats.add(cat_name)
        
        # Correlaciones
        corr_results = detectar_correlaciones(df, cols_info)
        
        # Tendencia temporal
        tendencia_result = detectar_tendencia_temporal(df, cols_info)
        
        # Insights avanzados
        insights_avanzados = generar_insights_avanzados(
            df, cols_info, pareto_results, outliers_results, corr_results, tendencia_result)
        
        # Guardar análisis avanzado en resultado
        resultado['analisis_avanzado'] = {
            'outliers': outliers_results[:5],
            'pareto': pareto_results[:6],
            'correlaciones': corr_results,
            'tendencia': tendencia_result,
            'insights': unique_non_empty_texts((workbook_profile.get('insights') or []) + (insights_avanzados or []), limit=MAX_INSIGHTS_AVANZADOS),
        }
        
        # ── KPIs AUTOMÁTICOS ─────────────────────────────────────────
        if es_comisiones:
            total_registros = len(df)
            valor_total = 0
            if 'Valor Total Solicitado' in df.columns:
                valor_profile = build_financial_series(df, 'Valor Total Solicitado')
                valor_total = float(valor_profile['series_valid'].sum())
            
            unique_solicitantes = int(df['Solicitante'].nunique()) if 'Solicitante' in df.columns else 0
            unique_ciudades = int(df['Ciudad Destino'].nunique()) if 'Ciudad Destino' in df.columns else 0
            
            unique_centros = 0
            if 'Centro de Costos' in df.columns:
                valid_cc = df[df['Centro de Costos'].astype(str).str.strip().str.len() > 1]
                unique_centros = int(valid_cc['Centro de Costos'].nunique())
            
            promedio_comision = valor_total / total_registros if total_registros > 0 else 0
            valor_max = float(valor_profile['series_valid'].max()) if 'Valor Total Solicitado' in df.columns and not valor_profile['series_valid'].empty else 0
            
            resultado['resumen_ejecutivo'] = {
                'total_comisiones': total_registros,
                'valor_total': valor_total,
                'unique_solicitantes': unique_solicitantes,
                'unique_ciudades': unique_ciudades,
                'unique_centros': unique_centros,
                'promedio_comision': promedio_comision,
                'valor_max_comision': valor_max,
                'conversion_resuelta': round(float(valor_profile.get('resolved_ratio') or 0), 3) if 'Valor Total Solicitado' in df.columns else 1.0,
                'monedas_detectadas': valor_profile.get('currencies_detected', ['COP']) if 'Valor Total Solicitado' in df.columns else ['COP'],
            }
        else:
            resultado['resumen_generico'] = {
                'hoja_principal': target_sheet,
                'total_filas': len(df),
                'total_columnas': int(df.shape[1]),
                'columnas_numericas': [c['nombre'] for c in cols_info if c['tipo'] == 'numerico'][:8],
                'columnas': df.columns.tolist()[:12]
            }
        
        # KPIs automáticos (para ambas rutas)
        kpis_auto = generar_kpis_automaticos(df, cols_info)
        if kpis_auto:
            resultado['kpis_automaticos'] = kpis_auto
        
        # Gráficas automáticas (para ruta genérica)
        if not es_comisiones:
            graficas_auto = []
            if workbook_profile.get('tipo_libro') != 'auditoria_control':
                graficas_auto = generar_graficas_automaticas(df, cols_info, tendencia_result)
            if graficas_auto:
                resultado['graficas_automaticas'] = graficas_auto
        
        # === TABLA PRINCIPAL ===
        cols = ['Id Comisión','Solicitante','Ciudad Destino',
                'Valor Total Solicitado','Estado','Centro de Costos']
        cols_exist = [c for c in cols if c in df.columns]
        if not cols_exist:
            top_cols = [c['nombre'] for c in cols_info[:MAX_TABLE_COLS]]
            cols_exist = [c for c in top_cols if c in df.columns]
            if not cols_exist:
                cols_exist = df.columns[:MAX_TABLE_COLS].tolist()
        
        df_slide = df[cols_exist].copy()
        # Filtrar filas completamente fantasma
        mask = df_slide.apply(lambda row: sum(1 for v in row if not es_valor_fantasma(v)) >= 2, axis=1)
        df_slide = df_slide[mask]
        
        for col in df_slide.columns:
            if pd.api.types.is_string_dtype(df_slide[col]) or df_slide[col].dtype == object:
                df_slide[col] = df_slide[col].astype(str).str[:70]
        
        tabla_data = {
            'encabezados': cols_exist,
            'filas': df_slide.values.tolist(),
            'hoja_origen': target_sheet,
        }
        if validar_tabla(tabla_data['encabezados'], tabla_data['filas']):
            resultado['muestra_tabla'] = tabla_data
        
        # === GRÁFICAS ESPECIALIZADAS (Comisiones) ===
        if es_comisiones and 'Estado' in df.columns:
            serie_estado = limpiar_serie_categorica(df['Estado'])
            estados = serie_estado.value_counts().head(MAX_CHART_CATEGORIES)
            labels = estados.index.tolist()
            valores = estados.values.tolist()
            labels, valores = compactar_categorias(labels, valores, max_items=MAX_PIE_CATEGORIES)
            if validar_grafica(labels, valores):
                resultado['grafica_estados'] = {
                    'tipo': 'doughnut',
                    'titulo': 'Distribución por Estado',
                    'labels': labels,
                    'valores': [int(v) for v in valores],
                    'colores': ['1E3A5F','4472C4','70AD47','ED7D31','FF0000','FFC000','9B59B6','3498DB'],
                    'hoja_origen': target_sheet,
                }
        
        # === GRÁFICA VALORES POR TIPO DE GASTO ===
        cols_valores = {
            'Tiquete': 'Valor Tiquete Solicitado',
            'Alimentación': 'Valor Alimentación Solicitado', 
            'Hospedaje': 'Valor Hospedaje Solicitado',
            'Transporte': 'Valor Transporte Solicitado'
        }
        vals = {}
        for nombre, col in cols_valores.items():
            if col in df.columns:
                profile = build_financial_series(df, col)
                v = float(profile['series_valid'].sum())
                if v > 0:
                    vals[nombre] = v
        
        if es_comisiones and vals:
            labels = list(vals.keys())
            valores = list(vals.values())
            if validar_grafica(labels, valores):
                resultado['grafica_valores'] = {
                    'tipo': 'bar',
                    'titulo': 'Total por Tipo de Gasto (COP)',
                    'labels': labels,
                    'valores': valores,
                    'colores': ['4472C4','ED7D31','A9D18E','FFC000'],
                    'hoja_origen': target_sheet,
                }
        
        # === TOP CIUDADES ===
        if es_comisiones and 'Ciudad Destino' in df.columns:
            serie_cd = limpiar_serie_categorica(df['Ciudad Destino'])
            ciudades = serie_cd.value_counts().head(MAX_CHART_CATEGORIES)
            labels = [str(c)[:25] for c in ciudades.index.tolist()]
            valores = [int(v) for v in ciudades.values.tolist()]
            labels, valores = compactar_categorias(labels, valores, max_items=MAX_BAR_CATEGORIES)
            if validar_grafica(labels, valores):
                resultado['grafica_ciudades'] = {
                    'tipo': 'bar',
                    'titulo': 'Top Ciudades de Destino',
                    'labels': labels,
                    'valores': valores,
                    'colores': ['4472C4'],
                    'hoja_origen': target_sheet,
                }
        
        # === TOP SOLICITANTES POR VALOR ===
        if es_comisiones and 'Solicitante' in df.columns and 'Valor Total Solicitado' in df.columns:
            df_vals = df.copy()
            df_vals['Solicitante'] = limpiar_serie_categorica(df_vals['Solicitante'])
            df_vals = df_vals[df_vals['Solicitante'].notna() & (df_vals['Solicitante'] != '')]
            valor_profile = build_financial_series(df_vals, 'Valor Total Solicitado')
            df_vals['Valor Total Solicitado'] = valor_profile['series_raw']
            df_vals = df_vals.dropna(subset=['Valor Total Solicitado'])
            top_sol = df_vals.groupby('Solicitante').agg(
                total_valor=('Valor Total Solicitado', 'sum'),
                num_comisiones=('Id Comisión', 'count') if 'Id Comisión' in df_vals.columns else ('Solicitante', 'count')
            ).sort_values('total_valor', ascending=False).head(8)
            
            labels = [str(s)[:30] for s in top_sol.index.tolist()]
            valores = [float(v) for v in top_sol['total_valor'].tolist()]
            
            if validar_grafica(labels, valores):
                resultado['top_solicitantes'] = {
                    'labels': labels,
                    'valores': valores,
                    'conteos': top_sol['num_comisiones'].tolist(),
                    'hoja_origen': target_sheet,
                }
        
        # === DISTRIBUCIÓN POR CENTRO DE COSTOS ===
        if es_comisiones and 'Centro de Costos' in df.columns and 'Valor Total Solicitado' in df.columns:
            df_cc = df.copy()
            df_cc['Centro de Costos'] = limpiar_serie_categorica(df_cc['Centro de Costos'])
            df_cc = df_cc[df_cc['Centro de Costos'].notna() & (df_cc['Centro de Costos'] != '')]
            valor_profile = build_financial_series(df_cc, 'Valor Total Solicitado')
            df_cc['Valor Total Solicitado'] = valor_profile['series_raw']
            df_cc = df_cc.dropna(subset=['Valor Total Solicitado'])
            cc_top = df_cc.groupby('Centro de Costos')['Valor Total Solicitado'].sum().sort_values(ascending=False).head(MAX_CHART_CATEGORIES)
            
            labels = cc_top.index.tolist()
            valores = [float(v) for v in cc_top.values.tolist()]
            labels, valores = compactar_categorias(labels, valores, max_items=MAX_BAR_CATEGORIES)
            
            if validar_grafica(labels, valores):
                resultado['centros_costos'] = {
                    'labels': labels,
                    'valores': valores,
                    'hoja_origen': target_sheet,
                }
        
        # === CONCLUSIONES INTELIGENTES ===
        conclusiones = generar_conclusiones(
            df, cols_info, kpis_auto, es_comisiones,
            pareto_results, outliers_results, corr_results, tendencia_result)
        conclusiones = unique_non_empty_texts((workbook_profile.get('conclusiones') or []) + (conclusiones or []), limit=MAX_CONCLUSIONES)
        if conclusiones:
            resultado['conclusiones'] = conclusiones
    
    # === Hallazgos y Oportunidades ===
    otras_tablas = {}
    for name, df in sheets.items():
        sheet_family = workbook_profile.get('familias_por_hoja', {}).get(name, 'general')
        if sheet_family in ('hallazgos', 'oportunidades'):
            df = remover_filas_basura(df)
            df = limpiar_df(df)
            if not df.empty:
                tabla_info = build_table_from_dataframe(
                    df,
                    name,
                    sheet_family=sheet_family,
                    max_cols=6,
                    max_rows=50,
                    text_limit=200,
                    min_meaningful_cells=1,
                )
                if tabla_info:
                    progress_data = None
                    for col in df.columns:
                        col_str = str(col).strip()
                        if col_str == '%' or 'porcentaje' in col_str.lower() or 'avance' in col_str.lower():
                            try:
                                progress_vals = normalize_numeric_series(df[col], col).fillna(0)
                                progress_data = progress_vals.tolist()
                            except:
                                pass
                            break
                    if progress_data:
                        tabla_info['progress'] = progress_data[:len(tabla_info.get('filas', []))]
                    otras_tablas[name] = tabla_info
    
    if otras_tablas: 
        resultado['otras_tablas'] = otras_tablas

    # === HOJAS RESTANTES ===
    genericas = {}
    for name, df in sheets.items():
        if name in processed_sheets or name == target_sheet:
            continue
        sheet_family = workbook_profile.get('familias_por_hoja', {}).get(name, 'general')
        if sheet_family in ('hallazgos', 'oportunidades'):
            continue
        if sheet_family in ('coso', 'distribucion'):
            continue
        
        if df.empty or df.shape[1] < 2 or df.shape[0] < 2:
            continue
            
        df = remover_filas_basura(df)
        df = limpiar_df(df)
        if not df.empty:
            filled_ratio = df.notna().sum().sum() / max(1, df.shape[0] * df.shape[1])
            if filled_ratio < MIN_FILL_RATIO:
                continue

            tabla_info = build_table_from_dataframe(
                df,
                name,
                sheet_family=sheet_family,
                max_cols=MAX_TABLE_COLS,
                max_rows=30,
                text_limit=150,
                min_meaningful_cells=1,
            )
            if tabla_info:
                genericas[name] = tabla_info
            
    if genericas:
        resultado['genericas'] = genericas

    # === COSO y TD ===
    coso = leer_coso(excel_path)
    if coso: resultado['coso'] = coso
    
    td = leer_distribucion_mes(excel_path)
    if td: resultado['distribucion_mes'] = td

    # === PRESUPUESTO DE SLIDES ===
    presupuesto = calcular_presupuesto_slides(resultado, es_comisiones)
    resultado['presupuesto_slides'] = presupuesto
    resultado['es_comisiones'] = es_comisiones
    if not resultado.get('conclusiones') and workbook_profile.get('conclusiones'):
        resultado['conclusiones'] = workbook_profile['conclusiones'][:MAX_CONCLUSIONES]
    if resultado.get('resumen_generico'):
        resultado['resumen_generico']['tipo_libro'] = workbook_profile.get('tipo_libro', 'general')
        resultado['resumen_generico']['familias_detectadas'] = workbook_profile.get('familias_detectadas', [])

    return resultado


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "No file path provided"}))
        sys.exit(1)
    
    import warnings
    warnings.filterwarnings('ignore')
    
    path_excel = sys.argv[1]
    if not os.path.exists(path_excel):
        print(json.dumps({"error": f"File not found: {path_excel}"}))
        sys.exit(1)
        
    try:
        data = preparar_datos_para_slides(path_excel)
        print(json.dumps(data, ensure_ascii=False, default=str))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False, default=str))
        sys.exit(1)
