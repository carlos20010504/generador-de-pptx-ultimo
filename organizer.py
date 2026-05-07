import pandas as pd
import numpy as np
import sys
import json
import os
import re
import hashlib
import time
import unicodedata
from datetime import datetime

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
MAX_TEXTUAL_BLOCKS = 6
MAX_TEXT_LINES_PER_BLOCK = 5
MAX_TEXTUAL_AI_BLOCKS = 3
TEXTUAL_AI_CACHE_FILE = os.path.join(os.path.dirname(__file__), '.cache', 'organizer_text_ai_cache.json')
EXECUTIVE_AI_CACHE_FILE = os.path.join(os.path.dirname(__file__), '.cache', 'organizer_executive_ai_cache.json')
UNIFIED_AI_CACHE_FILE = os.path.join(os.path.dirname(__file__), '.cache', 'organizer_unified_ai_cache.json')
AI_RUNTIME_STATE_FILE = os.path.join(os.path.dirname(__file__), '.cache', 'ai_runtime_state.json')
# Free-tier priority: OpenRouter (Hermes)
OPENROUTER_MODEL_PRIORITY = (
    'nousresearch/hermes-3-llama-3.1-405b:free',
)
AI_QUOTA_COOLDOWN_SECONDS = 5 * 60
AI_REQUEST_TIMEOUT_SECONDS = 90
AI_MAX_RETRIES = 3
OPENROUTER_APP_NAME = os.environ.get("OPENROUTER_APP_NAME", "Socya PPTX Generator")
OPENROUTER_SITE_URL = os.environ.get("OPENROUTER_SITE_URL", "http://localhost")
AI_WAIT_ON_RATE_LIMIT = os.environ.get("SOCYA_AI_WAIT_ON_RATE_LIMIT", "1").strip().lower() not in {"0", "false", "no"}
AI_MAX_WAIT_SECONDS = max(0, int(os.environ.get("SOCYA_AI_MAX_WAIT_SECONDS", "360") or 0))
AI_WAIT_POLL_SECONDS = max(5, int(os.environ.get("SOCYA_AI_WAIT_POLL_SECONDS", "15") or 15))
AI_EXECUTION_MODE = (os.environ.get("SOCYA_AI_EXECUTION_MODE", "best_effort") or "best_effort").strip().lower()
AI_HARD_DEADLINE_SECONDS = max(15, int(os.environ.get("SOCYA_AI_HARD_DEADLINE_SECONDS", "35") or 35))

PLACEHOLDER_VALS = {'???', '—', 'n/a', 'na', 'nan', 'none', '', '0', '-',
                    'null', 'sin datos', 'sin información', 'sin dato',
                    'no aplica', 'no disponible', 'nd', 's/d'}
STRICT_VISUAL_CURATION_MODE = True
LOCAL_PROMPT_STOPWORDS = {
    'para', 'sobre', 'desde', 'entre', 'hasta', 'hacia', 'como', 'este', 'esta', 'estos', 'estas',
    'solo', 'toda', 'todo', 'todos', 'todas', 'datos', 'excel', 'archivo', 'powerpoint', 'ppt',
    'presentacion', 'presentaciones', 'diapositiva', 'diapositivas', 'grafico', 'grafica', 'graficas',
    'tabla', 'tablas', 'prompt', 'usuario', 'analisis', 'enfoque', 'enfasis', 'prioriza', 'priorizar',
    'genera', 'generar', 'muestra', 'mostrar', 'real', 'reales', 'ejecutivo', 'ejecutiva', 'ejecutivas',
    'detallado', 'detallada', 'profesional', 'coherente', 'visual', 'atractivo', 'atractiva', 'preciso',
    'precisa', 'significativo', 'significativa', 'inteligente', 'estructurado', 'estructurada',
    'enfocada', 'enfocado', 'centrada', 'centrado', 'basado', 'basada', 'basados', 'basadas',
    'prioritario', 'prioritaria', 'prioritarios', 'prioritarias', 'decision', 'decisiones',
}

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

TEXTUAL_FAMILIES = {
    'auditoria',
    'checklist',
    'hallazgos',
    'oportunidades',
    'matriz_riesgos',
    'evidencias',
    'procedimiento',
    'cuestionario',
    'coso',
}

TEXT_ARTIFACT_REPLACEMENTS = {
    '├í': 'á',
    '├⌐': 'é',
    '├®': 'í',
    '├│': 'ó',
    '├║': 'ú',
    '├▒': 'ñ',
    '├ü': 'Ü',
    '├ä': 'Ä',
    '├ô': 'Ó',
    '├Ü': 'Ñ',
    'Ã¡': 'á',
    'Ã©': 'é',
    'Ã­': 'í',
    'Ã³': 'ó',
    'Ãº': 'ú',
    'Ã±': 'ñ',
    'Ã': 'A',
    'â€™': "'",
    'â€œ': '"',
    'â€\x9d': '"',
    'â€“': '-',
    'â€”': '-',
    'â€¢': '-',
    '’': "'",
    '‘': "'",
    '“': '"',
    '”': '"',
    '¾': 'ó',
    'ß': 'á',
    'Ý': 'í',
    'Ú': 'É',
    'Ë': 'Ó',
    '═': 'Í',
    '┐': '¿',
    '┬┐': '¿',
    'À': '-',
}

DEFAULT_PRESENTATION_THEME = {
    'name': 'Analitica Moderna',
    'primary_hex': '#0F172A',
    'accent_hex': '#2563EB',
    'text_hex': '#E5E7EB',
    'bg_hex': '#F8FAFC',
}

THEME_PRESETS = {
    'analitica-moderna': DEFAULT_PRESENTATION_THEME,
    'comite-ejecutivo': {
        'name': 'Comite Ejecutivo',
        'primary_hex': '#0B1F3A',
        'accent_hex': '#C0841A',
        'text_hex': '#E2E8F0',
        'bg_hex': '#F8FAFC',
    },
    'impacto-nocturno': {
        'name': 'Impacto Nocturno',
        'primary_hex': '#111827',
        'accent_hex': '#7C3AED',
        'text_hex': '#F3F4F6',
        'bg_hex': '#030712',
    },
    'socya-verde': {
        'name': 'Socya Verde',
        'primary_hex': '#14532D',
        'accent_hex': '#16A34A',
        'text_hex': '#E8F5EC',
        'bg_hex': '#F7FCF8',
    },
}

MOJIBAKE_HINT_CHARS = 'ÃÂ├┤┐╔╗═¾ÝËÐ�'

# ═══════════════════════════════════════════════════════════════════════════════
# UTILIDADES DE LIMPIEZA
# ═══════════════════════════════════════════════════════════════════════════════

def normalizar_columnas_unicas(columns):
    usadas = {}
    resultado = []
    for idx, col in enumerate(columns):
        if pd.notna(col) and str(col).strip() and not str(col).startswith('Unnamed'):
            nombre = repair_text_artifacts(str(col).strip())
        else:
            nombre = f"Col_{idx}"
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
        
        # Fallback: Si no encontró nada claro en las primeras 20, pero hay una fila 1 que parece cabecera
        if len(raw_data) > 1:
             valid_row1 = [x for x in raw_data[1] if pd.notna(x) and str(x).strip() and not str(x).startswith('Unnamed')]
             if len(valid_row1) >= 2:
                 df_new = pd.DataFrame(raw_data[2:], columns=raw_data[1])
                 df_new.columns = normalizar_columnas_unicas(df_new.columns)
                 return df_new
    except: pass
    return df


def limpiar_df(df):
    df = df.copy()
    df = df.dropna(how='all', axis=0).dropna(how='all', axis=1)
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
            df.iloc[:, idx] = serie.fillna('')
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
    
    # Identificación flexible de columnas importantes
    cols_norm = {normalize_semantic_text(c): c for c in df.columns}
    c_id = cols_norm.get('id comision') or cols_norm.get('id') or cols_norm.get('codigo') or cols_norm.get('consecutivo')
    c_sol = cols_norm.get('solicitante')
    c_val = cols_norm.get('valor total solicitado') or cols_norm.get('valor total')
    
    cols_check = [c for c in [c_id, c_sol, c_val] if c]
    if cols_check:
        df = df[df[cols_check].notna().any(axis=1)]
        if c_id:
            df = df[df[c_id].notna()]
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
    if kind_hint and is_temporal_header(kind_hint):
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
    text = repair_text_artifacts(value).replace('\xa0', ' ').strip().lower()
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


def contains_internal_slide_metadata(value):
    text = str(value or '').strip()
    if not text:
        return False
    normalized = normalize_semantic_text(text)
    if not normalized:
        return False
    if re.search(r'\b(title|subtitle|type)\s*:', text, flags=re.IGNORECASE):
        return True
    return normalized in {
        'text',
        'subtitle organizado automaticamente text',
        'organizado automaticamente text',
        'type text',
    }


def sanitize_executive_text(value, max_len=220):
    text = repair_text_artifacts(value).replace('_x000d_', ' ').replace('\r', ' ').replace('\n', ' ')
    text = re.sub(r'\s+', ' ', text).strip(" .:-")
    if not text:
        return ''
    text = text.replace('RECOMEDACIONES', 'RECOMENDACIONES')
    text = text.replace('recomedaciones', 'recomendaciones')
    text = re.sub(
        r'^(Hallazgo|Mejora|Riesgo|Procedimiento|Control|Prueba|Evidencia|Pregunta|Componente)\s*:\s*',
        '',
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r'^(TITLE|SUBTITLE|TYPE)\s*:\s*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\s+[—-]\s+TEXT$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'^(TEXT|LAYOUT)\s*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\s+', ' ', text).strip(" .:-")
    if not text:
        return ''
    if contains_internal_slide_metadata(text):
        return ''
    normalized = normalize_semantic_text(text)
    if normalized in {'title', 'subtitle', 'type', 'text', 'layout'}:
        return ''
    if len(normalized) < 4:
        return ''
    if text and text[0].islower():
        text = text[0].upper() + text[1:]
    return text[:max_len]


def mojibake_score(text):
    return sum(text.count(char) for char in MOJIBAKE_HINT_CHARS)


def attempt_redecode(text, source_encoding, target_encoding):
    try:
        return text.encode(source_encoding, errors='ignore').decode(target_encoding, errors='ignore')
    except Exception:
        return text


def repair_mojibake_text(text):
    current = str(text or '')
    if not current:
        return current
    best = current
    best_score = mojibake_score(current)
    candidates = [
        attempt_redecode(current, 'latin1', 'utf-8'),
        attempt_redecode(current, 'cp1252', 'utf-8'),
        attempt_redecode(current, 'latin1', 'cp1252'),
    ]
    for candidate in candidates:
        if not candidate:
            continue
        score = mojibake_score(candidate)
        if score < best_score and sum(char.isalnum() for char in candidate) >= max(2, sum(char.isalnum() for char in current) // 2):
            best = candidate
            best_score = score
    return best


def repair_text_artifacts(value):
    cleaned = repair_mojibake_text(str(value or ''))
    for source, target in TEXT_ARTIFACT_REPLACEMENTS.items():
        cleaned = cleaned.replace(source, target)
    cleaned = unicodedata.normalize('NFKC', cleaned)
    cleaned = re.sub(r'[\u200b-\u200f\u202a-\u202e]', '', cleaned)
    cleaned = cleaned.replace('\xa0', ' ').replace('_x000d_', ' ')
    cleaned = re.sub(r'\s+', ' ', cleaned).strip()
    cleaned = unicodedata.normalize('NFKD', cleaned).encode('ascii', 'ignore').decode('ascii')
    return re.sub(r'\s+', ' ', cleaned).strip()


def build_default_ai_curation_status():
    has_api_key = bool(get_openrouter_api_key())
    temporarily_blocked = is_ai_temporarily_blocked() if has_api_key else False
    reason = 'ready_for_ai' if has_api_key and not temporarily_blocked else (
        'ai_temporarily_blocked' if temporarily_blocked else 'ai_api_unavailable'
    )
    return {
        'strict_visual_mode': STRICT_VISUAL_CURATION_MODE,
        'api_key_available': has_api_key,
        'temporarily_blocked': temporarily_blocked,
        'unified_call_succeeded': False,
        'briefing_received': False,
        'visual_plan_received': False,
        'visual_curation_ready': False,
        'selected_chart_ids': [],
        'selected_table_ids': [],
        'reason': reason,
        'execution_mode': AI_EXECUTION_MODE,
        'provider_used': False,
        'local_fallback_used': False,
    }


def normalize_theme_context(theme_payload):
    if not isinstance(theme_payload, dict):
        return dict(DEFAULT_PRESENTATION_THEME)

    theme_key = sanitize_executive_text(theme_payload.get('key'), max_len=40).lower()
    preset = dict(THEME_PRESETS.get(theme_key) or DEFAULT_PRESENTATION_THEME)
    normalized = {
        'name': sanitize_executive_text(theme_payload.get('name') or preset.get('name'), max_len=60) or preset['name'],
        'primary_hex': sanitize_executive_text(theme_payload.get('primary_hex') or preset.get('primary_hex'), max_len=9) or preset['primary_hex'],
        'accent_hex': sanitize_executive_text(theme_payload.get('accent_hex') or preset.get('accent_hex'), max_len=9) or preset['accent_hex'],
        'text_hex': sanitize_executive_text(theme_payload.get('text_hex') or preset.get('text_hex'), max_len=9) or preset['text_hex'],
        'bg_hex': sanitize_executive_text(theme_payload.get('bg_hex') or preset.get('bg_hex'), max_len=9) or preset['bg_hex'],
    }
    return normalized


def parse_user_request_context(user_instructions):
    if isinstance(user_instructions, dict):
        payload = dict(user_instructions)
    else:
        raw = str(user_instructions or '').strip()
        payload = {}
        if raw.startswith('{') and raw.endswith('}'):
            try:
                parsed = json.loads(raw)
                if isinstance(parsed, dict):
                    payload = parsed
            except Exception:
                payload = {}
        if not payload:
            payload = {'prompt': raw}

    prompt = sanitize_executive_text(payload.get('prompt') or payload.get('userPrompt') or user_instructions, max_len=1200)
    audience = sanitize_executive_text(payload.get('audience'), max_len=40).lower() or 'ejecutivos'
    language = sanitize_executive_text(payload.get('language'), max_len=24) or 'Español'
    current_date = sanitize_executive_text(payload.get('current_date'), max_len=24) or datetime.now().strftime('%d/%m/%Y')
    theme = normalize_theme_context(payload.get('theme'))

    return {
        'prompt': prompt,
        'audience': audience,
        'language': language,
        'current_date': current_date,
        'theme': theme,
        'raw_payload': payload,
    }


def build_pandas_summary_for_ai(analisis_data):
    generico = analisis_data.get('resumen_generico') or {}
    metadatos = analisis_data.get('metadatos') or {}
    avanzado = analisis_data.get('analisis_avanzado') or {}
    muestra = analisis_data.get('muestra_tabla') or {}
    conclusions = unique_non_empty_texts(
        [sanitize_executive_text(item, max_len=160) for item in (analisis_data.get('conclusiones') or [])],
        limit=6,
    )
    charts = analisis_data.get('graficas_automaticas') or []
    kpis = analisis_data.get('kpis_automaticos') or []

    summary_lines = [
        f"Archivo: {sanitize_executive_text(metadatos.get('archivo'), max_len=80)}",
        f"Hoja principal: {sanitize_executive_text(generico.get('hoja_principal') or metadatos.get('hoja_principal'), max_len=80)}",
        f"Filas totales: {generico.get('total_filas') or 0}",
        f"Columnas totales: {generico.get('total_columnas') or 0}",
        "Columnas detectadas: " + ", ".join(
            sanitize_executive_text(item, max_len=30) for item in (generico.get('columnas') or [])[:10]
        ),
        "Columnas numericas: " + ", ".join(
            sanitize_executive_text(item, max_len=30) for item in (generico.get('columnas_numericas') or [])[:8]
        ),
        "KPIs detectados: " + ", ".join(
            sanitize_executive_text((item or {}).get('label') or (item or {}).get('valor'), max_len=60)
            for item in kpis[:5] if isinstance(item, dict)
        ),
        "Graficas candidatas: " + ", ".join(
            sanitize_executive_text((item or {}).get('titulo'), max_len=60)
            for item in charts[:5] if isinstance(item, dict)
        ),
        "Conclusiones base: " + " | ".join(conclusions),
    ]

    pareto = avanzado.get('pareto') or []
    for item in pareto[:3]:
        if isinstance(item, dict):
            top_item = (item.get('top_valores') or [{}])[0] if item.get('top_valores') else {}
            summary_lines.append(
                f"Pareto {sanitize_executive_text(item.get('columna'), max_len=40)}: "
                f"{sanitize_executive_text(top_item.get('valor'), max_len=40)} "
                f"({round(top_item.get('pct', 0), 1) if isinstance(top_item.get('pct'), (int, float)) else 'n/a'}%)"
            )

    trend = avanzado.get('tendencia') or {}
    if isinstance(trend, dict) and trend:
        summary_lines.append(
            "Tendencia temporal: "
            f"{sanitize_executive_text(trend.get('col_temporal'), max_len=40)} vs "
            f"{sanitize_executive_text(trend.get('col_valor'), max_len=40)} = "
            f"{sanitize_executive_text(trend.get('tendencia'), max_len=30)}"
        )

    headers = [sanitize_executive_text(item, max_len=30) for item in (muestra.get('encabezados') or [])[:6]]
    rows = []
    for row in (muestra.get('filas') or [])[:5]:
        if isinstance(row, (list, tuple)):
            rows.append([sanitize_executive_text(cell, max_len=30) for cell in row[:6]])
    if headers:
        summary_lines.append("Tabla principal encabezados: " + ", ".join(headers))
    if rows:
        summary_lines.append("Muestra filas: " + json.dumps(rows, ensure_ascii=False))

    return "\n".join(line for line in summary_lines if line and not line.endswith(": "))


def normalize_presentation_design_ai_payload(ai_payload):
    if not isinstance(ai_payload, dict):
        return None

    meta = ai_payload.get('presentation_meta') if isinstance(ai_payload.get('presentation_meta'), dict) else {}
    slides_raw = ai_payload.get('slides') if isinstance(ai_payload.get('slides'), list) else []
    global_design = ai_payload.get('global_design') if isinstance(ai_payload.get('global_design'), dict) else {}
    validation = ai_payload.get('validation') if isinstance(ai_payload.get('validation'), dict) else {}

    normalized_slides = []
    valid_types = {
        'title_slide', 'executive_summary', 'text_bullets', 'chart_full', 'chart_with_insight',
        'data_table', 'two_columns', 'section_divider', 'quote_highlight', 'closing_slide',
    }
    for index, slide in enumerate(slides_raw[:12], start=1):
        if not isinstance(slide, dict):
            continue
        slide_type = sanitize_executive_text(slide.get('type'), max_len=40)
        if slide_type not in valid_types:
            continue
        normalized_slide = {
            'slide_number': index,
            'type': slide_type,
            'title': sanitize_executive_text(slide.get('title'), max_len=120),
            'design_notes': sanitize_executive_text(slide.get('design_notes'), max_len=220),
        }
        if slide_type == 'title_slide':
            layout = slide.get('layout') if isinstance(slide.get('layout'), dict) else {}
            normalized_slide['title_text'] = sanitize_executive_text(((layout.get('title') or {}).get('text') if isinstance(layout.get('title'), dict) else slide.get('title')), max_len=120)
            normalized_slide['subtitle_text'] = sanitize_executive_text(((layout.get('subtitle') or {}).get('text') if isinstance(layout.get('subtitle'), dict) else meta.get('subtitle')), max_len=140)
        if slide_type in {'chart_full', 'chart_with_insight'}:
            chart = slide.get('chart') if isinstance(slide.get('chart'), dict) else {}
            insight_box = slide.get('insight_box') if isinstance(slide.get('insight_box'), dict) else {}
            normalized_slide['chart_type'] = sanitize_executive_text(chart.get('chart_type'), max_len=30)
            normalized_slide['insight_headline'] = sanitize_executive_text(insight_box.get('headline'), max_len=60)
            normalized_slide['insight_body'] = sanitize_executive_text(insight_box.get('body'), max_len=220)
        if slide_type == 'data_table':
            normalized_slide['footnote'] = sanitize_executive_text(slide.get('footnote'), max_len=140)
        if slide_type == 'executive_summary':
            normalized_slide['insight_text'] = sanitize_executive_text(slide.get('insight_text'), max_len=220)
        if slide_type == 'closing_slide':
            normalized_slide['cta_text'] = sanitize_executive_text(slide.get('cta_text'), max_len=160)
            normalized_slide['summary_bullets'] = unique_non_empty_texts(
                [sanitize_executive_text(item, max_len=140) for item in (slide.get('summary_bullets') or [])],
                limit=5,
            )
        normalized_slides.append(normalized_slide)

    if not normalized_slides:
        return None

    return {
        'presentation_meta': {
            'title': sanitize_executive_text(meta.get('title'), max_len=120),
            'subtitle': sanitize_executive_text(meta.get('subtitle'), max_len=160),
            'author': sanitize_executive_text(meta.get('author'), max_len=80),
            'date': sanitize_executive_text(meta.get('date'), max_len=24),
            'total_slides': min(12, max(1, int(meta.get('total_slides') or len(normalized_slides)))),
            'narrative_summary': sanitize_executive_text(meta.get('narrative_summary'), max_len=240),
        },
        'slides': normalized_slides,
        'global_design': {
            'font_primary': sanitize_executive_text(global_design.get('font_primary'), max_len=40),
            'font_secondary': sanitize_executive_text(global_design.get('font_secondary'), max_len=40),
            'footer_text': sanitize_executive_text(global_design.get('footer_text'), max_len=100),
        },
        'validation': {
            key: bool(validation.get(key))
            for key in (
                'total_slides_matches_array',
                'all_slides_have_required_fields',
                'no_slide_exceeds_5_bullets',
                'no_chart_exceeds_8_series',
                'json_is_valid',
            )
        },
    }


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

    def valor_real_tabla(val):
        if pd.isna(val):
            return False
        if isinstance(val, (int, float, np.integer, np.floating)) and not pd.isna(val):
            return True
        return not es_valor_fantasma(val)

    def normalizar_valor_tabla(val):
        if pd.isna(val):
            return ''
        if isinstance(val, (pd.Timestamp, datetime)):
            return val.strftime('%d/%m/%Y')
        if isinstance(val, (int, float, np.integer, np.floating)) and not pd.isna(val):
            if float(val).is_integer():
                return str(int(val))
            return str(float(val))
        return str(val).strip()[:text_limit]

    scored_cols = []
    for position, col in enumerate(selected_cols):
        serie = df[col]
        real_ratio = float(serie.apply(valor_real_tabla).mean()) if len(serie) else 0.0
        score = real_ratio * 100
        if es_dimension_ejecutiva(col):
            score += 14
        if infer_numeric_kind(col) in ('currency', 'percent'):
            score += 10
        if es_columna_identificador(col):
            score -= 18
        if es_columna_persona(col):
            score -= 6
        score -= position * 0.25
        scored_cols.append((score, col, real_ratio))

    scored_cols.sort(key=lambda item: item[0], reverse=True)
    dense_cols = [col for _score, col, real_ratio in scored_cols if real_ratio >= 0.55]
    ordered_cols = dense_cols + [col for _score, col, _ratio in scored_cols if col not in dense_cols]

    best_subset = None
    max_width = min(max_cols, len(ordered_cols))
    for width in range(max_width, 1, -1):
        subset_cols = ordered_cols[:width]
        df_candidate = df[subset_cols].copy()
        for col in subset_cols:
            if pd.api.types.is_string_dtype(df_candidate[col]) or df_candidate[col].dtype == object:
                df_candidate[col] = df_candidate[col].apply(normalizar_valor_tabla)
        row_mask = df_candidate.apply(lambda row: all(valor_real_tabla(value) for value in row), axis=1)
        dense_df = df_candidate[row_mask].head(max_rows).copy()
        if len(dense_df) >= max(2, min_meaningful_cells):
            best_subset = (subset_cols, dense_df)
            break

    if best_subset is None:
        return None

    selected_cols, df_res = best_subset

    tabla_info = {
        'encabezados': [str(col) for col in selected_cols],
        'filas': df_res.values.tolist(),
        'hoja_origen': sheet_name,
        'sheet_family': sheet_family,
    }
    if validar_tabla(tabla_info['encabezados'], tabla_info['filas']):
        return tabla_info
    return None


def build_textual_block_label(sheet_family):
    mapping = {
        'auditoria': 'Control y verificacion',
        'checklist': 'Cobertura de pruebas',
        'hallazgos': 'Hallazgos relevantes',
        'oportunidades': 'Oportunidades de mejora',
        'matriz_riesgos': 'Riesgos y mitigacion',
        'evidencias': 'Soportes y evidencias',
        'procedimiento': 'Procedimiento evaluado',
        'cuestionario': 'Cuestionario de control',
        'coso': 'Lectura COSO',
    }
    return mapping.get(sheet_family, 'Lectura narrativa')


def looks_like_meaningful_text(value):
    text = str(value or '').strip()
    if not text or es_valor_fantasma(text):
        return False
    if parse_numeric_value(text) is not None and '%' not in text and '$' not in text and len(text) < 8:
        return False
    words = [part for part in text.split() if part]
    return len(words) >= 4 or len(text) >= 28


def pick_textual_columns(df, sheet_family='general', max_cols=4):
    preferred_keywords = {
        'auditoria': ['pregunta', 'criterio', 'revision', 'verificacion', 'observacion', 'control'],
        'checklist': ['prueba', 'cumple', 'observacion', 'control', 'resultado'],
        'hallazgos': ['hallazgo', 'descripcion', 'detalle', 'riesgo', 'impacto', 'observacion', 'recomendacion'],
        'oportunidades': ['oportunidad', 'descripcion', 'detalle', 'accion', 'mejora', 'observacion', 'recomendacion'],
        'matriz_riesgos': ['riesgo', 'causa', 'consecuencia', 'control', 'recomendacion', 'accion'],
        'evidencias': ['evidencia', 'soporte', 'documento', 'descripcion', 'observacion'],
        'procedimiento': ['pregunta', 'respuesta', 'procedimiento', 'control', 'observacion', 'descripcion'],
        'cuestionario': ['pregunta', 'respuesta', 'observacion', 'estado'],
        'coso': ['componente', 'item', 'estado', 'accion', 'observacion'],
    }
    keywords = [normalize_semantic_text(item) for item in preferred_keywords.get(sheet_family, [])]
    scored = []
    for col in df.columns:
        col_name = str(col).strip()
        normalized_col = normalize_semantic_text(col_name)
        series = df[col].head(60)
        non_empty = [
            str(value).strip()
            for value in series
            if pd.notna(value) and not es_valor_fantasma(value)
        ]
        if not non_empty:
            continue
        text_like = sum(1 for value in non_empty if looks_like_meaningful_text(value))
        score = text_like * 3
        if keywords and any(keyword in normalized_col for keyword in keywords):
            score += 12
        if not es_columna_generica(col_name):
            score += 2
        if score > 0:
            scored.append((score, col))
    scored.sort(key=lambda item: item[0], reverse=True)
    selected = [col for _score, col in scored[:max_cols]]
    if selected:
        return selected
    return select_semantic_columns(df, sheet_family=sheet_family, max_cols=max_cols)


def format_textual_row(row_data, sheet_family='general'):
    ordered_items = [(str(col).strip(), str(value).strip()) for col, value in row_data if str(value).strip()]
    if not ordered_items:
        return None

    fragments = []
    skip_labels = {
        'title',
        'subtitle',
        'type',
        'text',
        'layout',
        'col 0',
        'col 1',
    }
    for label, value in ordered_items:
        clean_value = sanitize_executive_text(value, max_len=220)
        normalized_label = normalize_semantic_text(label)
        if (
            not clean_value
            or es_valor_fantasma(clean_value)
            or normalized_label in skip_labels
            or contains_internal_slide_metadata(label)
        ):
            continue
        if not fragments:
            fragments.append(clean_value)
            continue
        short_label = label[:40]
        if normalized_label in {'descripcion', 'detalle', 'respuesta', 'observacion', 'hallazgo', 'recomendacion', 'accion recomendada', 'causa', 'consecuencia'}:
            fragments.append(clean_value)
        else:
            fragments.append(f"{short_label}: {clean_value}")
        if len(fragments) >= 3:
            break

    if not fragments:
        return None
    line = ". ".join(fragments)
    line = re.sub(r'\s+', ' ', line).strip(" .")
    line = sanitize_executive_text(line, max_len=280)
    return line if line else None


def extract_textual_points_from_dataframe(df, sheet_name, sheet_family='general', max_points=MAX_TEXT_LINES_PER_BLOCK):
    if df is None or df.empty or sheet_family not in TEXTUAL_FAMILIES:
        return []

    selected_cols = pick_textual_columns(df, sheet_family=sheet_family, max_cols=4)
    if not selected_cols:
        return []

    points = []
    for _idx, row in df.head(40).iterrows():
        row_data = []
        for col in selected_cols:
            value = row.get(col, '')
            text = str(value).strip()
            if not text or es_valor_fantasma(text):
                continue
            if parse_numeric_value(text) is not None and '%' not in text and '$' not in text and len(text) < 8:
                continue
            row_data.append((col, text))
        line = format_textual_row(row_data, sheet_family=sheet_family)
        if line:
            points.append(line)

    return unique_non_empty_texts(points, limit=max_points)


def build_textual_block_from_dataframe(df, sheet_name, sheet_family='general'):
    lines = extract_textual_points_from_dataframe(df, sheet_name, sheet_family=sheet_family, max_points=MAX_TEXT_LINES_PER_BLOCK)
    if not lines:
        return None
    return {
        'title': sheet_name,
        'subtitle': build_textual_block_label(sheet_family),
        'hoja_origen': sheet_name,
        'sheet_family': sheet_family,
        'lines': lines,
        'source_mode': 'python',
    }


def get_openrouter_api_key():
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if api_key:
        return api_key.strip().replace('"', '').replace("'", "")

    try:
        for env_file in ['.env', '.env.local', '../.env']:
            if os.path.exists(env_file):
                with open(env_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                    if "OPENROUTER_API_KEY" in content:
                        for line in content.splitlines():
                            if line.startswith("OPENROUTER_API_KEY="):
                                value = line.split('=', 1)[1].strip().replace('"', '').replace("'", "")
                                if value: return value
    except Exception:
        pass
    return None


def load_textual_ai_cache():
    try:
        if os.path.exists(TEXTUAL_AI_CACHE_FILE):
            with open(TEXTUAL_AI_CACHE_FILE, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {}


def load_executive_ai_cache():
    try:
        if os.path.exists(EXECUTIVE_AI_CACHE_FILE):
            with open(EXECUTIVE_AI_CACHE_FILE, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {}


def save_textual_ai_cache(cache_data):
    try:
        os.makedirs(os.path.dirname(TEXTUAL_AI_CACHE_FILE), exist_ok=True)
        with open(TEXTUAL_AI_CACHE_FILE, 'w', encoding='utf-8') as handle:
            json.dump(cache_data, handle, ensure_ascii=False, indent=2)
    except Exception:
        pass


def save_executive_ai_cache(cache_data):
    try:
        os.makedirs(os.path.dirname(EXECUTIVE_AI_CACHE_FILE), exist_ok=True)
        with open(EXECUTIVE_AI_CACHE_FILE, 'w', encoding='utf-8') as handle:
            json.dump(cache_data, handle, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_ai_runtime_state():
    try:
        if os.path.exists(AI_RUNTIME_STATE_FILE):
            with open(AI_RUNTIME_STATE_FILE, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {}


def save_ai_runtime_state(state):
    try:
        os.makedirs(os.path.dirname(AI_RUNTIME_STATE_FILE), exist_ok=True)
        with open(AI_RUNTIME_STATE_FILE, 'w', encoding='utf-8') as handle:
            json.dump(state, handle, ensure_ascii=False, indent=2)
    except Exception:
        pass


def parse_retry_delay_seconds(message):
    text = str(message or '')
    patterns = [
        r'retry in\s+([0-9]+(?:\.[0-9]+)?)s',
        r'seconds:\s*([0-9]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            try:
                return max(0, int(float(match.group(1))))
            except Exception:
                continue
    return 0


def is_openrouter_rate_limit_message(message):
    text = str(message or '')
    if not text:
        return False
    normalized = text.lower()
    if 'provider returned error' in normalized and not any(
        token in normalized for token in ['rate limit', 'high demand', 'limit_rpm', '429', 'limited to']
    ):
        return False
    return (
        'rate limit exceeded' in normalized
        or 'limited to' in normalized
        or 'high demand for' in normalized
        or 'limit_rpm' in normalized
        or '429' in normalized
        or 'temporarily rate-limited upstream' in normalized
        or 'provider returned error' in normalized
    )


def get_ai_cooldown_remaining():
    state = load_ai_runtime_state()
    blocked_until = float(state.get('blocked_until') or 0)
    remaining = int(blocked_until - time.time())
    return max(0, remaining)


def is_ai_temporarily_blocked():
    return get_ai_cooldown_remaining() > 0


def mark_ai_quota_blocked(message=None):
    retry_seconds = parse_retry_delay_seconds(message)
    cooldown = max(AI_QUOTA_COOLDOWN_SECONDS, retry_seconds)
    state = {
        'blocked_until': time.time() + cooldown,
        'reason': 'quota',
        'retry_after_seconds': retry_seconds,
        'updated_at': int(time.time()),
    }
    save_ai_runtime_state(state)


def clear_ai_quota_block():
    state = load_ai_runtime_state()
    if state:
        save_ai_runtime_state({})


def is_transient_openrouter_message(message):
    text = str(message or '')
    if not text:
        return False
    normalized = text.lower()
    transient_tokens = (
        'provider returned error',
        'upstream error',
        'temporary',
        'timed out',
        'timeout',
        'connection reset',
        'service unavailable',
        'overloaded',
        'internal server error',
        'bad gateway',
        'gateway timeout',
    )
    return any(token in normalized for token in transient_tokens)


def extract_openrouter_error_message(payload):
    if isinstance(payload, dict):
        error = payload.get('error')
        if isinstance(error, dict):
            pieces = [
                str(error.get('message') or '').strip(),
                str(error.get('code') or '').strip(),
                str(((error.get('metadata') or {}).get('raw')) or '').strip() if isinstance(error.get('metadata'), dict) else '',
            ]
            return ' | '.join(part for part in pieces if part)
        return str(error or '').strip()
    return str(payload or '').strip()


def normalize_ai_json_response_text(text):
    raw = str(text or '').strip()
    if not raw:
        return None
    json_match = re.search(r'(\{.*\}|\[.*\])', raw.replace('\n', ' '), re.DOTALL)
    if json_match:
        return json_match.group(1)
    return raw.replace("```json", "").replace("```", "").strip()


def maybe_wait_for_ai_cooldown(wait_budget_seconds):
    if not AI_WAIT_ON_RATE_LIMIT:
        return 0, wait_budget_seconds
    remaining_budget = max(0, int(wait_budget_seconds or 0))
    waited = 0
    while remaining_budget > 0:
        cooldown_remaining = get_ai_cooldown_remaining()
        if cooldown_remaining <= 0:
            clear_ai_quota_block()
            break
        sleep_seconds = min(cooldown_remaining, AI_WAIT_POLL_SECONDS, remaining_budget)
        if sleep_seconds <= 0:
            break
        print(
            f"INFO: Esperando disponibilidad de Hermes/OpenRouter ({cooldown_remaining}s restantes, pausa {sleep_seconds}s)",
            file=sys.stderr,
        )
        time.sleep(sleep_seconds)
        waited += sleep_seconds
        remaining_budget -= sleep_seconds
    return waited, remaining_budget


def build_textual_block_cache_key(block):
    payload = {
        'title': block.get('title'),
        'subtitle': block.get('subtitle'),
        'sheet_family': block.get('sheet_family'),
        'lines': block.get('lines') or [],
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()


def build_executive_briefing_cache_key(payload):
    versioned_payload = {
        'prompt_version': 'executive-briefing-v2',
        'payload': payload,
    }
    raw = json.dumps(versioned_payload, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()


def merge_ai_textual_block(block, ai_payload):
    merged = dict(block)
    if isinstance(ai_payload, dict):
        title = sanitize_executive_text(ai_payload.get('title'), max_len=90)
        subtitle = sanitize_executive_text(ai_payload.get('subtitle'), max_len=90)
        lines = unique_non_empty_texts(
            [sanitize_executive_text(line, max_len=180) for line in (ai_payload.get('lines') or [])],
            limit=MAX_TEXT_LINES_PER_BLOCK,
        )
        if title:
            merged['title'] = title[:90]
        if subtitle:
            merged['subtitle'] = subtitle[:90]
        if lines:
            merged['lines'] = lines
            merged['source_mode'] = 'ai'
    return merged


def normalize_executive_briefing_ai_payload(ai_payload):
    if not isinstance(ai_payload, dict):
        return None

    normalized = {
        'de_que_trata': sanitize_executive_text(ai_payload.get('de_que_trata'), max_len=140),
        'datos_tecnicos': unique_non_empty_texts(
            [sanitize_executive_text(item, max_len=140) for item in (ai_payload.get('datos_tecnicos') or [])],
            limit=2,
        ),
        'planeamiento': unique_non_empty_texts(
            [sanitize_executive_text(item, max_len=140) for item in (ai_payload.get('planeamiento') or [])],
            limit=2,
        ),
        'puntos_a_tratar': unique_non_empty_texts(
            [sanitize_executive_text(item, max_len=140) for item in (ai_payload.get('puntos_a_tratar') or [])],
            limit=2,
        ),
        'breve_resumen': unique_non_empty_texts(
            [sanitize_executive_text(item, max_len=140) for item in (ai_payload.get('breve_resumen') or [])],
            limit=2,
        ),
        'objetivos': unique_non_empty_texts(
            [sanitize_executive_text(item, max_len=140) for item in (ai_payload.get('objetivos') or [])],
            limit=2,
        ),
        'elementos_prioritarios': unique_non_empty_texts(
            [sanitize_executive_text(item, max_len=140) for item in (ai_payload.get('elementos_prioritarios') or [])],
            limit=2,
        ),
    }
    if not any(normalized.values()):
        return None
    return normalized


def normalize_visual_plan_ai_payload(ai_payload):
    if not isinstance(ai_payload, dict):
        return None

    charts = []
    for item in ai_payload.get('charts') or []:
        if not isinstance(item, dict):
            continue
        visual_id = sanitize_executive_text(item.get('id'), max_len=80)
        message = sanitize_executive_text(item.get('mensaje_clave'), max_len=160)
        rationale = sanitize_executive_text(item.get('por_que_importa'), max_len=160)
        if visual_id:
            charts.append({
                'id': visual_id,
                'mensaje_clave': message,
                'por_que_importa': rationale,
            })

    tables = []
    for item in ai_payload.get('tables') or []:
        if not isinstance(item, dict):
            continue
        visual_id = sanitize_executive_text(item.get('id'), max_len=80)
        mode = sanitize_executive_text(item.get('modo'), max_len=20).lower()
        message = sanitize_executive_text(item.get('mensaje_clave'), max_len=160)
        rationale = sanitize_executive_text(item.get('por_que_importa'), max_len=160)
        if visual_id and mode in {'summary', 'detail', 'omit'}:
            tables.append({
                'id': visual_id,
                'modo': mode,
                'mensaje_clave': message,
                'por_que_importa': rationale,
            })

    storyline = unique_non_empty_texts(
        [sanitize_executive_text(item, max_len=160) for item in (ai_payload.get('storyline') or [])],
        limit=3,
    )

    if not charts and not tables and not storyline:
        return None
    return {
        'charts': charts[:4],
        'tables': tables[:4],
        'storyline': storyline,
    }


def attach_visual_ids(resultado):
    for idx, chart in enumerate(resultado.get('graficas_automaticas') or []):
        if isinstance(chart, dict):
            chart['_visual_ai_id'] = f'chart:auto:{idx}'

    main_table = resultado.get('muestra_tabla')
    if isinstance(main_table, dict):
        main_table['_visual_ai_id'] = 'table:main'

    for key, table in (resultado.get('otras_tablas') or {}).items():
        if isinstance(table, dict):
            table['_visual_ai_id'] = f"table:other:{normalize_semantic_text(key) or 'item'}"

    for key, table in (resultado.get('genericas') or {}).items():
        if isinstance(table, dict):
            table['_visual_ai_id'] = f"table:generic:{normalize_semantic_text(key) or 'item'}"


def build_visual_candidates_for_ai(analisis_data):
    chart_candidates = []
    for idx, chart in enumerate((analisis_data.get('graficas_automaticas') or [])[:4]):
        if not isinstance(chart, dict):
            continue
        labels = [sanitize_executive_text(label, max_len=30) for label in (chart.get('labels') or [])[:4]]
        values = (chart.get('valores') or [])[:4]
        if len(labels) < 3 or len(values) < 3:
            continue
        chart_candidates.append({
            'id': chart.get('_visual_ai_id') or f'chart:auto:{idx}',
            'titulo': sanitize_executive_text(chart.get('titulo'), max_len=80),
            'tipo': sanitize_executive_text(chart.get('tipo'), max_len=20),
            'labels_top': labels,
            'valores_top': values,
            'insight_base': sanitize_executive_text(chart.get('insight_auto'), max_len=140),
        })

    table_candidates = []
    candidate_tables = []
    if isinstance(analisis_data.get('muestra_tabla'), dict):
        candidate_tables.append(('table:main', analisis_data.get('muestra_tabla'), 'summary'))
    for key, table in list((analisis_data.get('otras_tablas') or {}).items())[:3]:
        candidate_tables.append((table.get('_visual_ai_id') or f"table:other:{normalize_semantic_text(key) or 'item'}", table, 'summary'))
    for key, table in list((analisis_data.get('genericas') or {}).items())[:3]:
        candidate_tables.append((table.get('_visual_ai_id') or f"table:generic:{normalize_semantic_text(key) or 'item'}", table, 'detail'))

    for visual_id, table, suggested_mode in candidate_tables[:6]:
        if not isinstance(table, dict):
            continue
        headers = [sanitize_executive_text(item, max_len=32) for item in (table.get('encabezados') or [])[:6]]
        rows = table.get('filas') or []
        if len(headers) < 2 or len(rows) < 2:
            continue
        sample_rows = []
        for row in rows[:2]:
            if isinstance(row, (list, tuple)):
                sample_rows.append([sanitize_executive_text(value, max_len=40) for value in row[:5]])
        table_candidates.append({
            'id': visual_id,
            'titulo': sanitize_executive_text(table.get('hoja_origen') or table.get('title') or visual_id, max_len=80),
            'encabezados': headers,
            'muestra_filas': sample_rows,
            'total_filas': len(rows),
            'total_columnas': len(headers),
            'modo_sugerido': suggested_mode,
        })

    return {
        'charts': chart_candidates,
        'tables': table_candidates,
    }


def extract_prompt_focus_terms(prompt, limit=8):
    normalized = normalize_semantic_text(prompt)
    if not normalized:
        return []

    terms = []
    seen = set()
    for token in re.findall(r'[a-z0-9]{4,}', normalized):
        if token in LOCAL_PROMPT_STOPWORDS or token.isdigit():
            continue
        if token in seen:
            continue
        seen.add(token)
        terms.append(token)
        if len(terms) >= limit:
            break
    return terms


def describe_focus_label(request_context, analisis_data):
    focus_terms = extract_prompt_focus_terms((request_context or {}).get('prompt'))
    if focus_terms:
        return focus_terms[0]

    generico = analisis_data.get('resumen_generico') or {}
    metadatos = analisis_data.get('metadatos') or {}
    fallback = generico.get('hoja_principal') or metadatos.get('hoja_principal') or metadatos.get('archivo')
    return sanitize_executive_text(fallback, max_len=48) or 'los datos priorizados'


def score_visual_candidate_for_prompt(candidate, focus_terms):
    haystack = normalize_semantic_text(" ".join([
        str(candidate.get('titulo') or ''),
        " ".join(str(item or '') for item in (candidate.get('labels_top') or [])),
        " ".join(str(item or '') for item in (candidate.get('encabezados') or [])),
        str(candidate.get('insight_base') or ''),
    ]))
    score = 1
    for term in focus_terms:
        if term in haystack:
            score += 4
    if candidate.get('insight_base'):
        score += 2
    if candidate.get('modo_sugerido') == 'summary':
        score += 1
    return score


def build_local_storyline(analisis_data, visual_plan=None):
    chart_messages = [
        sanitize_executive_text(item.get('mensaje_clave'), max_len=150)
        for item in ((visual_plan or {}).get('charts') or [])
        if isinstance(item, dict)
    ]
    table_messages = [
        sanitize_executive_text(item.get('mensaje_clave'), max_len=150)
        for item in ((visual_plan or {}).get('tables') or [])
        if isinstance(item, dict) and item.get('modo') != 'omit'
    ]
    conclusions = [
        sanitize_executive_text(item, max_len=160)
        for item in (analisis_data.get('conclusiones') or [])[:4]
    ]
    insights = [
        sanitize_executive_text(item.get('texto') if isinstance(item, dict) else item, max_len=160)
        for item in ((analisis_data.get('analisis_avanzado') or {}).get('insights') or [])[:3]
    ]
    return unique_non_empty_texts(chart_messages + table_messages + conclusions + insights, limit=3)


def build_local_visual_plan(analisis_data):
    request_context = analisis_data.get('presentation_request') or {}
    focus_terms = extract_prompt_focus_terms(request_context.get('prompt'))
    focus_label = describe_focus_label(request_context, analisis_data)
    prompt_normalized = normalize_semantic_text(request_context.get('prompt'))
    candidates = build_visual_candidates_for_ai(analisis_data)

    charts_ranked = sorted(
        (candidates.get('charts') or []),
        key=lambda item: score_visual_candidate_for_prompt(item, focus_terms),
        reverse=True,
    )
    tables_ranked = sorted(
        (candidates.get('tables') or []),
        key=lambda item: score_visual_candidate_for_prompt(item, focus_terms),
        reverse=True,
    )

    chart_limit = 2 if charts_ranked else 0
    table_limit = 2 if tables_ranked else 0
    if any(token in prompt_normalized for token in ('grafica', 'graficas', 'tendencia', 'comparativa', 'comparativo')):
        chart_limit = min(3, len(charts_ranked))
    if any(token in prompt_normalized for token in ('tabla', 'tablas', 'detalle', 'hallazgo', 'hallazgos', 'riesgo', 'riesgos')):
        table_limit = min(3, len(tables_ranked))

    charts = []
    for candidate in charts_ranked[:chart_limit]:
        title = sanitize_executive_text(candidate.get('titulo'), max_len=70) or 'Visual clave'
        charts.append({
            'id': candidate.get('id'),
            'mensaje_clave': sanitize_executive_text(
                candidate.get('insight_base') or f"{title} concentra la lectura ejecutiva sobre {focus_label}.",
                max_len=160,
            ),
            'por_que_importa': sanitize_executive_text(
                f"Se prioriza porque conecta el Excel con el foco solicitado sobre {focus_label}.",
                max_len=160,
            ),
        })

    tables = []
    for candidate in tables_ranked[:table_limit]:
        detail_requested = any(token in prompt_normalized for token in ('tabla', 'tablas', 'detalle', 'hallazgo', 'hallazgos', 'riesgo', 'riesgos'))
        mode = 'detail' if detail_requested or (candidate.get('total_filas') or 0) > 12 else 'summary'
        headers = ", ".join((candidate.get('encabezados') or [])[:3])
        tables.append({
            'id': candidate.get('id'),
            'modo': mode,
            'mensaje_clave': sanitize_executive_text(
                f"La tabla resume evidencia puntual de {focus_label}" + (f" con columnas {headers}." if headers else "."),
                max_len=160,
            ),
            'por_que_importa': sanitize_executive_text(
                "Sirve como soporte verificable para la narrativa priorizada del archivo.",
                max_len=160,
            ),
        })

    storyline = build_local_storyline(analisis_data, {'charts': charts, 'tables': tables})
    return normalize_visual_plan_ai_payload({
        'charts': charts,
        'tables': tables,
        'storyline': storyline,
    })


def build_local_recommendation_text(request_context, focus_label):
    prompt_normalized = normalize_semantic_text((request_context or {}).get('prompt'))
    if any(token in prompt_normalized for token in ('riesgo', 'riesgos', 'hallazgo', 'hallazgos', 'control')):
        return sanitize_executive_text(
            f"Priorizar acciones de control y seguimiento sobre {focus_label} con evidencia trazable del Excel.",
            max_len=160,
        )
    if any(token in prompt_normalized for token in ('costo', 'costos', 'gasto', 'gastos', 'presupuesto', 'monto', 'montos')):
        return sanitize_executive_text(
            f"Concentrar la decision en los montos y concentraciones que explican la mayor parte del comportamiento de {focus_label}.",
            max_len=160,
        )
    return sanitize_executive_text(
        f"Usar la evidencia priorizada para decidir los siguientes pasos sobre {focus_label}.",
        max_len=160,
    )


def chart_index_from_visual_id(visual_id):
    match = re.search(r'chart:auto:(\d+)', str(visual_id or ''))
    return int(match.group(1)) if match else None


def build_local_executive_summary(analisis_data, visual_plan):
    request_context = analisis_data.get('presentation_request') or {}
    focus_label = describe_focus_label(request_context, analisis_data)
    storyline = build_local_storyline(analisis_data, visual_plan)
    conclusions = [
        sanitize_executive_text(item, max_len=160)
        for item in (analisis_data.get('conclusiones') or [])[:4]
    ]
    table_message = next((
        sanitize_executive_text(item.get('mensaje_clave'), max_len=160)
        for item in ((visual_plan or {}).get('tables') or [])
        if isinstance(item, dict) and item.get('modo') != 'omit'
    ), '')

    insights_graficas = []
    for item in ((visual_plan or {}).get('charts') or [])[:3]:
        idx = chart_index_from_visual_id(item.get('id'))
        if idx is None:
            continue
        insight = sanitize_executive_text(item.get('mensaje_clave'), max_len=160)
        if insight:
            insights_graficas.append({'id': idx, 'insight': insight})

    return {
        'vision_general': storyline[0] if storyline else (
            conclusions[0] if conclusions else sanitize_executive_text(
                f"Lectura ejecutiva construida desde datos reales del Excel sobre {focus_label}.",
                max_len=160,
            )
        ),
        'alerta_principal': storyline[1] if len(storyline) > 1 else (
            conclusions[1] if len(conclusions) > 1 else sanitize_executive_text(
                f"El archivo requiere priorizacion para evitar una lectura dispersa de {focus_label}.",
                max_len=160,
            )
        ),
        'recomendacion': build_local_recommendation_text(request_context, focus_label),
        'insight_tabla': table_message or sanitize_executive_text(
            f"La tabla principal sirve como base verificable para profundizar en {focus_label}.",
            max_len=160,
        ),
        'insights_graficas': insights_graficas,
    }


def build_local_briefing_payload(analisis_data, visual_plan):
    request_context = analisis_data.get('presentation_request') or {}
    generico = analisis_data.get('resumen_generico') or {}
    metadatos = analisis_data.get('metadatos') or {}
    focus_label = describe_focus_label(request_context, analisis_data)
    columns = [
        sanitize_executive_text(item, max_len=36)
        for item in (generico.get('columnas') or [])[:4]
    ]
    storyline = build_local_storyline(analisis_data, visual_plan)
    payload = {
        'de_que_trata': sanitize_executive_text(
            f"Lectura ejecutiva de {metadatos.get('archivo') or 'Excel'} con foco en {focus_label}.",
            max_len=160,
        ),
        'datos_tecnicos': unique_non_empty_texts([
            sanitize_executive_text(f"Hoja principal: {generico.get('hoja_principal') or metadatos.get('hoja_principal')}", max_len=120),
            sanitize_executive_text(f"Volumen analizado: {generico.get('total_filas') or 0} filas y {generico.get('total_columnas') or 0} columnas.", max_len=120),
            sanitize_executive_text(f"Columnas priorizadas: {', '.join(columns)}", max_len=120) if columns else '',
        ], limit=3),
        'planeamiento': unique_non_empty_texts([
            sanitize_executive_text(f"Abrir con el foco ejecutivo en {focus_label}.", max_len=120),
            sanitize_executive_text("Presentar solo evidencia con trazabilidad a hojas y columnas reales.", max_len=120),
            sanitize_executive_text("Cerrar con una decision o accion concreta sustentada en los hallazgos.", max_len=120),
        ], limit=3),
        'puntos_a_tratar': unique_non_empty_texts(storyline or [
            sanitize_executive_text(f"Priorizar la lectura mas util sobre {focus_label}.", max_len=120),
        ], limit=3),
        'breve_resumen': unique_non_empty_texts((analisis_data.get('conclusiones') or [])[:3], limit=3),
        'objetivos': unique_non_empty_texts([
            sanitize_executive_text(f"Responder el foco solicitado sobre {focus_label}.", max_len=120),
            sanitize_executive_text(f"Traducir el Excel en una narrativa clara para {request_context.get('audience') or 'ejecutivos'}.", max_len=120),
        ], limit=3),
        'elementos_prioritarios': unique_non_empty_texts([
            sanitize_executive_text(item.get('mensaje_clave'), max_len=120)
            for item in ((visual_plan or {}).get('charts') or [])[:2]
        ] + [
            sanitize_executive_text(item.get('mensaje_clave'), max_len=120)
            for item in ((visual_plan or {}).get('tables') or [])[:2]
            if isinstance(item, dict) and item.get('modo') != 'omit'
        ], limit=4),
    }
    return normalize_executive_briefing_ai_payload(payload)


def build_local_ai_curation_bundle(analisis_data):
    visual_plan = build_local_visual_plan(analisis_data)
    if not visual_plan:
        return None
    return {
        'resumen_ejecutivo_ia': build_local_executive_summary(analisis_data, visual_plan),
        'briefing_ejecutivo_ia': build_local_briefing_payload(analisis_data, visual_plan),
        'visual_plan_ia': visual_plan,
    }


def enrich_textual_blocks_with_ai(blocks):
    if not blocks:
        return []

    if is_ai_temporarily_blocked():
        return blocks

    api_key = get_openrouter_api_key()
    if not api_key:
        return blocks

    cache = load_textual_ai_cache()
    enriched = []
    pending = []
    for block in blocks:
        cache_key = build_textual_block_cache_key(block)
        cached = cache.get(cache_key)
        if isinstance(cached, dict):
            merged = merge_ai_textual_block(block, cached)
            merged['cache_hit'] = True
            enriched.append(merged)
        else:
            block_copy = dict(block)
            block_copy['_cache_key'] = cache_key
            pending.append(block_copy)
            enriched.append(block_copy)
            
    pending_for_ai = [block for block in pending if block.get('sheet_family') in TEXTUAL_FAMILIES][:MAX_TEXTUAL_AI_BLOCKS]
    if not pending_for_ai:
        for block in enriched:
            if isinstance(block, dict):
                block.pop('_cache_key', None)
        return enriched

    prompt_payload = []
    for index, block in enumerate(pending_for_ai):
        prompt_payload.append({
            'id': index,
            'title': block.get('title'),
            'subtitle': block.get('subtitle'),
            'sheet_family': block.get('sheet_family'),
            'lines': block.get('lines') or [],
        })

    prompt = f"""
    Eres consultor senior de auditoria y control interno.
    Reescribe estos bloques narrativos para PowerPoint ejecutivo manteniendo SOLO hechos presentes.
    No inventes cifras, estados, responsables ni conclusiones no evidenciadas.
    Conserva el sentido de hallazgos, procedimiento, riesgos y oportunidades.
    Devuelve como maximo {MAX_TEXT_LINES_PER_BLOCK} lineas por bloque, en tono profesional y humano.

    BLOQUES:
    {json.dumps(prompt_payload, ensure_ascii=False)}

    Responde exclusivamente con este JSON:
    {{
      "blocks": [
        {{
          "id": 0,
          "title": "Titulo breve",
          "subtitle": "Subtitulo ejecutivo",
          "lines": ["Linea 1", "Linea 2"]
        }}
      ]
    }}
    """

    response_text = call_ai_api(prompt)
    if response_text:
        try:
            # Robust extraction of JSON from markdown or text
            json_match = re.search(r'(\{.*\})', response_text.replace('\n', ' '), re.DOTALL)
            if json_match:
                response_text = json_match.group(1)
            else:
                response_text = response_text.replace("```json", "").replace("```", "").strip()
            
            parsed = json.loads(response_text)
            if isinstance(parsed, dict) and 'blocks' in parsed:
                for item in parsed.get('blocks', []):
                    if not isinstance(item, dict): continue
                    item_id = item.get('id')
                    if not isinstance(item_id, int) or not (0 <= item_id < len(pending_for_ai)):
                        continue
                    cache_key = pending_for_ai[item_id].get('_cache_key')
                    if cache_key:
                        cache[cache_key] = {
                            'title': item.get('title'),
                            'subtitle': item.get('subtitle'),
                            'lines': item.get('lines') or [],
                        }
                save_textual_ai_cache(cache)
        except Exception as exc:
            print(f"INFO: Error parseando IA textual: {exc}", file=sys.stderr)

    final_blocks = []
    for block in enriched:
        if not isinstance(block, dict):
            final_blocks.append(block)
            continue
            
        cache_key = block.get('_cache_key')
        if cache_key and cache_key in cache:
            merged = merge_ai_textual_block(block, cache.get(cache_key))
            merged['cache_hit'] = block.get('cache_hit', False)
            merged.pop('_cache_key', None)
            final_blocks.append(merged)
        else:
            block.pop('_cache_key', None)
            final_blocks.append(block)
    return final_blocks


def call_ai_api(prompt, response_mime_type="application/json"):
    api_key = get_openrouter_api_key()
    if not api_key:
        return None
    best_effort_mode = AI_EXECUTION_MODE != "blocking"
    wait_budget_seconds = 0 if best_effort_mode else AI_MAX_WAIT_SECONDS
    request_timeout_seconds = max(10, min(AI_REQUEST_TIMEOUT_SECONDS, AI_HARD_DEADLINE_SECONDS)) if best_effort_mode else AI_REQUEST_TIMEOUT_SECONDS
    max_retries = 1 if best_effort_mode else AI_MAX_RETRIES
    if is_ai_temporarily_blocked():
        if best_effort_mode:
            return None
        waited, wait_budget_seconds = maybe_wait_for_ai_cooldown(wait_budget_seconds)
        if waited <= 0 and is_ai_temporarily_blocked():
            return None

    import requests
    system_msg = "You must output strictly valid JSON." if response_mime_type == "application/json" else "You are a helpful assistant."
    payload = {
        "model": OPENROUTER_MODEL_PRIORITY[0],
        "messages": [
            {"role": "system", "content": system_msg},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.2 if response_mime_type == "application/json" else 0.4,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "HTTP-Referer": OPENROUTER_SITE_URL,
        "X-Title": OPENROUTER_APP_NAME,
    }

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.post(
                url="https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                json=payload,
                timeout=request_timeout_seconds,
            )
            response_text = resp.text or ""

            if resp.status_code == 429:
                mark_ai_quota_blocked(response_text)
                if best_effort_mode:
                    return None
                waited, wait_budget_seconds = maybe_wait_for_ai_cooldown(wait_budget_seconds)
                if waited > 0 and attempt < max_retries:
                    continue
                return None

            if not resp.ok:
                error_message = extract_openrouter_error_message(
                    resp.json() if "application/json" in str(resp.headers.get("content-type", "")).lower() else response_text
                )
                if is_openrouter_rate_limit_message(error_message):
                    mark_ai_quota_blocked(error_message)
                    if best_effort_mode:
                        return None
                    waited, wait_budget_seconds = maybe_wait_for_ai_cooldown(wait_budget_seconds)
                    if waited > 0 and attempt < max_retries:
                        continue
                    return None
                if is_transient_openrouter_message(error_message) and attempt < max_retries:
                    time.sleep(min(2 ** (attempt - 1), 6))
                    continue
                print(f"INFO: Error HTTP OpenRouter ({resp.status_code}): {error_message}", file=sys.stderr)
                return None

            data = resp.json()
            if isinstance(data, dict) and 'choices' in data and len(data['choices']) > 0:
                content = ((data['choices'][0] or {}).get('message') or {}).get('content')
                normalized = normalize_ai_json_response_text(content) if response_mime_type == "application/json" else content
                if normalized:
                    clear_ai_quota_block()
                    return normalized
                last_error = "respuesta sin contenido util"
            else:
                error_message = extract_openrouter_error_message(data)
                if is_openrouter_rate_limit_message(error_message):
                    mark_ai_quota_blocked(error_message)
                    return None
                last_error = error_message or "respuesta OpenRouter sin choices"
        except Exception as exc:
            if is_openrouter_rate_limit_message(exc):
                mark_ai_quota_blocked(exc)
                if best_effort_mode:
                    return None
                waited, wait_budget_seconds = maybe_wait_for_ai_cooldown(wait_budget_seconds)
                if waited > 0 and attempt < max_retries:
                    continue
                return None
            last_error = str(exc)
            if is_transient_openrouter_message(exc) and attempt < max_retries:
                time.sleep(min(2 ** (attempt - 1), 6))
                continue
            if attempt < max_retries:
                time.sleep(min(2 ** (attempt - 1), 6))
                continue

    if last_error:
        print(f"INFO: Error cargando OpenRouter: {last_error}", file=sys.stderr)
    return None


def build_data_block_inventory(resultado):
    inventory = []
    if resultado.get('muestra_tabla'):
        inventory.append({'type': 'table', 'name': 'Tabla principal', 'source': resultado['muestra_tabla'].get('hoja_origen')})
    for chart in resultado.get('graficas_automaticas', []) or []:
        inventory.append({'type': 'chart', 'name': chart.get('titulo'), 'source': chart.get('hoja_origen')})
    for name, table in (resultado.get('otras_tablas') or {}).items():
        inventory.append({'type': 'table', 'name': name, 'source': table.get('hoja_origen')})
    for name, table in (resultado.get('genericas') or {}).items():
        inventory.append({'type': 'table', 'name': name, 'source': table.get('hoja_origen')})
    return inventory


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
    name_l = normalize_semantic_text(name)
    rows, cols = df.shape
    score = rows * cols
    if cols < 2 or rows < 2:
        return -1
    family = classify_sheet_family(name, df)
    score += PRIMARY_SHEET_FAMILY_SCORES.get(family, 0)
    if any(k in name_l for k in ['muestra total', 'consolidado total', 'resumen ejecutivo', 'tabla principal']):
        score += 500
    if any(k in name_l for k in ['ventas', 'inventario', 'stock', 'productos', 'resumen', 'dashboard', 'datos']):
        score += 180
    if any(k in name_l for k in ['hallazgo', 'oportunidad', 'mejora', 'coso', 'td', 'distribucion']):
        score -= 200
    headers = [normalize_semantic_text(c) for c in df.columns if pd.notna(c)]
    if any(h in headers for h in ['solicitante', 'valor total solicitado', 'valor total', 'ventas totales']):
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
    return any(token in nombre for token in ['id', 'codigo', 'consecutivo', 'numero', 'nro', 'folio', 'radicado', 'centro', 'comprobante'])


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


def is_temporal_header(col):
    normalized = normalize_semantic_text(col)
    if not normalized:
        return False
    return any(
        token in normalized
        for token in ('fecha', 'date', 'periodo', 'period', 'mes', 'month', 'semana', 'week', 'dia', 'day', 'anio', 'ano', 'year')
    )


def should_prioritize_primary_story(primary_sheet_name, primary_df):
    if primary_df is None or getattr(primary_df, 'empty', True):
        return False
    rows, cols = primary_df.shape
    if rows < 150 or cols < 5:
        return False
    primary_family = classify_sheet_family(primary_sheet_name, primary_df)
    if primary_family != 'general':
        return False
    signature = build_sheet_semantic_signature(primary_sheet_name, primary_df)
    meaningful_tokens = ('comision', 'valor', 'costo', 'gasto', 'centro', 'ciudad', 'fecha', 'viaje', 'solicit')
    return any(token in signature for token in meaningful_tokens)


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

    if is_temporal_header(col):
        return {
            'kind': 'temporal',
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
    if profile.get('kind') in {'temporal', 'identifier'}:
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
        is_temporal = is_temporal_header(col)
        is_identifier_header = es_columna_identificador(col)
        is_exec_dimension = es_dimension_ejecutiva(col)
        
        # Detectar tipo
        serie_num = normalize_numeric_series(serie, col)
        ratio_num = serie_num.notna().sum() / max(1, len(serie))
        unique_vals = serie.astype(str).nunique()
        total_vals = len(serie)
        ratio_unique = unique_vals / max(1, total_vals)

        if is_temporal:
            info['tipo'] = 'temporal'
            info['importancia'] += 8
        elif ratio_num >= 0.6 and not is_identifier_header:
            info['tipo'] = 'numerico'
            info['importancia'] += 20
            stats = analizar_columna_numerica(df, col)
            if stats:
                info['stats'] = stats
                if is_probable_numeric_identifier(serie_num.dropna().tolist()) and not is_exec_dimension:
                    info['tipo'] = 'identificador'
                    info['importancia'] -= 18
                elif ratio_unique > 0.92 and infer_numeric_kind(col) == 'number' and not is_exec_dimension:
                    info['tipo'] = 'identificador'
                    info['importancia'] -= 14
                if stats['total'] > 1000000:
                    info['importancia'] += 30
                elif stats['total'] > 10000:
                    info['importancia'] += 15
        else:
            if is_exec_dimension and unique_vals >= 2 and unique_vals <= min(250, max(15, int(total_vals * 0.45))):
                info['tipo'] = 'categorica'
                info['importancia'] += 18
                info['valores_unicos'] = unique_vals
            elif is_identifier_header and not is_exec_dimension:
                info['tipo'] = 'identificador'
                info['importancia'] += 5
            elif ratio_unique <= 0.3 and unique_vals <= 20:
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

def generar_conclusiones(df, cols_info, kpis, _unused_specialization=False,
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
# PRESUPUESTO DE SLIDES INTELIGENTE
# ═══════════════════════════════════════════════════════════════════════════════

def calcular_presupuesto_slides(resultado, _unused_specialization=False):
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
        'genericas': 0,
        'conclusiones': 0,
        'cierre': 1
    }
    
    slots_disponibles = MAX_SLIDES - 2  # -2 por portada y cierre
    
    # Prioridad 1: Resumen ejecutivo / KPIs (siempre)
    if resultado.get('resumen_ejecutivo') or resultado.get('resumen_generico') or resultado.get('kpis_automaticos'):
        presupuesto['resumen_kpis'] = 1
        slots_disponibles -= 1
    
    total_hojas = len(resultado.get('metadatos', {}).get('hojas_encontradas', []) or [])
    has_textual_blocks = bool(resultado.get('bloques_textuales'))

    # Prioridad 2: Gráficas (reservar al menos 1 si hay graficas reales)
    num_graficas = len(resultado.get('graficas_automaticas', []))
    graficas_slots = min(num_graficas, MAX_AUTO_CHARTS, slots_disponibles)
    if num_graficas > 0 and slots_disponibles > 0:
        graficas_slots = max(1, graficas_slots)
    presupuesto['graficas'] = graficas_slots
    slots_disponibles -= graficas_slots
    
    # Prioridad 3: Tabla principal (máximo 2 páginas, ESTRICTO)
    num_filas = len(resultado.get('muestra_tabla', {}).get('filas', []))
    if num_filas > 0 and slots_disponibles > 0:
        tabla_pages = min(2, max(1, num_filas // ROWS_PER_TABLE_SLIDE + (1 if num_filas % ROWS_PER_TABLE_SLIDE else 0)))
        tabla_pages = min(tabla_pages, slots_disponibles)
        presupuesto['tabla_principal'] = tabla_pages
        slots_disponibles -= tabla_pages
    
    # Prioridad 6: Conclusiones (siempre reservamos espacio)
    if (resultado.get('conclusiones') or resultado.get('analisis_avanzado')) and slots_disponibles > 0:
        presupuesto['conclusiones'] = 1
        slots_disponibles -= 1
    
    # Prioridad 7: Estructura del archivo
    if resultado.get('metadatos', {}).get('hojas_encontradas') and slots_disponibles > 0:
        presupuesto['estructura'] = 1
        slots_disponibles -= 1
    
    # Prioridad 8: Tablas adicionales genéricas
    num_genericas = len(resultado.get('otras_tablas', {})) + len(resultado.get('genericas', {}))
    if num_genericas > 0 and slots_disponibles > 0:
        cap_genericas = 4 if has_textual_blocks or total_hojas >= 10 else 3
        gen_slots = min(num_genericas, cap_genericas, slots_disponibles)
        presupuesto['genericas'] = gen_slots
        slots_disponibles -= gen_slots
    
    presupuesto['slots_restantes'] = slots_disponibles
    presupuesto['total_estimado'] = MAX_SLIDES - slots_disponibles
    
    return presupuesto


# ═══════════════════════════════════════════════════════════════════════════════
# PIPELINE PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

import requests

def sintetizar_resumen_ia(analisis_data):
    """
    Sintetiza un resumen ejecutivo usando SOLO conclusiones, tablas y graficas ya calculadas por Python.
    """
    conclusiones = analisis_data.get('conclusiones', [])
    insights = analisis_data.get('analisis_avanzado', {}).get('insights', [])
    
    textos_base = conclusiones[:5] + insights[:3]
    if not textos_base:
        return None

    graficas = analisis_data.get('graficas_automaticas', [])
    info_graficas = []
    for i, g in enumerate(graficas):
        info_graficas.append({
            "id": i,
            "titulo": g.get('titulo', ''),
            "labels_top_3": g.get('labels', [])[:3],
            "valores_top_3": g.get('valores', [])[:3]
        })

    info_tabla = ""
    if 'muestra_tabla' in analisis_data:
        headers = analisis_data['muestra_tabla'].get('encabezados', [])
        rows = analisis_data['muestra_tabla'].get('filas', [])[:3]
        info_tabla = f"Tabla Principal: Encabezados: {headers}, Top 3 Filas: {rows}"

    prompt = f"""
    Eres un Consultor Senior. Sintetiza estos hallazgos estadísticos reales en un breve Resumen Ejecutivo para la presentación corporativa.
    Adicionalmente, redacta 1 línea de "insight" de negocio brillante para cada gráfica proporcionada, y 1 para la tabla principal.
    
    HALLAZGOS BASE:
    {json.dumps(textos_base, ensure_ascii=False)}

    DATOS DE GRÁFICAS:
    {json.dumps(info_graficas, ensure_ascii=False)}

    DATOS DE TABLAS:
    {info_tabla}

    Responde ESTRICTAMENTE con este JSON sin formato markdown extra:
    {{
      "vision_general": "Párrafo resumiendo la situación (máx 40 palabras)",
      "alerta_principal": "El riesgo, anomalía o concentración más crítica detectada (máx 25 palabras)",
      "recomendacion": "Acción estratégica a tomar (máx 20 palabras)",
      "insight_tabla": "Insight de negocio analizando los datos de la tabla (máx 20 palabras)",
      "insights_graficas": [
          {{"id": 0, "insight": "Fuerte concentración en X, indicando oportunidad de mejora."}}
      ]
    }}
    """

    response_text = call_ai_api(prompt)
    if response_text:
        try:
            # Robust extraction of JSON
            json_match = re.search(r'(\{.*\})', response_text.replace('\n', ' '), re.DOTALL)
            if json_match:
                response_text = json_match.group(1)
            else:
                response_text = response_text.replace("```json", "").replace("```", "").strip()
            
            parsed = json.loads(response_text)
            if isinstance(parsed, dict):
                return parsed
        except Exception as e:
            print(f"INFO: Error parseando resumen IA: {e}", file=sys.stderr)
    return None


def sintetizar_briefing_ejecutivo_ia(analisis_data):
    """
    Genera un briefing ejecutivo corto para las diapositivas 2 y 3.
    """
    metadatos = analisis_data.get('metadatos', {}) or {}
    resumen = analisis_data.get('resumen_ejecutivo', {}) or {}
    generico = analisis_data.get('resumen_generico', {}) or {}
    conclusiones = unique_non_empty_texts(analisis_data.get('conclusiones') or [], limit=4)
    insights = []
    for item in (analisis_data.get('analisis_avanzado', {}) or {}).get('insights', []) or []:
        if isinstance(item, dict):
            text = sanitize_executive_text(item.get('texto'), max_len=140)
        else:
            text = sanitize_executive_text(item, max_len=140)
        if text:
            insights.append(text)
    insights = unique_non_empty_texts(insights, limit=3)

    graficas = []
    for item in (analisis_data.get('graficas_automaticas') or [])[:3]:
        graficas.append({
            'titulo': sanitize_executive_text(item.get('titulo'), max_len=80),
            'insight': sanitize_executive_text(item.get('insight_auto'), max_len=120),
            'labels': [sanitize_executive_text(label, max_len=40) for label in (item.get('labels') or [])[:3]],
            'valores': (item.get('valores') or [])[:3],
        })

    tabla = None
    muestra = analisis_data.get('muestra_tabla') or {}
    if muestra:
        tabla = {
            'encabezados': [sanitize_executive_text(item, max_len=40) for item in (muestra.get('encabezados') or [])[:5]],
            'filas': [
                [sanitize_executive_text(value, max_len=60) for value in row[:5]]
                for row in (muestra.get('filas') or [])[:2]
                if isinstance(row, (list, tuple))
            ],
        }

    prompt_payload = {
        'archivo': metadatos.get('archivo'),
        'hoja_principal': generico.get('hoja_principal') or metadatos.get('hoja_principal'),
        'tipo_libro': metadatos.get('tipo_libro'),
        'familias_detectadas': metadatos.get('familias_detectadas') or [],
        'resumen_ejecutivo': resumen,
        'resumen_generico': generico,
        'conclusiones': conclusiones,
        'insights': insights,
        'graficas': graficas,
        'tabla_principal': tabla,
    }
    cache_key = build_executive_briefing_cache_key(prompt_payload)
    cache = load_executive_ai_cache()
    cached = normalize_executive_briefing_ai_payload(cache.get(cache_key))
    if cached:
        return cached

    prompt = f"""
    Eres consultor senior de gerencia y control interno.
    Debes redactar contenido ejecutivo de alto impacto para las diapositivas iniciales de un PowerPoint.
    Usa SOLO los hechos presentes en el JSON. No inventes cifras, procesos, objetivos, areas ni riesgos.
    El tono debe ser de comite de gerencia: sobrio, claro, accionable y orientado a decisiones.
    Responde en espanol.
    Cada bullet debe ser concreto, con lectura de negocio y sin jerga tecnica.
    Evita mencionar IDs, nombres de columnas, nombres de hojas, conteos tecnicos de tablas/graficas o detalles de sistema.
    Prioriza mensajes sobre concentracion, impacto financiero, riesgo, oportunidad y acciones de seguimiento.
    No uses puntos suspensivos (...).

    DATOS:
    {json.dumps(prompt_payload, ensure_ascii=False)}

    Responde EXCLUSIVAMENTE con este JSON:
    {{
      "de_que_trata": "frase de maximo 22 palabras con enfoque gerencial",
      "datos_tecnicos": ["bullet 1", "bullet 2"],
      "planeamiento": ["bullet 1", "bullet 2"],
      "puntos_a_tratar": ["bullet 1", "bullet 2"],
      "breve_resumen": ["bullet 1", "bullet 2"],
      "objetivos": ["bullet 1", "bullet 2"],
      "elementos_prioritarios": ["bullet 1", "bullet 2"]
    }}
    """

    response_text = call_ai_api(prompt)
    if response_text:
        try:
            # Robust extraction of JSON
            json_match = re.search(r'(\{.*\})', response_text.replace('\n', ' '), re.DOTALL)
            if json_match:
                response_text = json_match.group(1)
            else:
                response_text = response_text.replace("```json", "").replace("```", "").strip()
                
            parsed = json.loads(response_text)
            normalized = normalize_executive_briefing_ai_payload(parsed)
            if normalized:
                cache[cache_key] = normalized
                save_executive_ai_cache(cache)
            return normalized
        except Exception as exc:
            print(f"INFO: Error parseando briefing IA: {exc}", file=sys.stderr)
    return None


# ─────────────────────────────────────────────────────────────────────────────
# LLAMADA UNIFICADA: todo en 1 request → máximo ahorro de cuota
# ─────────────────────────────────────────────────────────────────────────────

def load_unified_ai_cache():
    try:
        if os.path.exists(UNIFIED_AI_CACHE_FILE):
            with open(UNIFIED_AI_CACHE_FILE, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {}


def save_unified_ai_cache(cache_data):
    try:
        os.makedirs(os.path.dirname(UNIFIED_AI_CACHE_FILE), exist_ok=True)
        with open(UNIFIED_AI_CACHE_FILE, 'w', encoding='utf-8') as fh:
            json.dump(cache_data, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass


def build_unified_cache_key(analisis_data):
    """Hash del contenido analítico esencial — misma clave si el mismo Excel."""
    request_context = analisis_data.get('presentation_request') or {}
    theme = request_context.get('theme') or {}
    fingerprint = {
        'archivo': (analisis_data.get('metadatos') or {}).get('archivo'),
        'conclusiones': (analisis_data.get('conclusiones') or [])[:5],
        'insights': [(i.get('texto') if isinstance(i, dict) else i)
                     for i in ((analisis_data.get('analisis_avanzado') or {}).get('insights') or [])[:3]],
        'graficas_titulos': [g.get('titulo') for g in (analisis_data.get('graficas_automaticas') or [])[:4]],
        'bloques_titulos': [b.get('title') for b in (analisis_data.get('bloques_textuales') or [])[:3]],
        'user_prompt': request_context.get('prompt'),
        'audience': request_context.get('audience'),
        'language': request_context.get('language'),
        'theme': theme.get('name'),
        'prompt_version': 'unified-v4-design',
    }
    raw = json.dumps(fingerprint, ensure_ascii=False, sort_keys=True)
    import hashlib
    return hashlib.sha256(raw.encode('utf-8')).hexdigest()


def sintetizar_todo_con_ia(analisis_data):
    """
    UNA sola llamada a la IA que produce:
    - resumen_ejecutivo_ia   (visión, alerta, recomendación, insights gráficas)
    - briefing_ejecutivo_ia  (diapositivas 2-3)
    - bloques_textuales enriquecidos (max 3)
    """
    cache = load_unified_ai_cache()
    cache_key = build_unified_cache_key(analisis_data)
    cached = cache.get(cache_key)
    if isinstance(cached, dict) and cached.get('resumen_ejecutivo_ia'):
        print("INFO: IA unificada — resultado desde cache.", file=sys.stderr)
        return cached

    # ── Construir payload compacto ──────────────────────────────────────────
    conclusiones = (analisis_data.get('conclusiones') or [])[:5]
    insights_raw = (analisis_data.get('analisis_avanzado') or {}).get('insights') or []
    insights = [(i.get('texto') if isinstance(i, dict) else i) for i in insights_raw][:3]
    textos_base = unique_non_empty_texts(
        [sanitize_executive_text(t, max_len=160) for t in conclusiones + insights], limit=6
    )

    graficas_info = []
    for idx, g in enumerate((analisis_data.get('graficas_automaticas') or [])[:4]):
        graficas_info.append({
            'id': idx,
            'titulo': sanitize_executive_text(g.get('titulo'), max_len=60),
            'top_labels': [sanitize_executive_text(l, max_len=30) for l in (g.get('labels') or [])[:3]],
            'top_valores': (g.get('valores') or [])[:3],
        })

    bloques_raw = []
    for b in (analisis_data.get('bloques_textuales') or [])[:MAX_TEXTUAL_AI_BLOCKS]:
        bloques_raw.append({
            'title': sanitize_executive_text(b.get('title'), max_len=60),
            'sheet_family': b.get('sheet_family'),
            'lines': [sanitize_executive_text(l, max_len=160) for l in (b.get('lines') or [])[:4]],
        })

    metadatos = analisis_data.get('metadatos') or {}
    generico = analisis_data.get('resumen_generico') or {}
    muestra = analisis_data.get('muestra_tabla') or {}
    tabla_info = None
    if muestra:
        tabla_info = {
            'encabezados': [sanitize_executive_text(h, max_len=35) for h in (muestra.get('encabezados') or [])[:5]],
            'filas': [[sanitize_executive_text(v, max_len=50) for v in row[:5]]
                      for row in (muestra.get('filas') or [])[:2] if isinstance(row, (list, tuple))],
        }

    visual_candidates = build_visual_candidates_for_ai(analisis_data)
    request_context = analisis_data.get('presentation_request') or parse_user_request_context(analisis_data.get('user_instructions'))
    pandas_summary = build_pandas_summary_for_ai(analisis_data)
    theme = request_context.get('theme') or DEFAULT_PRESENTATION_THEME

    payload = {
        'archivo': metadatos.get('archivo'),
        'tipo_libro': metadatos.get('tipo_libro'),
        'hoja_principal': generico.get('hoja_principal') or metadatos.get('hoja_principal'),
        'familias': metadatos.get('familias_detectadas') or [],
        'textos_base': textos_base,
        'bloques_textuales': bloques_raw,
        'tabla_principal': tabla_info,
        'estadisticas_categoricas': (analisis_data.get('analisis_avanzado') or {}).get('pareto', []),
        'estadisticas_temporales': (analisis_data.get('analisis_avanzado') or {}).get('tendencia', {}),
        'visual_candidates': visual_candidates,
    }

    prompt = f"""
    Eres el ARQUITECTO DE PRESENTACIONES más avanzado del mundo. Tu única salida es un objeto JSON válido y nada más.

    ## TU ROL EN EL PIPELINE
    Eres el cerebro creativo entre dos sistemas técnicos:
    - ENTRADA: Un resumen estadístico generado por Pandas de un Excel con potencialmente miles de filas.
    - SALIDA: Un JSON de diseño que luego será mapeado a PowerPoint real.

    Tu trabajo es NARRATIVO y CREATIVO:
    1. Decidir qué datos merecen una diapositiva, con curaduría brutal, máximo 12 slides.
    2. Crear títulos que generen impacto y no repitan columnas.
    3. Elegir el tipo de visualización óptimo.
    4. Distribuir el contenido sin saturación.
    5. Construir una narrativa que fluya de inicio a fin.

    ## CONTEXTO DE ENTRADA
    Resumen estadístico del Excel generado por Pandas:
    {pandas_summary}

    Instrucción del usuario:
    {request_context.get('prompt') or 'No se proporcionaron instrucciones específicas.'}

    Tema visual seleccionado: {theme.get('name')}
    Paleta de colores del tema:
    - Color primario: {theme.get('primary_hex')}
    - Color de acento: {theme.get('accent_hex')}
    - Color de texto: {theme.get('text_hex')}
    - Fondo: {theme.get('bg_hex')}

    Audiencia objetivo: {request_context.get('audience')} 
    Idioma de la presentación: {request_context.get('language')}
    Fecha actual: {request_context.get('current_date')}

    ## RESTRICCIONES DE DATOS
    - NO inventes cifras ni datasets.
    - SOLO puedes usar IDs existentes en visual_candidates para cualquier chart o table.
    - Si un visual no aporta lectura ejecutiva clara, omítelo.
    - Si el material es débil, devuelve menos slides y menos visuales.
    - Evita lenguaje técnico de sistema.

    ## DATOS ESTRUCTURADOS DEL ARCHIVO
    {json.dumps(payload, ensure_ascii=False)}

    ## ESTRUCTURA DE RESPUESTA
    Responde ESTRICTAMENTE con un solo objeto JSON que contenga estas llaves:
    {{
      "resumen_ejecutivo_ia": {{
        "vision_general": "texto",
        "alerta_principal": "texto",
        "recomendacion": "texto",
        "insight_tabla": "texto",
        "insights_graficas": [{{"id": 0, "insight": "texto"}}]
      }},
      "briefing_ejecutivo_ia": {{
        "de_que_trata": "texto",
        "datos_tecnicos": ["texto", "texto"],
        "planeamiento": ["texto", "texto"],
        "puntos_a_tratar": ["texto", "texto"],
        "breve_resumen": ["texto", "texto"],
        "objetivos": ["texto", "texto"],
        "elementos_prioritarios": ["texto", "texto"]
      }},
      "bloques_textuales_enriquecidos": [
        {{"title": "texto", "subtitle": "texto", "lines": ["texto"]}}
      ],
      "visual_plan_ia": {{
        "charts": [
          {{"id": "chart:auto:0", "mensaje_clave": "texto", "por_que_importa": "texto"}}
        ],
        "tables": [
          {{"id": "table:main", "modo": "summary|detail|omit", "mensaje_clave": "texto", "por_que_importa": "texto"}}
        ],
        "storyline": ["texto", "texto", "texto"]
      }},
      "presentation_design_ia": {{
        "presentation_meta": {{
          "title": "texto",
          "subtitle": "texto",
          "author": "Generado con Excel → PowerPoint Pro",
          "date": "{request_context.get('current_date')}",
          "total_slides": 0,
          "narrative_summary": "texto"
        }},
        "slides": [
          {{
            "slide_number": 1,
            "type": "title_slide",
            "layout": {{
              "title": {{"text": "texto"}},
              "subtitle": {{"text": "texto"}},
              "date": {{"text": "{request_context.get('current_date')}"}}
            }},
            "design_notes": "texto"
          }}
        ],
        "global_design": {{
          "font_primary": "Calibri",
          "font_secondary": "Calibri Light",
          "footer_text": "Confidencial - {request_context.get('current_date')}"
        }},
        "validation": {{
          "total_slides_matches_array": true,
          "all_slides_have_required_fields": true,
          "no_slide_exceeds_5_bullets": true,
          "no_chart_exceeds_8_series": true,
          "json_is_valid": true
        }}
      }}
    }}
    No uses markdown. JSON puro. Se brutal con la curaduría y enfócate en decisiones.
    """

    response_text = call_ai_api(prompt)
    if response_text:
        try:
            # Robust extraction of JSON
            json_match = re.search(r'(\{.*\})', response_text.replace('\n', ' '), re.DOTALL)
            if json_match:
                response_text = json_match.group(1)
            else:
                response_text = response_text.replace("```json", "").replace("```", "").strip()
            
            parsed = json.loads(response_text)
            
            # Validar estructura mínima
            if not isinstance(parsed, dict): return None
            
            result = {}
            if isinstance(parsed.get('resumen_ejecutivo_ia'), dict):
                result['resumen_ejecutivo_ia'] = parsed['resumen_ejecutivo_ia']
            if isinstance(parsed.get('briefing_ejecutivo_ia'), dict):
                result['briefing_ejecutivo_ia'] = normalize_executive_briefing_ai_payload(parsed['briefing_ejecutivo_ia'])
            if isinstance(parsed.get('bloques_textuales_enriquecidos'), list):
                result['bloques_textuales_enriquecidos'] = parsed['bloques_textuales_enriquecidos']
            visual_plan = normalize_visual_plan_ai_payload(parsed.get('visual_plan_ia'))
            if visual_plan:
                result['visual_plan_ia'] = visual_plan
            presentation_design = normalize_presentation_design_ai_payload(parsed.get('presentation_design_ia'))
            if presentation_design:
                result['presentation_design_ia'] = presentation_design
            
            if result:
                cache[cache_key] = result
                save_unified_ai_cache(cache)
                return result
        except Exception as exc:
            print(f"INFO: Error parseando IA unificada: {exc}", file=sys.stderr)
            
    return None


def generar_sugerencias_ia(analisis_data):
    """
    Genera sugerencias inteligentes y específicas para el usuario, basándose
    en los datos reales del Excel ya analizados.
    """
    cache = load_unified_ai_cache()
    cache_key = build_unified_cache_key(analisis_data) + "_sugerencias_v2"
    cached = cache.get(cache_key)
    if cached:
        return cached

    # ── Payload rico con datos reales del Excel ──
    metadatos = analisis_data.get('metadatos') or {}
    generico  = analisis_data.get('resumen_generico') or {}
    muestra   = analisis_data.get('muestra_tabla') or {}
    avanzado  = analisis_data.get('analisis_avanzado') or {}

    # Columnas de la hoja principal
    columnas = muestra.get('encabezados') or []

    # Muestra de filas (máx 3 filas, máx 6 columnas)
    filas_raw = muestra.get('filas') or []
    sample_rows = []
    for row in filas_raw[:3]:
        if isinstance(row, (list, tuple)):
            sample_rows.append([sanitize_executive_text(str(v), max_len=40) for v in row[:6]])

    # KPIs detectados automáticamente
    kpis = [k.get('label') for k in (analisis_data.get('kpis_automaticos') or []) if isinstance(k, dict)]

    # Pareto (columna más dominante por frecuencia)
    pareto = avanzado.get('pareto') or []
    pareto_summary = []
    for p in pareto[:3]:
        if isinstance(p, dict):
            pareto_summary.append({
                'columna': p.get('columna'),
                'top_valor': (p.get('top_valores') or [{}])[0].get('valor') if p.get('top_valores') else None,
                'top_pct': (p.get('top_valores') or [{}])[0].get('pct') if p.get('top_valores') else None,
            })

    # Tendencia temporal
    tendencia = avanzado.get('tendencia') or {}
    tendencia_info = {
        'columna_temporal': tendencia.get('col_temporal'),
        'columna_valor': tendencia.get('col_valor'),
        'tendencia_tipo': tendencia.get('tendencia'),
    } if tendencia else None

    # Bloques textuales detectados
    temas_textuales = [b.get('title') for b in (analisis_data.get('bloques_textuales') or [])[:4] if isinstance(b, dict)]

    payload = {
        'archivo': metadatos.get('archivo'),
        'hojas': metadatos.get('hojas_encontradas'),
        'hoja_principal': generico.get('hoja_principal') or metadatos.get('hoja_principal'),
        'total_filas': generico.get('total_filas'),
        'columnas': columnas[:12],
        'muestra_filas': sample_rows,
        'kpis_detectados': kpis[:8],
        'pareto': pareto_summary,
        'tendencia': tendencia_info,
        'temas_hojas': temas_textuales,
    }

    prompt = f"""Eres un Experto en Storytelling de Datos y Presentaciones Ejecutivas.
El usuario ha subido un Excel con la siguiente información real:

{json.dumps(payload, ensure_ascii=False, indent=2)}

Tu misión: Genera EXACTAMENTE 5 sugerencias breves, concretas y accionables que el usuario podría escribir como instrucciones para personalizar la generación de su PowerPoint. 

REGLAS:
- Cada sugerencia debe hacer referencia a datos REALES del Excel (columnas, valores, hojas).
- Sé específico: menciona nombres de columnas o métricas cuando sea relevante.
- Mezcla tipos: enfoque en gráficas, en tablas, en conclusiones, en comparativas, en alertas.
- Longitud: máximo 12 palabras por sugerencia.
- Idioma: español.

Responde ÚNICAMENTE con este JSON:
{{
  "sugerencias": ["sugerencia 1", "sugerencia 2", "sugerencia 3", "sugerencia 4", "sugerencia 5"]
}}"""

    response_text = call_ai_api(prompt)
    if response_text:
        try:
            json_match = re.search(r'(\{.*\})', response_text.replace('\n', ' '), re.DOTALL)
            clean = json_match.group(1) if json_match else response_text.replace("```json", "").replace("```", "").strip()
            parsed = json.loads(clean)
            if isinstance(parsed.get('sugerencias'), list) and len(parsed['sugerencias']) >= 2:
                result = {'sugerencias': [s for s in parsed['sugerencias'] if isinstance(s, str) and len(s) > 5][:5]}
                cache[cache_key] = result
                save_unified_ai_cache(cache)
                return result
        except Exception as exc:
            print(f"INFO: Error parseando sugerencias IA: {exc}", file=sys.stderr)

    # Fallback con columnas reales si la IA falla
    fallback = []
    if columnas:
        fallback.append(f"Resalta los valores más altos de '{columnas[0]}'")
    if len(columnas) > 1:
        fallback.append(f"Crea una gráfica comparativa de '{columnas[1]}'")
    if kpis:
        fallback.append(f"Pon los KPIs principales en la portada: {', '.join(kpis[:2])}")
    fallback += ["Genera conclusiones ejecutivas con alertas en rojo", "Agrupa los datos por categoría y muestra tendencias"]
    return {"sugerencias": fallback[:5]}


def preparar_datos_para_slides(excel_path, user_instructions=None):
    try:
        sheets = pd.read_excel(excel_path, sheet_name=None)
    except Exception as e:
        return {"error": str(e)}

    resultado = {}
    request_context = parse_user_request_context(user_instructions)
    resultado['user_instructions'] = request_context.get('prompt')
    resultado['presentation_request'] = request_context
    resultado['ai_curation_status'] = build_default_ai_curation_status()
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
    resultado['bloques_textuales'] = []

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
    
    if target_sheet:
        resultado['metadatos']['hoja_principal'] = target_sheet
        processed_sheets.add(target_sheet)
        df = sheets[target_sheet]
        df = remover_filas_basura(df)
        df = limpiar_df(df)
        df.attrs['sheet_name'] = target_sheet
        if should_prioritize_primary_story(target_sheet, df):
            workbook_profile = dict(workbook_profile)
            workbook_profile['tipo_libro'] = 'general'
            workbook_profile['conclusiones'] = []
            workbook_profile['insights'] = []
            resultado['perfil_libro'] = workbook_profile
            resultado['metadatos']['tipo_libro'] = 'general'
        
        cols_norm = {normalize_semantic_text(c): c for c in df.columns}
        
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
        
        # ── RESUMEN UNIVERSAL ────────────────────────────────────────
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
        
        # Gráficas automáticas (fallback)
        graficas_auto = generar_graficas_automaticas(df, cols_info, tendencia_result)
        if graficas_auto:
            resultado['graficas_automaticas'] = graficas_auto
        
        # === TABLA PRINCIPAL ===
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
        
        # === GRÁFICAS ESPECIALIZADAS ELIMINADAS ===
        # La IA o el algoritmo genérico deciden dinámicamente las gráficas.
        
        # === CONCLUSIONES INTELIGENTES ===
        conclusiones = generar_conclusiones(
            df, cols_info, kpis_auto, False,
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

    bloques_textuales = []
    for name, df in sheets.items():
        sheet_family = workbook_profile.get('familias_por_hoja', {}).get(name, 'general')
        if sheet_family not in TEXTUAL_FAMILIES:
            continue
        df = remover_filas_basura(df)
        df = limpiar_df(df)
        if df.empty:
            continue
        block = build_textual_block_from_dataframe(df, name, sheet_family=sheet_family)
        if block:
            bloques_textuales.append(block)

    bloques_textuales = bloques_textuales[:MAX_TEXTUAL_BLOCKS]
    if bloques_textuales:
        resultado['bloques_textuales'] = enrich_textual_blocks_with_ai(bloques_textuales)
        resultado['metadatos']['bloques_textuales'] = len(resultado['bloques_textuales'])
    else:
        resultado['metadatos']['bloques_textuales'] = 0

    # === PRESUPUESTO DE SLIDES ===
    presupuesto = calcular_presupuesto_slides(resultado, False)
    resultado['presupuesto_slides'] = presupuesto
    attach_visual_ids(resultado)
    if not resultado.get('conclusiones') and workbook_profile.get('conclusiones'):
        resultado['conclusiones'] = workbook_profile['conclusiones'][:MAX_CONCLUSIONES]
    if resultado.get('resumen_generico'):
        resultado['resumen_generico']['tipo_libro'] = workbook_profile.get('tipo_libro', 'general')
        resultado['resumen_generico']['familias_detectadas'] = workbook_profile.get('familias_detectadas', [])
    resultado['bloques_datos'] = build_data_block_inventory(resultado)

    # === IA UNIFICADA: 1 sola llamada cubre resumen + briefing + bloques ===
    ia_todo = sintetizar_todo_con_ia(resultado)
    ai_status = dict(resultado.get('ai_curation_status') or build_default_ai_curation_status())

    if ia_todo:
        ai_status['unified_call_succeeded'] = True
        ai_status['provider_used'] = True
        # Resumen ejecutivo
        ia_resumen = ia_todo.get('resumen_ejecutivo_ia')
        if isinstance(ia_resumen, dict):
            resultado['resumen_ejecutivo_ia'] = ia_resumen
            graficas = resultado.get('graficas_automaticas', [])
            for item in ia_resumen.get('insights_graficas', []):
                if not isinstance(item, dict): continue
                idx = item.get('id')
                if isinstance(idx, int) and 0 <= idx < len(graficas):
                    graficas[idx]['insight_auto'] = item.get('insight', graficas[idx].get('insight_auto'))

        # Briefing diapositivas 2-3
        ia_briefing = ia_todo.get('briefing_ejecutivo_ia')
        if ia_briefing:
            resultado['briefing_ejecutivo_ia'] = ia_briefing
            ai_status['briefing_received'] = True

        visual_plan = ia_todo.get('visual_plan_ia')
        if isinstance(visual_plan, dict):
            resultado['visual_plan_ia'] = visual_plan
            ai_status['visual_plan_received'] = True
            ai_status['visual_curation_ready'] = True
            ai_status['selected_chart_ids'] = [
                item.get('id') for item in (visual_plan.get('charts') or [])
                if isinstance(item, dict) and item.get('id')
            ]
            ai_status['selected_table_ids'] = [
                item.get('id') for item in (visual_plan.get('tables') or [])
                if isinstance(item, dict) and item.get('id') and item.get('modo') != 'omit'
            ]
            ai_status['reason'] = 'visual_plan_received'

            chart_lookup = {
                chart.get('_visual_ai_id'): chart
                for chart in (resultado.get('graficas_automaticas') or [])
                if isinstance(chart, dict) and chart.get('_visual_ai_id')
            }
            for chart_item in visual_plan.get('charts') or []:
                if not isinstance(chart_item, dict):
                    continue
                chart = chart_lookup.get(chart_item.get('id'))
                if not chart:
                    continue
                if chart_item.get('mensaje_clave'):
                    chart['insight_auto'] = sanitize_executive_text(chart_item.get('mensaje_clave'), max_len=180)
                chart['_visual_ai_selected'] = True
                chart['_visual_ai_rationale'] = sanitize_executive_text(chart_item.get('por_que_importa'), max_len=180)

            table_candidates = []
            if isinstance(resultado.get('muestra_tabla'), dict):
                table_candidates.append(resultado['muestra_tabla'])
            table_candidates.extend(
                table for table in (resultado.get('otras_tablas') or {}).values()
                if isinstance(table, dict)
            )
            table_candidates.extend(
                table for table in (resultado.get('genericas') or {}).values()
                if isinstance(table, dict)
            )
            table_lookup = {
                table.get('_visual_ai_id'): table
                for table in table_candidates
                if table.get('_visual_ai_id')
            }
            for table_item in visual_plan.get('tables') or []:
                if not isinstance(table_item, dict):
                    continue
                table = table_lookup.get(table_item.get('id'))
                if not table:
                    continue
                table['_visual_ai_mode'] = table_item.get('modo')
                table['_visual_ai_selected'] = table_item.get('modo') != 'omit'
                table['_visual_ai_message'] = sanitize_executive_text(table_item.get('mensaje_clave'), max_len=180)
                table['_visual_ai_rationale'] = sanitize_executive_text(table_item.get('por_que_importa'), max_len=180)

        # Bloques textuales enriquecidos
        bloques_enriquecidos = ia_todo.get('bloques_textuales_enriquecidos')
        if isinstance(bloques_enriquecidos, list) and bloques_enriquecidos:
            bloques_actuales = resultado.get('bloques_textuales') or []
            for i, bloque_ia in enumerate(bloques_enriquecidos[:len(bloques_actuales)]):
                if i < len(bloques_actuales) and isinstance(bloque_ia, dict):
                    b = dict(bloques_actuales[i])
                    if bloque_ia.get('title'):
                        b['title'] = sanitize_executive_text(bloque_ia['title'], max_len=90)
                    if bloque_ia.get('subtitle'):
                        b['subtitle'] = sanitize_executive_text(bloque_ia['subtitle'], max_len=90)
                    if bloque_ia.get('lines'):
                        b['lines'] = [sanitize_executive_text(l, max_len=180) for l in bloque_ia['lines'] if isinstance(l, str)][:MAX_TEXT_LINES_PER_BLOCK]
                    b['source_mode'] = 'ai'
                    bloques_actuales[i] = b
            resultado['bloques_textuales'] = bloques_actuales
        if not ai_status.get('visual_plan_received'):
            ai_status['reason'] = 'unified_ai_without_visual_plan'
            local_bundle = build_local_ai_curation_bundle(resultado)
            if local_bundle and isinstance(local_bundle.get('visual_plan_ia'), dict):
                visual_plan = local_bundle['visual_plan_ia']
                resultado['visual_plan_ia'] = visual_plan
                ai_status['visual_plan_received'] = True
                ai_status['visual_curation_ready'] = True
                ai_status['local_fallback_used'] = True
                ai_status['selected_chart_ids'] = [
                    item.get('id') for item in (visual_plan.get('charts') or [])
                    if isinstance(item, dict) and item.get('id')
                ]
                ai_status['selected_table_ids'] = [
                    item.get('id') for item in (visual_plan.get('tables') or [])
                    if isinstance(item, dict) and item.get('id') and item.get('modo') != 'omit'
                ]
                ai_status['reason'] = 'local_visual_plan_generated'

                chart_lookup = {
                    chart.get('_visual_ai_id'): chart
                    for chart in (resultado.get('graficas_automaticas') or [])
                    if isinstance(chart, dict) and chart.get('_visual_ai_id')
                }
                for chart_item in visual_plan.get('charts') or []:
                    if not isinstance(chart_item, dict):
                        continue
                    chart = chart_lookup.get(chart_item.get('id'))
                    if not chart:
                        continue
                    if chart_item.get('mensaje_clave'):
                        chart['insight_auto'] = sanitize_executive_text(chart_item.get('mensaje_clave'), max_len=180)
                    chart['_visual_ai_selected'] = True
                    chart['_visual_ai_rationale'] = sanitize_executive_text(chart_item.get('por_que_importa'), max_len=180)

                table_candidates = []
                if isinstance(resultado.get('muestra_tabla'), dict):
                    table_candidates.append(resultado['muestra_tabla'])
                table_candidates.extend(
                    table for table in (resultado.get('otras_tablas') or {}).values()
                    if isinstance(table, dict)
                )
                table_candidates.extend(
                    table for table in (resultado.get('genericas') or {}).values()
                    if isinstance(table, dict)
                )
                table_lookup = {
                    table.get('_visual_ai_id'): table
                    for table in table_candidates
                    if table.get('_visual_ai_id')
                }
                for table_item in visual_plan.get('tables') or []:
                    if not isinstance(table_item, dict):
                        continue
                    table = table_lookup.get(table_item.get('id'))
                    if not table:
                        continue
                    table['_visual_ai_mode'] = table_item.get('modo')
                    table['_visual_ai_selected'] = table_item.get('modo') != 'omit'
                    table['_visual_ai_message'] = sanitize_executive_text(table_item.get('mensaje_clave'), max_len=180)
                    table['_visual_ai_rationale'] = sanitize_executive_text(table_item.get('por_que_importa'), max_len=180)
    else:
        ai_status['reason'] = 'unified_ai_unavailable'
        local_bundle = build_local_ai_curation_bundle(resultado)
        if local_bundle:
            ai_status['local_fallback_used'] = True

            ia_resumen = local_bundle.get('resumen_ejecutivo_ia')
            if isinstance(ia_resumen, dict):
                resultado['resumen_ejecutivo_ia'] = ia_resumen
                graficas = resultado.get('graficas_automaticas', [])
                for item in ia_resumen.get('insights_graficas', []):
                    if not isinstance(item, dict):
                        continue
                    idx = item.get('id')
                    if isinstance(idx, int) and 0 <= idx < len(graficas):
                        graficas[idx]['insight_auto'] = item.get('insight', graficas[idx].get('insight_auto'))

            briefing_ia = local_bundle.get('briefing_ejecutivo_ia')
            if isinstance(briefing_ia, dict):
                resultado['briefing_ejecutivo_ia'] = briefing_ia
                ai_status['briefing_received'] = True

            visual_plan = local_bundle.get('visual_plan_ia')
            if isinstance(visual_plan, dict):
                resultado['visual_plan_ia'] = visual_plan
                ai_status['visual_plan_received'] = True
                ai_status['visual_curation_ready'] = True
                ai_status['selected_chart_ids'] = [
                    item.get('id') for item in (visual_plan.get('charts') or [])
                    if isinstance(item, dict) and item.get('id')
                ]
                ai_status['selected_table_ids'] = [
                    item.get('id') for item in (visual_plan.get('tables') or [])
                    if isinstance(item, dict) and item.get('id') and item.get('modo') != 'omit'
                ]
                ai_status['reason'] = 'local_visual_plan_generated'

                chart_lookup = {
                    chart.get('_visual_ai_id'): chart
                    for chart in (resultado.get('graficas_automaticas') or [])
                    if isinstance(chart, dict) and chart.get('_visual_ai_id')
                }
                for chart_item in visual_plan.get('charts') or []:
                    if not isinstance(chart_item, dict):
                        continue
                    chart = chart_lookup.get(chart_item.get('id'))
                    if not chart:
                        continue
                    if chart_item.get('mensaje_clave'):
                        chart['insight_auto'] = sanitize_executive_text(chart_item.get('mensaje_clave'), max_len=180)
                    chart['_visual_ai_selected'] = True
                    chart['_visual_ai_rationale'] = sanitize_executive_text(chart_item.get('por_que_importa'), max_len=180)

                table_candidates = []
                if isinstance(resultado.get('muestra_tabla'), dict):
                    table_candidates.append(resultado['muestra_tabla'])
                table_candidates.extend(
                    table for table in (resultado.get('otras_tablas') or {}).values()
                    if isinstance(table, dict)
                )
                table_candidates.extend(
                    table for table in (resultado.get('genericas') or {}).values()
                    if isinstance(table, dict)
                )
                table_lookup = {
                    table.get('_visual_ai_id'): table
                    for table in table_candidates
                    if table.get('_visual_ai_id')
                }
                for table_item in visual_plan.get('tables') or []:
                    if not isinstance(table_item, dict):
                        continue
                    table = table_lookup.get(table_item.get('id'))
                    if not table:
                        continue
                    table['_visual_ai_mode'] = table_item.get('modo')
                    table['_visual_ai_selected'] = table_item.get('modo') != 'omit'
                    table['_visual_ai_message'] = sanitize_executive_text(table_item.get('mensaje_clave'), max_len=180)
                    table['_visual_ai_rationale'] = sanitize_executive_text(table_item.get('por_que_importa'), max_len=180)

    resultado['ai_curation_status'] = ai_status

    return resultado


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "No file path provided"}))
        sys.exit(1)
    
    import warnings
    warnings.filterwarnings('ignore')

    if sys.argv[1] == "--suggestions":
        if len(sys.argv) < 3:
            print(json.dumps({"error": "No file path provided for suggestions"}))
            sys.exit(1)
        path_excel = sys.argv[2]
        try:
            data = preparar_datos_para_slides(path_excel)
            sugerencias = generar_sugerencias_ia(data)
            print(json.dumps(sugerencias, ensure_ascii=False))
        except Exception as e:
            print(json.dumps({"error": str(e)}, ensure_ascii=False))
        sys.exit(0)

    if sys.argv[1] == "--panel-report":
        if len(sys.argv) < 3:
            print(json.dumps({"error": "No file path provided for panel report"}))
            sys.exit(1)
        path_excel = sys.argv[2]
        try:
            user_instructions = sys.argv[3] if len(sys.argv) > 3 else None
            data = preparar_datos_para_slides(path_excel, user_instructions)
            sugerencias = generar_sugerencias_ia(data)
            report = {
                "analysis": data,
                "suggestions": sugerencias.get("sugerencias", []) if isinstance(sugerencias, dict) else [],
                "model": OPENROUTER_MODEL_PRIORITY[0],
            }
            print(json.dumps(report, ensure_ascii=False, default=str))
        except Exception as e:
            print(json.dumps({"error": str(e)}, ensure_ascii=False, default=str))
        sys.exit(0)

    path_excel = sys.argv[1]
    if not os.path.exists(path_excel):
        print(json.dumps({"error": f"File not found: {path_excel}"}))
        sys.exit(1)
        
    try:
        user_instructions = None
        if len(sys.argv) > 2:
            user_instructions = sys.argv[2]
            
        data = preparar_datos_para_slides(path_excel, user_instructions)
        print(json.dumps(data, ensure_ascii=False, default=str))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False, default=str))
        sys.exit(1)
