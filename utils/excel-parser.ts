// @ts-nocheck
/**
 * Universal Excel → PPTX Intelligence Engine v4.0
 * Works with ANY Excel file — auto-detects structure, types, and best visualization
 */
'use strict';
import * as XLSX from 'xlsx';

// ══════════════════════════════════════════════════════════════════════════════
// §1  6 ROTATING THEMES
// ══════════════════════════════════════════════════════════════════════════════

export const THEMES = [
  {
    id: 'dark_professional', name: 'Dark Professional', isDark: true,
    coverBg: '0D1117', slideBg: '161B22', cardBg: '21262D',
    primary: '00D4FF', secondary: 'F59E0B', accent: '3B82F6',
    text: 'F0F6FF', textMuted: '8B949E',
    tableHeaderBg: '1E40AF', tableHeaderText: 'F0F6FF',
    tableRow1: '0D1117', tableRow2: '161B22', tableText: 'D4E6F1',
    footerBg: '0A0E14', footerText: '8B949E', kpiCardBg: '21262D',
    divider: '30363D', sidebarW: 0.22,
    chartColors: ['00D4FF', 'F59E0B', '10B981', 'F87171', 'A78BFA', '34D399', 'FB923C', '818CF8'],
  },
  {
    id: 'minimal_white', name: 'Minimalista Blanco', isDark: false,
    coverBg: 'FFFFFF', slideBg: 'FAFAFA', cardBg: 'F3F4F6',
    primary: '1E40AF', secondary: '6366F1', accent: 'F59E0B',
    text: '111827', textMuted: '6B7280',
    tableHeaderBg: '1E3A8A', tableHeaderText: 'FFFFFF',
    tableRow1: 'F9FAFB', tableRow2: 'FFFFFF', tableText: '111827',
    footerBg: '1E3A8A', footerText: 'FFFFFF', kpiCardBg: 'FFFFFF',
    divider: 'E5E7EB', sidebarW: 0,
    chartColors: ['1E40AF', 'F59E0B', '10B981', 'EF4444', '8B5CF6', 'F97316', '06B6D4', 'EC4899'],
  },
  {
    id: 'gradient_modern', name: 'Gradiente Moderno', isDark: false,
    coverBg: '4F46E5', slideBg: 'EEF2FF', cardBg: 'FFFFFF',
    primary: '4F46E5', secondary: '7C3AED', accent: 'EC4899',
    text: '1E1B4B', textMuted: '6366F1',
    tableHeaderBg: '4F46E5', tableHeaderText: 'FFFFFF',
    tableRow1: 'F5F3FF', tableRow2: 'FFFFFF', tableText: '1E1B4B',
    footerBg: '4F46E5', footerText: 'FFFFFF', kpiCardBg: 'FFFFFF',
    divider: 'C7D2FE', sidebarW: 0,
    chartColors: ['4F46E5', '7C3AED', 'EC4899', '06B6D4', '10B981', 'F59E0B', 'F87171', '8B5CF6'],
  },
  {
    id: 'corporate_bold', name: 'Corporativo Audaz', isDark: true,
    coverBg: '001F5B', slideBg: '002171', cardBg: '0A2472',
    primary: 'FF6B35', secondary: 'FCD34D', accent: 'FFFFFF',
    text: 'FFFFFF', textMuted: 'BAC4D0',
    tableHeaderBg: 'FF6B35', tableHeaderText: 'FFFFFF',
    tableRow1: '001F5B', tableRow2: '0A2472', tableText: 'FFFFFF',
    footerBg: '000D33', footerText: 'BAC4D0', kpiCardBg: '0A2472',
    divider: '1A3A6B', sidebarW: 0.25,
    chartColors: ['FF6B35', 'FCD34D', '34D399', '60A5FA', 'F472B6', 'A78BFA', 'FB923C', '2DD4BF'],
  },
  {
    id: 'green_data', name: 'Verde Datos', isDark: true,
    coverBg: '022C22', slideBg: '022C22', cardBg: '065F46',
    primary: '10B981', secondary: '34D399', accent: 'FCD34D',
    text: 'ECFDF5', textMuted: '6EE7B7',
    tableHeaderBg: '10B981', tableHeaderText: '022C22',
    tableRow1: '022C22', tableRow2: '064E3B', tableText: 'ECFDF5',
    footerBg: '014737', footerText: '6EE7B7', kpiCardBg: '065F46',
    divider: '047857', sidebarW: 0.2,
    chartColors: ['10B981', '34D399', 'FCD34D', '60A5FA', 'F472B6', 'FB923C', '818CF8', '2DD4BF'],
  },
  {
    id: 'sunset_analytics', name: 'Sunset Analytics', isDark: true,
    coverBg: '2D1B69', slideBg: '2D1B69', cardBg: '4C1D95',
    primary: 'F97316', secondary: 'FB923C', accent: 'FCD34D',
    text: 'FEF3C7', textMuted: 'DDD6FE',
    tableHeaderBg: 'F97316', tableHeaderText: 'FFFFFF',
    tableRow1: '2D1B69', tableRow2: '3D1674', tableText: 'FEF3C7',
    footerBg: '1E0F4E', footerText: 'DDD6FE', kpiCardBg: '4C1D95',
    divider: '5B21B6', sidebarW: 0.22,
    chartColors: ['F97316', 'FCD34D', 'A78BFA', 'EC4899', '34D399', '60A5FA', 'F87171', 'FB923C'],
  },
];

// ══════════════════════════════════════════════════════════════════════════════
// §2  HELPERS
// ══════════════════════════════════════════════════════════════════════════════

const log = {
  info: (...a) => console.log('[PARSER]', ...a),
  warn: (...a) => console.warn('[PARSER WARN]', ...a),
  fix: (...a) => console.log('[AUTOFIX]', ...a),
};

// ── Randomization Pools (every presentation is unique) ──────────────────────

const TAGLINES = [
  'Análisis basado en datos reales',
  'Insights automatizados para la toma de decisiones',
  'Vista consolidada de indicadores clave',
  'Generado por el Motor de Inteligencia Socya',
  'Data-driven insights para ejecutivos',
  'Reporte de análisis multidimensional',
  'Panorama estratégico de datos operativos',
  'Tablero de control automatizado',
  'Diagnóstico integral de datos corporativos',
  'Transformando datos en decisiones',
  'Resumen ejecutivo de operaciones',
  'Análisis automatizado de indicadores',
  'Evaluación cuantitativa y cualitativa',
  'Informe de gestión basado en evidencia',
  'Métricas clave para la alta dirección',
  'Dashboard ejecutivo automatizado',
  'Síntesis analítica de datos operacionales',
  'Control y seguimiento de indicadores',
  'Visión 360° del desempeño organizacional',
  'Inteligencia de datos para líderes',
  'Análisis profundo de rendimiento',
  'Reporte estratégico de gestión',
];

const CLOSING_MSGS = [
  'Información generada automáticamente a partir de datos reales.',
  'Este reporte fue creado por el Motor de Inteligencia Socya.',
  'Análisis automatizado — datos verificados y consolidados.',
  'Generación inteligente de reportes corporativos.',
  'Reporte producido por Socya PPTX Generator v4.0.',
  'Datos procesados y visualizados automáticamente.',
  'Motor de Análisis Socya — Precisión & Automatización.',
  'Información procesada mediante algoritmos de detección automática.',
  'Socya — Excelencia en automatización de reportes.',
  'Análisis completado con éxito. Datos íntegros preservados.',
  'Automatización inteligente al servicio de la gestión.',
  'Reporte consolidado con validación WCAG integrada.',
];

const CLOSING_TITLES = ['¡Gracias!', 'Fin del Reporte', 'Resumen Completo', 'Gracias por su Atención', '¡Gracias por su Tiempo!'];

function pickRandom(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

function toNum(val) {
  if (val === null || val === undefined || val === '') return null;
  const s = String(val).trim().replace(/[$€£¥,\s]/g, '').replace(/\((\d+\.?\d*)\)/, '$1').replace(/COP|USD|EUR/ig, '');
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function isExcelDateSerial(v) { return typeof v === 'number' && v > 1 && v < 100000; }

function fmtDate(val) {
  if (!val) return '—';
  try {
    let d = isExcelDateSerial(val) ? new Date((val - 25569) * 86400000) : new Date(String(val));
    return isNaN(d.getTime()) ? String(val) : d.toLocaleDateString('es-CO', { day: '2-digit', month: 'short', year: 'numeric' });
  } catch { return String(val); }
}

function fmtCurrency(val) {
  const n = toNum(val);
  if (n === null) return '—';
  return new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 }).format(n);
}

function fmtPercent(val) {
  const n = toNum(val);
  if (n === null) return '—';
  return `${(n <= 1 ? n * 100 : n).toFixed(1)}%`;
}

function fmtNumber(val) {
  const n = toNum(val);
  if (n === null) return '—';
  return n.toLocaleString('es-CO', { maximumFractionDigits: 2 });
}

function truncate(str, max = 45) {
  const s = String(str ?? '').trim();
  return s.length > max ? s.slice(0, max - 1) + '…' : (s || '—');
}

function normKey(s) {
  return String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase().trim();
}

const BADGE_MAP = {
  'CONTABILIZADO': { bg: 'D1FAE5', text: '065F46', icon: '✓' },
  'LEGALIZADO': { bg: 'DBEAFE', text: '1E40AF', icon: '✓' },
  'RECHAZADO': { bg: 'FEE2E2', text: '991B1B', icon: '✗' },
  'SOLICITADO': { bg: 'E0F2FE', text: '0C4A6E', icon: '◷' },
  'APROBADO': { bg: 'D1FAE5', text: '065F46', icon: '✓' },
  'PENDIENTE': { bg: 'FEF3C7', text: '92400E', icon: '◷' },
  'COMPLETADO': { bg: 'D1FAE5', text: '065F46', icon: '✓' },
  'ENPROCESO': { bg: 'FEF3C7', text: '92400E', icon: '◷' },
  'EFECTIVO': { bg: 'D1FAE5', text: '065F46', icon: '✓' },
  'OPORTUNIDADDEMEJORA': { bg: 'FEF3C7', text: '92400E', icon: '!' },
  'SI': { bg: 'D1FAE5', text: '065F46', icon: '✓' },
  'NO': { bg: 'FEE2E2', text: '991B1B', icon: '✗' },
  'ALTO': { bg: 'FEE2E2', text: '991B1B', icon: '▲' },
  'MEDIO': { bg: 'FEF3C7', text: '92400E', icon: '▶' },
  'BAJO': { bg: 'D1FAE5', text: '065F46', icon: '▼' },
  'ACTIVO': { bg: 'D1FAE5', text: '065F46', icon: '✓' },
  'INACTIVO': { bg: 'F3F4F6', text: '6B7280', icon: '○' },
  'CRITICO': { bg: 'FEE2E2', text: '991B1B', icon: '⚠' },
};

function getBadge(val) {
  if (!val) return null;
  const k = normKey(String(val));
  if (k.length < 2) return null;
  if (BADGE_MAP[k]) return { value: truncate(val, 30), ...BADGE_MAP[k] };
  for (const [mk, st] of Object.entries(BADGE_MAP)) {
    if (k.includes(mk) || mk.includes(k)) return { value: truncate(val, 30), ...st };
  }
  if (/^(SI|YES|OK|TRUE|APROB|EFECT|LEGALIZ|CONTABILIZ|COMPLET)/.test(k)) return { value: truncate(val, 30), ...BADGE_MAP['SI'] };
  if (/^(NO|FALSE|ERROR|RECHAZ|CANCEL|IMPROCEDENTE)/.test(k)) return { value: truncate(val, 30), ...BADGE_MAP['NO'] };
  return null;
}

function formatValue(val, colType) {
  if (val === null || val === undefined || val === '') return '—';
  switch (colType) {
    case 'currency': return fmtCurrency(val);
    case 'date': return fmtDate(val);
    case 'percent': return fmtPercent(val);
    case 'numeric': return fmtNumber(val);
    case 'categorical':
    case 'boolean': return getBadge(val) ?? truncate(val, 35);
    case 'id': return truncate(val, 20);
    default: return truncate(val, 45);
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// §3  COLUMN TYPE DETECTION
// ══════════════════════════════════════════════════════════════════════════════

function detectColumnType(colName, values) {
  const sample = values.filter(v => v !== null && v !== undefined && String(v).trim() !== '').slice(0, 40);
  if (!sample.length) return 'text';
  const nk = normKey(colName);

  // Name-based fast paths
  if (/ID$|^ID|COD|CODIGO|FOLIO|NUMERO|CONS/.test(nk)) return 'id';
  if (/FECHA|DATE|DIA|MES|AÑO|YEAR|MONTH|PERIODO/.test(nk)) return 'date';
  if (/PORC|PCT|PERCENT|TASA|RATE|EFICIENCIA/.test(nk) || sample.some(v => typeof v === 'number' && v > 0 && v <= 1)) return 'percent';
  if (/VALOR|TOTAL|COSTO|PRECIO|MONTO|SALAR|GASTO|INGRES|COMIS|EGRESO|PAGO|DESEMBOLSO|APOYO/.test(nk)) return 'currency';

  const nums = sample.map(v => toNum(v)).filter(n => n !== null);
  const numRatio = nums.length / sample.length;

  if (numRatio >= 0.75) {
    const avg = nums.reduce((a, b) => a + b, 0) / nums.length;
    const allSmall01 = nums.every(n => n >= 0 && n <= 1);
    if (allSmall01 && nums.length >= 5) return 'percent';
    if (avg > 10000) return 'currency';
    return 'numeric';
  }

  const unique = new Set(sample.map(v => normKey(String(v))));
  if (unique.size <= 2 && (unique.has('SI') || unique.has('NO') || unique.has('TRUE') || unique.has('FALSE'))) return 'boolean';
  if (unique.size <= 10 && sample.length >= 5) return 'categorical';
  return 'text';
}

// ══════════════════════════════════════════════════════════════════════════════
// §4  CHART TYPE SELECTION
// ══════════════════════════════════════════════════════════════════════════════

function selectChartType(labelType, distinctCount, hasTimeSeries, values) {
  if (hasTimeSeries || labelType === 'date') return 'line';
  const total = values.reduce((a, b) => a + b, 0);
  const allPos = values.every(v => v >= 0);
  if (distinctCount <= 7 && allPos && total > 0) return 'donut';
  if (distinctCount > 8) return 'bar_horizontal';
  return 'bar';
}

// ══════════════════════════════════════════════════════════════════════════════
// §5  COLUMN WIDTH CALCULATOR (prevents table overflow)
// ══════════════════════════════════════════════════════════════════════════════

const TYPE_WIDTH_INCH = {
  id: 0.75, date: 1.1, currency: 1.35, percent: 0.9,
  numeric: 1.1, boolean: 1.1, categorical: 1.4, text: 0,
};
const SLIDE_TABLE_W = 12.3;
const MAX_COLS_IN_SLIDE = 8;

function calcColWidths(colDefs) {
  // Limit columns
  const cols = colDefs.slice(0, MAX_COLS_IN_SLIDE);
  const textCols = cols.filter(c => c.type === 'text');
  const nonTextW = cols.filter(c => c.type !== 'text').reduce((s, c) => s + TYPE_WIDTH_INCH[c.type], 0);
  const textW = Math.max(SLIDE_TABLE_W - nonTextW, textCols.length * 1.2);
  const perText = textCols.length ? textW / textCols.length : 0;

  const widths = cols.map(c => c.type === 'text' ? parseFloat(perText.toFixed(2)) : TYPE_WIDTH_INCH[c.type]);

  // Normalize to exactly SLIDE_TABLE_W
  const sumW = widths.reduce((a, b) => a + b, 0);
  if (sumW > 0) {
    const scale = SLIDE_TABLE_W / sumW;
    const scaled = widths.map(w => parseFloat((w * scale).toFixed(2)));
    const diff = parseFloat((SLIDE_TABLE_W - scaled.reduce((a, b) => a + b, 0)).toFixed(2));
    scaled[scaled.length - 1] += diff;
    return scaled;
  }
  return widths;
}

// ══════════════════════════════════════════════════════════════════════════════
// §6  DATA AGGREGATION FOR CHARTS
// ══════════════════════════════════════════════════════════════════════════════

function groupByForChart(rows, labelKey, valueKey, topN, chartColors) {
  const map = {};
  for (const r of rows) {
    const lv = r[labelKey];
    const nv = toNum(r[valueKey]);
    if (lv === null || lv === undefined) continue;
    const k = truncate(lv, 38);
    map[k] = (map[k] ?? 0) + (nv ?? 1);
  }
  let entries = Object.entries(map).sort((a, b) => b[1] - a[1]);
  if (entries.length > topN) {
    const top = entries.slice(0, topN - 1);
    const othersV = entries.slice(topN - 1).reduce((s, [, v]) => s + v, 0);
    if (othersV > 0) top.push(['Otros', othersV]);
    entries = top;
  }
  return entries.map(([label, value], i) => ({
    label: String(label),
    value: Number(value) || 0,
    color: chartColors[i % chartColors.length],
  }));
}

function countByForChart(rows, labelKey, chartColors) {
  return groupByForChart(rows, labelKey, null, 12, chartColors);
}

// ══════════════════════════════════════════════════════════════════════════════
// §6b HEADLINE INFERENCE
// ══════════════════════════════════════════════════════════════════════════════

function generateInferredHeadline(sheetName, colDefs) {
  // Build an intelligent headline from sheet name + column types
  const name = (sheetName || '').trim();
  const hasCurrency = colDefs.some(c => c.type === 'currency');
  const hasPercent = colDefs.some(c => c.type === 'percent');
  const hasCat = colDefs.some(c => c.type === 'categorical' || c.type === 'boolean');
  const hasDate = colDefs.some(c => c.type === 'date');
  const hasNumeric = colDefs.some(c => c.type === 'numeric');

  // If the sheet already has a descriptive name (>3 chars, not just "Sheet1"), use it directly
  if (name.length > 3 && !/^(hoja|sheet|data|datos)\s*\d*$/i.test(name)) {
    return name;
  }

  // Infer from column types
  const parts = [];
  if (hasCurrency) parts.push('Financiero');
  else if (hasPercent) parts.push('Indicadores');
  else if (hasNumeric) parts.push('Datos Cuantitativos');

  if (hasCat) parts.push('por Categoría');
  if (hasDate) parts.push('Evolutivo');

  if (parts.length > 0) {
    return `Análisis ${parts.join(' ')}`;
  }

  // Fallback: use visible column names
  const colNames = colDefs.slice(0, 3).map(c => c.header || c.key).join(', ');
  return colNames ? `Datos: ${colNames}` : `Análisis de ${name || 'Datos'}`;
}

// ══════════════════════════════════════════════════════════════════════════════
// §6c AUTO-INSIGHTS FOR CHART DATA
// ══════════════════════════════════════════════════════════════════════════════

function generateAutoInsights(data, labelHeader, valueHeader, valueType) {
  if (!data || data.length < 2) return [];

  const insights = [];
  const values = data.map(d => d.value);
  const total = values.reduce((a, b) => a + b, 0);
  const max = Math.max(...values);
  const min = Math.min(...values);
  const maxItem = data.find(d => d.value === max);
  const minItem = data.find(d => d.value === min);
  const fmtVal = (v) => valueType === 'currency' ? fmtCurrency(v) : fmtNumber(v);

  // Top item
  if (maxItem && total > 0) {
    const pct = ((max / total) * 100).toFixed(1);
    insights.push(`"${maxItem.label}" lidera con ${fmtVal(max)} (${pct}% del total).`);
  }

  // Bottom item
  if (minItem && minItem !== maxItem && total > 0) {
    const pct = ((min / total) * 100).toFixed(1);
    insights.push(`"${minItem.label}" registra el menor valor: ${fmtVal(min)} (${pct}%).`);
  }

  // Concentration check
  if (maxItem && total > 0 && (max / total) > 0.5) {
    insights.push(`Alta concentración: "${maxItem.label}" representa más de la mitad del total.`);
  }

  // Spread
  if (max > 0 && min >= 0 && data.length >= 3) {
    const ratio = max / (min || 1);
    if (ratio > 10) {
      insights.push(`Amplia dispersión: el valor máximo es ${ratio.toFixed(0)}× el mínimo.`);
    }
  }

  return insights.slice(0, 3);
}

// ══════════════════════════════════════════════════════════════════════════════
// §7  KPI EXTRACTION
// ══════════════════════════════════════════════════════════════════════════════

function extractKPIs(rows, colDefs) {
  const kpis = [];
  const numCols = colDefs.filter(c => ['currency', 'numeric', 'percent'].includes(c.type));
  for (const col of numCols.slice(0, 4)) {
    const vals = rows.map(r => toNum(r[col.key])).filter(n => n !== null);
    if (!vals.length) continue;
    const sum = vals.reduce((a, b) => a + b, 0);
    const label = col.header || col.key;
    const value = col.type === 'currency' ? fmtCurrency(sum)
      : col.type === 'percent' ? fmtPercent(sum / vals.length)
        : fmtNumber(sum);
    kpis.push({ label: truncate(label, 25), value });
  }
  // Always add count
  kpis.unshift({ label: 'Total Registros', value: rows.length.toLocaleString('es-CO') });
  return kpis.slice(0, 5);
}

// ══════════════════════════════════════════════════════════════════════════════
// §8  SHEET PARSING — Universal (any Excel)
// ══════════════════════════════════════════════════════════════════════════════

function detectHeaderRow(rawRows) {
  for (let i = 0; i < Math.min(rawRows.length, 15); i++) {
    const row = rawRows[i] ?? [];
    const nonEmpty = row.filter(c => c !== null && c !== undefined && String(c).trim() !== '');
    if (nonEmpty.length >= 2) return i;
  }
  return 0;
}

function extractTitle(rawRows, headerIdx) {
  let title = '', subtitle = '';
  for (let i = 0; i < headerIdx && i < 8; i++) {
    const row = rawRows[i] ?? [];
    const cell = String(row[0] ?? '').trim();
    if (cell.toUpperCase().startsWith('TITLE:')) { title = cell.substring(6).trim(); continue; }
    if (cell.toUpperCase().startsWith('SUBTITLE:')) { subtitle = cell.substring(9).trim(); continue; }
    const nonempty = row.filter(c => c !== null && String(c).trim());
    if (nonempty.length === 1 && cell.length > 3 && !title) title = cell;
    else if (nonempty.length === 1 && cell.length > 3 && !subtitle && title) subtitle = cell;
  }
  return { title, subtitle };
}

function parseSheetUniversal(wb, sheetName) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!rawRows.length) return null;

  // Filter completely empty rows
  const cleanRows = rawRows.filter(r => r && r.some(c => c !== null && c !== undefined && String(c).trim() !== ''));
  if (!cleanRows.length) return null;

  const headerIdx = detectHeaderRow(cleanRows);
  const { title, subtitle } = extractTitle(cleanRows, headerIdx);
  const rawHeaders = (cleanRows[headerIdx] ?? []).map(h => String(h ?? '').trim());

  // Parse body rows into objects
  const bodyRaw = cleanRows.slice(headerIdx + 1).filter(r =>
    r && r.some(c => c !== null && String(c).trim() !== '')
  );

  // Build col defs (detect types from first 30 rows)
  const colDefs = rawHeaders.map((h, idx) => {
    if (!h || h === '') return null;
    const vals = bodyRaw.slice(0, 30).map(r => r[idx] ?? null);
    const type = detectColumnType(h, vals);
    return { key: h, header: h, idx, type };
  }).filter(Boolean);

  // Convert rows to objects
  const rows = bodyRaw.map(r => {
    const obj = {};
    colDefs.forEach(c => { obj[c.key] = r[c.idx] ?? null; });
    return obj;
  });

  return { rows, colDefs, rawHeaders, sheetName, meta: { title, subtitle } };
}

// ══════════════════════════════════════════════════════════════════════════════
// §9  SLIDE BUILDERS
// ══════════════════════════════════════════════════════════════════════════════

function buildTableSlides(sheetData, sheetName, theme, maxContentRows = 18) {
  const { rows, colDefs, meta } = sheetData;
  if (!rows.length || !colDefs.length) return [];

  const subtitle = meta.subtitle || '';

  // Select best columns (max 8)
  const hasBadgeCols = colDefs.filter(c => ['categorical', 'boolean'].includes(c.type));
  const hasNumCols = colDefs.filter(c => ['currency', 'numeric', 'percent'].includes(c.type));
  const hasTextCols = colDefs.filter(c => c.type === 'text');
  const hasIdCols = colDefs.filter(c => c.type === 'id');

  let selCols = [...hasIdCols.slice(0, 1), ...hasNumCols.slice(0, 3), ...hasBadgeCols.slice(0, 2), ...hasTextCols.slice(0, 2)];
  if (selCols.length < 3) selCols = colDefs.slice(0, 8);
  selCols = selCols.slice(0, MAX_COLS_IN_SLIDE);
  const colW = calcColWidths(selCols);

  const slides = [];
  const totalRows = rows.length;
  const numSlides = Math.ceil(totalRows / maxContentRows);

  for (let sIdx = 0; sIdx < numSlides; sIdx++) {
    const chunk = rows.slice(sIdx * maxContentRows, (sIdx + 1) * maxContentRows);
    const tableRows = chunk.map(r => {
      const rowObj = {};
      selCols.forEach(c => {
        const raw = r[c.key] ?? null;
        const fmtd = formatValue(raw, c.type);
        const badge = (['categorical', 'boolean'].includes(c.type)) ? getBadge(raw) : null;
        rowObj[c.key] = badge ?? (fmtd === '—' && raw !== null ? String(raw) : fmtd);
      });
      return rowObj;
    });

    const kpis = sIdx === 0 ? extractKPIs(rows, colDefs) : [];
    const overflow = totalRows - (sIdx + 1) * maxContentRows;
    const suffix = numSlides > 1 ? ` (${sIdx + 1}/${numSlides})` : '';

    slides.push({
      type: 'table',
      title: (meta.title || generateInferredHeadline(sheetName, colDefs)) + suffix,
      subtitle,
      kpis,
      columns: selCols,
      colW,
      rows: tableRows,
      overflow: Math.max(0, overflow),
      totalRows,
    });
  }
  return slides;
}

function buildChartSlide(sheetData, sheetName, theme) {
  const { rows, colDefs, meta } = sheetData;
  if (!rows.length || !colDefs.length) return null;

  const title = meta.title || generateInferredHeadline(sheetName, colDefs);
  const catCols = colDefs.filter(c => ['categorical', 'boolean', 'text', 'id'].includes(c.type));
  const numCols = colDefs.filter(c => ['currency', 'numeric', 'percent'].includes(c.type));

  if (!catCols.length || !numCols.length) {
    // Fallback: count by first categorical
    const col = catCols[0] || colDefs[0];
    if (!col) return null;
    const data = countByForChart(rows, col.key, theme.chartColors);
    if (!data.length) return null;
    return {
      type: 'chart',
      title,
      subtitle: meta.subtitle || '',
      kpis: extractKPIs(rows, colDefs),
      chart: {
        type: data.length <= 7 ? 'donut' : 'bar',
        title: `Distribución por ${col.header}`,
        data,
      },
    };
  }

  const labelCol = catCols[0];
  const valueCol = numCols[0];
  const hasTimePattern = labelCol.type === 'date' || /mes|month|año|year|periodo/i.test(labelCol.key);
  const distinct = new Set(rows.map(r => r[labelCol.key])).size;
  const data = groupByForChart(rows, labelCol.key, valueCol.key, 12, theme.chartColors);
  const vals = data.map(d => d.value);
  const chartType = selectChartType(labelCol.type, distinct, hasTimePattern, vals);

  const insights = generateAutoInsights(data, labelCol.header, valueCol.header, valueCol.type);

  return {
    type: 'chart',
    title,
    subtitle: meta.subtitle || '',
    kpis: extractKPIs(rows, colDefs),
    chart: { type: chartType, title: `${valueCol.header} por ${labelCol.header}`, data },
    insights: insights,
  };
}

function buildMixedSlide(sheetData, sheetName, theme) {
  const { rows, colDefs, meta } = sheetData;
  if (!rows.length || !colDefs.length) return null;

  const catCols = colDefs.filter(c => ['categorical', 'boolean'].includes(c.type));
  const numCols = colDefs.filter(c => ['currency', 'numeric', 'percent'].includes(c.type));

  // Table: show max 4 columns to fit in 50% width
  const selCols = colDefs.slice(0, 4);
  const tableColW = selCols.map(() => parseFloat((6.0 / selCols.length).toFixed(2)));
  const tableRows = rows.slice(0, 15).map(r => {
    const obj = {};
    selCols.forEach(c => {
      const raw = r[c.key] ?? null;
      const badge = ['categorical', 'boolean'].includes(c.type) ? getBadge(raw) : null;
      obj[c.key] = badge ?? formatValue(raw, c.type);
    });
    return obj;
  });

  // Chart
  const labelCol = catCols[0] || colDefs.find(c => c.type !== 'numeric' && c.type !== 'currency') || colDefs[0];
  const valueCol = numCols[0] || colDefs.find(c => ['numeric', 'currency'].includes(c.type));
  let chartData = [];
  if (labelCol && valueCol && labelCol.key !== valueCol.key) {
    chartData = groupByForChart(rows, labelCol.key, valueCol.key, 8, theme.chartColors);
  } else if (catCols.length) {
    chartData = countByForChart(rows, catCols[0].key, theme.chartColors);
  }

  const chartType = chartData.length <= 6 ? 'donut' : 'bar';

  return {
    type: 'mixed',
    title: meta.title || generateInferredHeadline(sheetName, colDefs),
    subtitle: meta.subtitle || '',
    kpis: extractKPIs(rows, colDefs),
    columns: selCols,
    colW: tableColW,
    rows: tableRows,
    overflow: Math.max(0, rows.length - 15),
    chart: chartData.length ? { type: chartType, title: `Por ${(catCols[0] || labelCol)?.header}`, data: chartData } : null,
  };
}

function buildDualChartSlide(sheetData, sheetName, theme) {
  const { rows, colDefs, meta } = sheetData;
  const catCols = colDefs.filter(c => ['categorical', 'boolean'].includes(c.type));
  const numCols = colDefs.filter(c => ['currency', 'numeric', 'percent'].includes(c.type));
  if (numCols.length < 2 || !catCols.length || rows.length < 3) return null;

  const labelCol = catCols[0];
  const leftData = groupByForChart(rows, labelCol.key, numCols[0].key, 8, theme.chartColors);
  const rightColors = [...theme.chartColors].reverse();
  const rightData = groupByForChart(rows, labelCol.key, numCols[1].key, 8, rightColors);
  if (leftData.length < 2 || rightData.length < 2) return null;

  const leftType = leftData.length <= 6 ? 'bar' : 'bar_horizontal';
  const rightType = rightData.length <= 6 ? 'donut' : 'bar';

  return {
    type: 'dual_chart',
    title: `Comparativa: ${meta.title || sheetName}`,
    subtitle: `${numCols[0].header} vs ${numCols[1].header} por ${labelCol.header}`,
    kpis: extractKPIs(rows, colDefs),
    chartLeft: { type: leftType, title: numCols[0].header, data: leftData },
    chartRight: { type: rightType, title: numCols[1].header, data: rightData },
  };
}

// ══════════════════════════════════════════════════════════════════════════════
// §10  SLIDE LAYOUT DECISION
// ══════════════════════════════════════════════════════════════════════════════

function decideLayout(colDefs, rowCount) {
  const numericCount = colDefs.filter(c => ['currency', 'numeric', 'percent'].includes(c.type)).length;
  const catCount = colDefs.filter(c => ['categorical', 'boolean'].includes(c.type)).length;
  const totalCols = colDefs.length;

  // Chart: exactly 2 cols with one label + one value — pure aggregated
  if (totalCols === 2 && numericCount === 1) return 'chart';
  // Pure chart: small categorical+numeric dataset, ideal for visualization
  if (rowCount <= 15 && numericCount >= 1 && catCount >= 1 && totalCols === 2) return 'chart';
  // Mixed: moderate size with both categorical and numeric cols
  if (rowCount <= 30 && numericCount >= 1 && catCount >= 1 && totalCols >= 3 && totalCols <= 7) return 'mixed';
  // Mixed: few cols with numeric+cat, good for side-by-side
  if (numericCount >= 1 && catCount >= 1 && totalCols <= 4 && rowCount > 10) return 'mixed';
  // Table for everything else (many cols, text-heavy, or large datasets)
  return 'table';
}

// ══════════════════════════════════════════════════════════════════════════════
// §11  EXECUTIVE SUMMARY
// ══════════════════════════════════════════════════════════════════════════════

function buildExecutiveSummary(allSheetData, theme) {
  let totalRows = 0;
  let sheetCount = 0;
  const allBadges = {};
  const currencyTotals = {};

  for (const sd of allSheetData) {
    totalRows += sd.rows.length;
    sheetCount++;
    sd.rows.forEach(r => {
      // Collect badge/status values
      sd.colDefs.filter(c => c.type === 'categorical' || c.type === 'boolean').forEach(c => {
        const v = r[c.key];
        if (v !== null && v !== undefined) {
          const k = truncate(String(v), 28);
          allBadges[k] = (allBadges[k] ?? 0) + 1;
        }
      });
      // Collect currency totals across all sheets
      sd.colDefs.filter(c => c.type === 'currency').slice(0, 2).forEach(c => {
        const n = toNum(r[c.key]);
        if (n !== null) {
          currencyTotals[c.header] = (currencyTotals[c.header] ?? 0) + n;
        }
      });
    });
  }

  const statusDist = Object.entries(allBadges)
    .sort((a, b) => b[1] - a[1]).slice(0, 7)
    .map(([label, count], i) => ({ label, count, badge: getBadge(label), color: theme.chartColors[i] }));

  const globalKPIs = [
    { label: 'Total Registros', value: totalRows.toLocaleString('es-CO') },
    { label: 'Hojas Analizadas', value: sheetCount.toString() },
  ];

  // Add top currency KPIs found across all sheets (max 3)
  const topCurrencyEntries = Object.entries(currencyTotals)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3);
  for (const [label, sum] of topCurrencyEntries) {
    globalKPIs.push({ label: truncate(label, 22), value: fmtCurrency(sum) });
    if (globalKPIs.length >= 5) break;
  }

  // If no currency found, fall back to numeric from first sheet
  if (topCurrencyEntries.length === 0 && allSheetData.length > 0) {
    const first = allSheetData[0];
    const numCol = first.colDefs.find(c => c.type === 'numeric');
    if (numCol) {
      const vals = first.rows.map(r => toNum(r[numCol.key])).filter(n => n !== null);
      const sum = vals.reduce((a, b) => a + b, 0);
      globalKPIs.push({ label: truncate(numCol.header, 22), value: fmtNumber(sum) });
    }
  }

  return {
    type: 'executive_summary',
    title: 'Resumen Ejecutivo',
    subtitle: 'Vista consolidada del análisis',
    kpis: globalKPIs.slice(0, 5),
    statusDistribution: statusDist,
    sheetSummary: allSheetData.map(sd => ({
      name: sd.meta.title || sd.sheetName,
      rows: sd.rows.length,
      layout: decideLayout(sd.colDefs, sd.rows.length),
    })),
  };
}

// ══════════════════════════════════════════════════════════════════════════════
// §12  VALIDATION & AUTOCORRECTION
// ══════════════════════════════════════════════════════════════════════════════

function validateAndFix(slides) {
  const fixed = [];
  let fixCount = 0;

  for (const s of slides) {
    if (!s || !s.type) continue;

    // Fix 1: Empty title
    if (!s.title || !s.title.trim()) {
      s.title = `Sección ${fixed.length + 1}`;
      log.fix('Título vacío corregido →', s.title);
      fixCount++;
    }

    // Fix 2: Table with empty cells → 'N/D'
    if (s.rows && s.columns) {
      s.rows = s.rows.map(r => {
        const nr = { ...r };
        s.columns.forEach(c => {
          const v = nr[c.key];
          if (v === null || v === undefined || v === '' || v === '—' && r[c.key] !== '—') {
            nr[c.key] = 'N/D';
            fixCount++;
          }
        });
        return nr;
      });
    }

    // Fix 3: Chart with no data
    if (s.chart && (!s.chart.data || !s.chart.data.length)) {
      log.fix('Gráfico sin datos eliminado en slide:', s.title);
      s.chart = null;
      if (s.type === 'chart') { s.type = 'table'; }
      if (s.type === 'mixed') { s.type = 'table'; }
      fixCount++;
    }

    // Fix 4: Donut/Pie with >7 categories → group into 'Otros'
    if (s.chart && s.chart.type === 'donut' && s.chart.data && s.chart.data.length > 7) {
      const top6 = s.chart.data.slice(0, 6);
      const othersV = s.chart.data.slice(6).reduce((sum, d) => sum + d.value, 0);
      if (othersV > 0) top6.push({ label: 'Otros', value: othersV, color: '94A3B8' });
      s.chart.data = top6;
      log.fix('Donut con >7 categorías agrupadas en "Otros"');
      fixCount++;
    }

    // Fix 5: Chart values that include NaN → remove entry
    if (s.chart && s.chart.data) {
      const before = s.chart.data.length;
      s.chart.data = s.chart.data.filter(d => !isNaN(d.value) && d.value !== null);
      if (s.chart.data.length < before) { log.fix(`${before - s.chart.data.length} puntos NaN eliminados del gráfico`); fixCount++; }
      if (!s.chart.data.length) { s.chart = null; fixCount++; }
    }

    // Fix 6: colW sum must match SLIDE_TABLE_W
    if (s.colW && s.columns) {
      const sum = s.colW.reduce((a, b) => a + b, 0);
      if (Math.abs(sum - SLIDE_TABLE_W) > 0.05) {
        const scale = SLIDE_TABLE_W / sum;
        s.colW = s.colW.map(w => parseFloat((w * scale).toFixed(2)));
        const diff = parseFloat((SLIDE_TABLE_W - s.colW.reduce((a, b) => a + b, 0)).toFixed(2));
        s.colW[s.colW.length - 1] += diff;
        fixCount++;
      }
    }

    fixed.push(s);
  }

  if (fixCount > 0) log.info(`✅ Autocorrección: ${fixCount} problemas corregidos automáticamente`);
  return fixed;
}

// ══════════════════════════════════════════════════════════════════════════════
// §13  MAX 25 SLIDES ENFORCEMENT
// ══════════════════════════════════════════════════════════════════════════════

function enforceSlideLimit(slides, maxSlides = 25) {
  if (slides.length <= maxSlides) return slides;

  log.warn(`${slides.length} diapositivas exceeden el límite de ${maxSlides}. Consolidando...`);
  const cover = slides.find(s => s.type === 'cover');
  const index = slides.find(s => s.type === 'index');
  const exec = slides.find(s => s.type === 'executive_summary');
  const insights = slides.find(s => s.type === 'insights');
  const closing = slides.find(s => s.type === 'closing');

  // Filter out structural slides to get pure content
  const structural = ['cover', 'index', 'executive_summary', 'closing', 'insights'];
  const content = slides.filter(s => !structural.includes(s.type));

  // When over limit, first remove section dividers to save space
  const reservedSlots = [cover, index, exec, insights, closing].filter(Boolean).length;
  const availableForContent = maxSlides - reservedSlots - 1; // -1 for overflow

  let keepContent;
  if (content.length > availableForContent) {
    // Remove section dividers first to save space
    const withoutDividers = content.filter(s => s.type !== 'section_divider');
    keepContent = withoutDividers.slice(0, availableForContent);
  } else {
    keepContent = content.slice(0, availableForContent);
  }

  const dropped = content.filter(s => s.type !== 'section_divider').filter(s => !keepContent.includes(s));

  const result = [cover, index, exec, ...keepContent].filter(Boolean);

  // Add insights before overflow
  if (insights) result.push(insights);

  if (dropped.length > 0) {
    const totalDroppedRows = dropped.reduce((sum, s) => sum + (s.totalRows || s.rows?.length || 0), 0);
    result.push({
      type: 'overflow_detail',
      title: 'Información Complementaria',
      subtitle: `${dropped.length} secciones · ${totalDroppedRows.toLocaleString('es-CO')} registros adicionales`,
      sections: dropped.map(s => ({ name: s.title, rows: s.totalRows || (s.rows?.length ?? 0) })),
      message: 'Los datos completos están disponibles en el archivo Excel original adjunto. Navegue a la diapositiva 1 para volver al inicio.',
    });
  }

  if (closing) result.push(closing);
  return result;
}


// ══════════════════════════════════════════════════════════════════════════════
// §14  STAT CALLOUT BUILDER
// ══════════════════════════════════════════════════════════════════════════════

function buildStatCallout(sheetData, sheetName) {
  const { rows, colDefs, meta } = sheetData;
  const numCols = colDefs.filter(c => ['currency', 'numeric', 'percent'].includes(c.type));
  if (numCols.length < 3 || rows.length < 2) return null;

  const stats = numCols.slice(0, 4).map((col, i) => {
    const vals = rows.map(r => toNum(r[col.key])).filter(n => n !== null);
    if (!vals.length) return null;
    const sum = vals.reduce((a, b) => a + b, 0);
    const avg = sum / vals.length;
    const max = Math.max(...vals);

    const icons = ['📊', '💰', '📈', '🎯'];
    const isC = col.type === 'currency';

    return {
      icon: icons[i % icons.length],
      value: isC ? fmtCurrency(sum) : fmtNumber(sum),
      label: truncate(col.header, 25),
      description: `Promedio: ${isC ? fmtCurrency(avg) : fmtNumber(avg)} · Máx: ${isC ? fmtCurrency(max) : fmtNumber(max)}`,
    };
  }).filter(Boolean);

  if (stats.length < 2) return null;

  return {
    type: 'stat_callout',
    title: `Indicadores: ${meta.title || sheetName}`,
    subtitle: `${rows.length} registros analizados · ${stats.length} métricas clave`,
    stats,
  };
}

// ══════════════════════════════════════════════════════════════════════════════
// §15  AUTO-DETECT INSIGHTS/FINDINGS
// ══════════════════════════════════════════════════════════════════════════════

function buildAutoInsights(allSheetData) {
  const findings = [];

  for (const sd of allSheetData) {
    const { rows, colDefs, sheetName, meta } = sd;
    const title = meta.title || sheetName;

    // Finding 1: Missing data detection
    for (const col of colDefs.slice(0, 6)) {
      const total = rows.length;
      const empty = rows.filter(r => {
        const v = r[col.key];
        return v === null || v === undefined || String(v).trim() === '';
      }).length;
      const pct = (empty / total) * 100;
      if (pct > 20 && empty > 5) {
        findings.push({
          title: `Datos faltantes en "${truncate(col.header, 25)}"`,
          description: `${empty} de ${total} registros (${pct.toFixed(0)}%) vacíos en hoja "${title}". Considere validar la fuente de datos.`,
          severity: pct > 50 ? 'high' : 'medium',
          sheet: title,
        });
      }
    }

    // Finding 2: Outlier detection in currency/numeric columns
    for (const col of colDefs.filter(c => ['currency', 'numeric'].includes(c.type)).slice(0, 3)) {
      const vals = rows.map(r => toNum(r[col.key])).filter(n => n !== null);
      if (vals.length < 5) continue;
      const avg = vals.reduce((a, b) => a + b, 0) / vals.length;
      const max = Math.max(...vals);
      if (max > avg * 5 && avg > 0) {
        findings.push({
          title: `Valor atípico en "${truncate(col.header, 25)}"`,
          description: `Valor máximo ${col.type === 'currency' ? fmtCurrency(max) : fmtNumber(max)} excede ×${(max / avg).toFixed(0)} el promedio en "${title}".`,
          severity: max > avg * 10 ? 'high' : 'medium',
          sheet: title,
        });
      }
    }

    // Finding 3: High concentration in categorical columns
    for (const col of colDefs.filter(c => c.type === 'categorical').slice(0, 2)) {
      const counts = {};
      rows.forEach(r => {
        const v = String(r[col.key] ?? '').trim();
        if (v) counts[v] = (counts[v] || 0) + 1;
      });
      const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
      if (entries.length >= 2) {
        const topPct = (entries[0][1] / rows.length) * 100;
        if (topPct > 60) {
          findings.push({
            title: `Alta concentración: "${truncate(col.header, 20)}"`,
            description: `"${truncate(entries[0][0], 20)}" representa el ${topPct.toFixed(0)}% de los registros en "${title}".`,
            severity: 'low',
            sheet: title,
          });
        }
      }
    }
  }

  // Limit and deduplicate
  const unique = findings.slice(0, 6);
  if (!unique.length) return null;

  return {
    type: 'insights',
    title: 'Hallazgos Automáticos',
    subtitle: `${unique.length} observaciones detectadas en el análisis de datos`,
    findings: unique,
  };
}

// ══════════════════════════════════════════════════════════════════════════════
// §16  MAIN EXPORT
// ══════════════════════════════════════════════════════════════════════════════

export function parsePresentationFromWorkbook(wb, options = {}) {
  const selectedTheme = options.theme ?? THEMES[Math.floor(Math.random() * THEMES.length)];
  log.info(`Tema seleccionado: ${selectedTheme.name}`);

  const slides = [];
  const allSheetData = [];

  // Cover
  const fileName = options.sourceFile || 'Reporte';
  slides.push({
    type: 'cover',
    title: options.title || 'Informe de Análisis',
    subtitle: options.subtitle || pickRandom(TAGLINES),
    organization: options.organization || 'Socya',
    date: new Date().toLocaleDateString('es-CO', { year: 'numeric', month: 'long', day: 'numeric' }),
    fileName,
    themeId: selectedTheme.id,
  });

  // Parse all sheets
  for (const sn of wb.SheetNames) {
    const data = parseSheetUniversal(wb, sn);
    if (!data || !data.rows.length) {
      log.warn(`Hoja "${sn}" vacía o sin datos válidos. Omitida.`);
      continue;
    }
    log.info(`Hoja "${sn}": ${data.rows.length} filas, ${data.colDefs.length} columnas. Tipos: ${data.colDefs.map(c => c.type).join(', ')}`);
    allSheetData.push(data);
  }

  // Executive summary (after cover, before content)
  if (allSheetData.length > 0) {
    slides.push(buildExecutiveSummary(allSheetData, selectedTheme));
  }

  // Content slides for each sheet — with section dividers
  let sectionNum = 0;
  for (const sd of allSheetData) {
    if (!sd.rows.length || !sd.colDefs.length) continue;
    sectionNum++;

    const sheetTitle = sd.meta.title || sd.sheetName;
    const layout = decideLayout(sd.colDefs, sd.rows.length);
    log.info(`Hoja "${sd.sheetName}" → layout: ${layout}`);

    // Add section divider for each sheet group (only when multiple sheets or many rows)
    if (allSheetData.length > 1) {
      slides.push({
        type: 'section_divider',
        sectionNumber: sectionNum,
        title: sheetTitle,
        subtitle: `${sd.rows.length} registros · ${sd.colDefs.length} columnas analizadas`,
      });
    }

    let newSlides = [];

    // Build stat callout for data-heavy sheets (before tables)
    const numCols = sd.colDefs.filter(c => ['currency', 'numeric', 'percent'].includes(c.type));
    if (numCols.length >= 3 && sd.rows.length >= 5) {
      const statSlide = buildStatCallout(sd, sd.sheetName, selectedTheme);
      if (statSlide) newSlides.push(statSlide);
    }

    if (layout === 'table' || layout === 'kpi_cards') {
      newSlides.push(...buildTableSlides(sd, sd.sheetName, selectedTheme));
    } else if (layout === 'chart') {
      const s = buildChartSlide(sd, sd.sheetName, selectedTheme);
      if (s) newSlides.push(s);
    } else if (layout === 'mixed') {
      const s = buildMixedSlide(sd, sd.sheetName, selectedTheme);
      if (s) newSlides.push(s);
      else newSlides.push(...buildTableSlides(sd, sd.sheetName, selectedTheme));
    } else {
      newSlides.push(...buildTableSlides(sd, sd.sheetName, selectedTheme));
    }

    // Add chart companion for table slides that have numeric + categorical data
    if (layout === 'table' && newSlides.length > 0) {
      const catCols = sd.colDefs.filter(c => ['categorical', 'boolean'].includes(c.type));
      if (numCols.length && catCols.length && sd.rows.length >= 3) {
        let bestData = [];
        let bestCat = null;
        for (const cc of catCols.slice(0, 3)) {
          const d = groupByForChart(sd.rows, cc.key, numCols[0].key, 10, selectedTheme.chartColors);
          if (d.length >= 2 && d.length > bestData.length) { bestData = d; bestCat = cc; }
        }
        if (bestData.length >= 2 && bestCat) {
          const hasTime = bestCat.type === 'date' || /mes|month|año|year|periodo/i.test(bestCat.key);
          const ct = selectChartType(bestCat.type, bestData.length, hasTime, bestData.map(d => d.value));
          newSlides.push({
            type: 'chart',
            title: `Análisis: ${sheetTitle}`,
            subtitle: `${numCols[0].header} por ${bestCat.header}`,
            kpis: [],
            chart: { type: ct, title: `${numCols[0].header} por ${bestCat.header}`, data: bestData },
          });
        }
      }
    }

    // Add dual chart for sheets with 2+ numeric columns and categories
    if (numCols.length >= 2 && sd.rows.length >= 3) {
      const dualSlide = buildDualChartSlide(sd, sd.sheetName, selectedTheme);
      if (dualSlide) newSlides.push(dualSlide);
    }

    slides.push(...newSlides);
  }

  // Auto-insights slide (before closing)
  if (allSheetData.length > 0) {
    const insightsSlide = buildAutoInsights(allSheetData);
    if (insightsSlide) {
      slides.push(insightsSlide);
      log.info(`✅ ${insightsSlide.findings.length} hallazgos automáticos detectados`);
    }
  }

  // Closing with randomized message for unique presentations
  slides.push({
    type: 'closing',
    title: pickRandom(CLOSING_TITLES),
    message: `${pickRandom(CLOSING_MSGS)} · ${new Date().toLocaleDateString('es-CO')}`,
    organization: options.organization || 'Socya',
  });

  // Validate and fix
  const validated = validateAndFix(slides);

  // Enforce 24 slide max (reserve 1 slot for index added after)
  const final = enforceSlideLimit(validated, 24);

  // Build index AFTER enforcement so it only references surviving slides
  const indexItems = final
    .filter(s => !['cover', 'index', 'section_divider', 'closing', 'overflow_detail'].includes(s.type))
    .map((s, i) => ({ num: i + 1, label: s.title || `Sección ${i + 1}` }))
    .slice(0, 16);
  const execIdx = final.findIndex(s => s.type === 'executive_summary');
  final.splice(execIdx >= 0 ? execIdx + 1 : 1, 0, { type: 'index', title: 'Contenido', items: indexItems });

  log.info(`📊 Presentación final: ${final.length} diapositivas (tema: ${selectedTheme.name})`);

  return {
    slides: final,
    theme: selectedTheme,
    meta: {
      title: options.title || 'Reporte',
      sourceFile: fileName,
      generatedAt: new Date().toISOString(),
      totalSlides: final.length,
    },
  };
}
