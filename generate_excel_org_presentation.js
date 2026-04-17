const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const XLSX = require('xlsx');
const pptxgen = require('pptxgenjs');
const {
  assertNoUnauthorizedDataDependencies,
  buildAuditRecord,
  createProcessingContext,
  decideSafeMode,
  ensureExplicitInputDependencies,
  validateOrganizerConsistency,
  writeAuditRecord,
} = require('./utils/presentation-integrity.cjs');
const pptx = { ShapeType: { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse' } };

const cliInput = process.argv[2];
const PROCESSING_CONTEXT = createProcessingContext(cliInput, __dirname);
ensureExplicitInputDependencies(PROCESSING_CONTEXT);
const INPUT_FILE = PROCESSING_CONTEXT.inputFile;
const SOURCE_FILE_LABEL = PROCESSING_CONTEXT.inputLabel;
const OUTPUT_FILE = PROCESSING_CONTEXT.outputFile;
const REQUESTED_VISUAL_MODE = String(process.env.SOCYA_PRESENTATION_MODE || '').trim().toLowerCase();
const PRESENTATION_VISUAL_MODE = ['charts', 'tables', 'mixed'].includes(REQUESTED_VISUAL_MODE) ? REQUESTED_VISUAL_MODE : 'mixed';

const COLORS = {
  navy: '94C24A',
  navy2: '7EAA38',
  navy3: '47661E',
  panel: 'F2F8E7',
  panelAlt: 'D7EAAE',
  borderDark: 'B9D98A',
  slate: '5B6B7C',
  green: '94C24A',
  greenSoft: 'F2F9E6',
  amber: 'BCD88B',
  amberSoft: 'F5FAEC',
  gold: 'A4CD62',
  goldSoft: 'F1F8E5',
  aqua: '85BC84',
  aquaSoft: 'EDF7EA',
  red: 'C44F4F',
  redSoft: 'FBE9E8',
  ink: '1E293B',
  muted: '5B7253',
  textLight: 'FFFFFF',
  textSoft: 'F7FBF0',
  line: 'D9E8C4',
  bg: 'FAFDF5',
  card: 'FFFFFF',
  white: 'FFFFFF',
};

const CHART_COLORS = [
  COLORS.green,
  COLORS.aqua,
  COLORS.amber,
  'A8CF69',
  'CAE7A6',
  '7EAA38',
  'D64550',
  '9CA3AF',
];

function hashString(input = '') {
  let hash = 2166136261;
  for (let i = 0; i < input.length; i += 1) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

const GREEN_LAYOUT_FAMILIES = [
  {
    id: 'bosque-claro',
    accentOrder: [COLORS.green, COLORS.aqua, COLORS.gold, COLORS.amber],
    chartPalette: [COLORS.green, COLORS.aqua, COLORS.gold, 'A8CF69', COLORS.amber],
    sidePanelColor: COLORS.panelAlt,
    sidePanelTransparency: 16,
    topBandColor: COLORS.navy,
    topBandAccent: COLORS.navy2,
    surfaceSoft: COLORS.greenSoft,
    fontPair: { head: 'Calibri', body: 'Calibri' },
  },
  {
    id: 'jardin-editorial',
    accentOrder: [COLORS.aqua, COLORS.green, COLORS.amber, COLORS.gold],
    chartPalette: [COLORS.aqua, COLORS.green, COLORS.amber, COLORS.gold, 'B7DB80'],
    sidePanelColor: COLORS.aquaSoft,
    sidePanelTransparency: 10,
    topBandColor: COLORS.navy2,
    topBandAccent: COLORS.green,
    surfaceSoft: 'F4FBF6',
    fontPair: { head: 'Cambria', body: 'Calibri' },
  },
  {
    id: 'musgo-premium',
    accentOrder: [COLORS.gold, COLORS.green, COLORS.aqua, COLORS.amber],
    chartPalette: [COLORS.gold, COLORS.green, COLORS.aqua, 'B8D77B', COLORS.amber],
    sidePanelColor: COLORS.goldSoft,
    sidePanelTransparency: 9,
    topBandColor: COLORS.navy3,
    topBandAccent: COLORS.navy2,
    surfaceSoft: 'FBFDF5',
    fontPair: { head: 'Trebuchet MS', body: 'Calibri' },
  },
  {
    id: 'hoja-viva',
    accentOrder: ['A8CF69', COLORS.green, COLORS.aqua, COLORS.gold],
    chartPalette: ['A8CF69', COLORS.green, COLORS.aqua, COLORS.gold, COLORS.amber],
    sidePanelColor: COLORS.greenSoft,
    sidePanelTransparency: 7,
    topBandColor: COLORS.navy,
    topBandAccent: COLORS.green,
    surfaceSoft: 'F5FBEC',
    fontPair: { head: 'Arial', body: 'Calibri' },
  },
  {
    id: 'eucalipto-suave',
    accentOrder: [COLORS.green, 'A8CF69', COLORS.aqua, COLORS.gold],
    chartPalette: [COLORS.green, 'A8CF69', COLORS.aqua, COLORS.gold, COLORS.amber],
    sidePanelColor: 'EEF6E2',
    sidePanelTransparency: 6,
    topBandColor: '89BC3F',
    topBandAccent: '6E9A2D',
    surfaceSoft: 'F8FCF1',
    fontPair: { head: 'Cambria', body: 'Calibri' },
  },
  {
    id: 'selva-ligera',
    accentOrder: [COLORS.aqua, '9BCA57', COLORS.green, COLORS.gold],
    chartPalette: [COLORS.aqua, '9BCA57', COLORS.green, COLORS.gold, 'C8E29B'],
    sidePanelColor: 'EAF5EA',
    sidePanelTransparency: 11,
    topBandColor: '7FB64D',
    topBandAccent: '5E8F30',
    surfaceSoft: 'F4FBF4',
    fontPair: { head: 'Trebuchet MS', body: 'Calibri' },
  },
];

function buildDeckStyle(seedSource) {
  const family = GREEN_LAYOUT_FAMILIES[hashString(seedSource) % GREEN_LAYOUT_FAMILIES.length];
  const variant = (scope, count) => (count <= 1 ? 0 : hashString(`${seedSource}|${scope}`) % count);

  return {
    ...family,
    seedSource,
    variants: {
      decor: variant('decor', 3),
      header: variant('header', 3),
      metricCard: variant('metricCard', 3),
      insightBox: variant('insightBox', 3),
      commissionsCover: variant('commissionsCover', 3),
      commissionsClosing: variant('commissionsClosing', 3),
      genericCover: variant('genericCover', 3),
      genericOverview: variant('genericOverview', 3),
      genericFocus: variant('genericFocus', 3),
      genericMap: variant('genericMap', 2),
      genericClosing: variant('genericClosing', 3),
    },
  };
}

let DECK_STYLE = buildDeckStyle(`${SOURCE_FILE_LABEL}|socya`);

function resolveDeckSeed() {
  const explicitSeed = cleanText(process.env.SOCYA_LAYOUT_SEED);
  if (explicitSeed) return explicitSeed;
  const baseKey = path.resolve(INPUT_FILE).toLowerCase();
  const mode = cleanText(process.env.SOCYA_LAYOUT_MODE).toLowerCase() || 'stable';
  const safetyTag = process.env.SOCYA_SAFE_MODE === '1' ? 'safe' : 'standard';
  return `${baseKey}|${mode}|${safetyTag}|${PRESENTATION_VISUAL_MODE}`;
}

function configureDeckStyle(seedSource) {
  DECK_STYLE = buildDeckStyle(seedSource || `${SOURCE_FILE_LABEL}|socya`);
}

function getDeckVariant(scope) {
  return DECK_STYLE.variants[scope] || 0;
}

function getDeckFont(role = 'body') {
  return role === 'head' ? DECK_STYLE.fontPair.head : DECK_STYLE.fontPair.body;
}

function getDeckAccent(index = 0) {
  return DECK_STYLE.accentOrder[index % DECK_STYLE.accentOrder.length];
}

function getDeckChartPalette(count = 4) {
  return Array.from({ length: count }, (_, index) => DECK_STYLE.chartPalette[index % DECK_STYLE.chartPalette.length]);
}

function getScopedVariant(scope, key, count) {
  return count <= 1 ? 0 : hashString(`${DECK_STYLE.seedSource}|${scope}|${key}`) % count;
}

function getSoftAccentColor(accent) {
  const palette = {
    [COLORS.green]: COLORS.greenSoft,
    [COLORS.aqua]: COLORS.aquaSoft,
    [COLORS.gold]: COLORS.goldSoft,
    [COLORS.amber]: COLORS.amberSoft,
    A8CF69: 'F5FBEC',
  };
  return palette[accent] || DECK_STYLE.surfaceSoft || COLORS.greenSoft;
}

function isSoftAccentColor(color) {
  return [
    COLORS.greenSoft,
    COLORS.aquaSoft,
    COLORS.goldSoft,
    COLORS.amberSoft,
    COLORS.panel,
    COLORS.panelAlt,
    'F5FBEC',
    'F8FCF1',
    'F4FBF4',
    'EEF6E2',
    'EAF5EA',
  ].includes(color);
}

function getHeaderLayoutMetrics(headerVariant, title, subtitle) {
  const titleX = headerVariant === 2 ? 0.82 : 0.55;
  const titleY = headerVariant === 1 ? 1.12 : 1.18;
  const titleW = headerVariant === 1 ? 8.95 : headerVariant === 2 ? 8.35 : 8.6;
  const titleLines = Math.max(1, Math.min(3, estimateWrappedLines(title, titleW, 28)));
  const titleH = Math.max(0.6, Math.min(1.05, 0.22 + titleLines * 0.22));
  const subtitleW = headerVariant === 1 ? 9.7 : headerVariant === 2 ? 9.0 : 9.5;
  const subtitleLines = subtitle ? Math.max(1, Math.min(3, estimateWrappedLines(subtitle, subtitleW, 10.5))) : 0;
  const subtitleH = subtitle ? Math.max(0.3, Math.min(0.54, 0.1 + subtitleLines * 0.14)) : 0;
  const subtitleY = titleY + titleH + 0.08;

  return {
    titleX,
    titleY,
    titleW,
    titleLines,
    titleH,
    subtitleW,
    subtitleLines,
    subtitleH,
    subtitleY,
    contentBottom: subtitle ? subtitleY + subtitleH : titleY + titleH,
  };
}

function getSafeContentStartY(title, subtitle, desiredY = 1.95, extraGap = 0.28) {
  const headerMetrics = getHeaderLayoutMetrics(getDeckVariant('header'), title, subtitle);
  return Math.max(desiredY, headerMetrics.contentBottom + extraGap);
}

const MONTH_INDEX = {
  enero: 1, ene: 1,
  febrero: 2, feb: 2,
  marzo: 3, mar: 3,
  abril: 4, abr: 4,
  mayo: 5, may: 5,
  junio: 6, jun: 6,
  julio: 7, jul: 7,
  agosto: 8, ago: 8,
  septiembre: 9, setiembre: 9, sep: 9, sept: 9,
  octubre: 10, oct: 10,
  noviembre: 11, nov: 11,
  diciembre: 12, dic: 12,
};

function cleanText(value, maxLen) {
  if (value === null || value === undefined) return '';
  let text = String(value)
    .replace(/_x000d_/gi, ' ')
    .replace(/x000d_/gi, ' ')
    .replace(/\r?\n+/g, ' ')
    .replace(/\u00a0/g, ' ')
    .replace(/^[\u00FC\u00F9\u00B3\u2022\-]+\s*/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  const mojibakeMap = [
    ['ÔÇö', '—'],
    ['ÔÇ£', '"'],
    ['ÔÇ¥', "'"],
    ['├ü', 'Á'],
    ['├ü', 'Á'],
    ['├│', 'ó'],
    ['├ô', 'Ó'],
    ['├®', 'é'],
    ['├ë', 'É'],
    ['├¡', 'á'],
    ['├Ä', 'Á'],
    ['├®', 'é'],
    ['├¡', 'á'],
    ['├¡', 'á'],
    ['├║', 'ú'],
    ['├Ü', 'Ú'],
    ['├▒', 'ñ'],
    ['├æ', 'Ñ'],
    ['├¡', 'á'],
    ['├®', 'é'],
    ['├¡', 'á'],
    ['├¡', 'á'],
    ['Â', ''],
  ];
  mojibakeMap.forEach(([from, to]) => {
    text = text.split(from).join(to);
  });

  text = text
    .replace(/Gesti¾n/g, 'Gestión')
    .replace(/gesti¾n/g, 'gestión')
    .replace(/informaci¾n/g, 'información')
    .replace(/revisi¾n/g, 'revisión')
    .replace(/Revisi¾n/g, 'Revisión')
    .replace(/ejecuci¾n/g, 'ejecución')
    .replace(/concentraci¾n/g, 'concentración')
    .replace(/Presentaci¾n/g, 'Presentación')
    .replace(/distribuci¾n/g, 'distribución')
    .replace(/evaluaci¾n/g, 'evaluación')
    .replace(/tÚrminos/g, 'términos')
    .replace(/comitÚ/g, 'comité')
    .replace(/electr¾nicos/g, 'electrónicos')
    .replace(/MetodologÝa/g, 'Metodología')
    .replace(/·tiles/g, 'útiles')
    .replace(/·nicos/g, 'únicos');

  if (!text || text === '—' || text === 'ù') return '';
  if (maxLen && text.length > maxLen) return `${text.slice(0, maxLen - 1).trim()}…`;
  return text;
}

function humanizeLabel(value, maxLen) {
  const text = cleanText(value, maxLen);
  if (!text) return '';
  const normalized = text.replace(/[_-]+/g, ' ');
  return normalized.replace(/\s+/g, ' ').trim();
}

function normalizeHeaderKey(value) {
  return cleanText(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function ensureUniqueHeaders(headers) {
  const counts = new Map();
  return headers.map((header, index) => {
    const base = cleanText(header) || `Col_${index + 1}`;
    const seen = counts.get(base) || 0;
    counts.set(base, seen + 1);
    return seen ? `${base}_${seen + 1}` : base;
  });
}

function monthSortIndex(text) {
  const normalized = normalizeHeaderKey(text);
  return MONTH_INDEX[normalized] || null;
}

function safeUpper(value) {
  return cleanText(value).toUpperCase();
}

function fmtCurrency(value) {
  const num = Number(value || 0);
  return new Intl.NumberFormat('es-CO', {
    style: 'currency',
    currency: 'COP',
    maximumFractionDigits: 0,
  }).format(num);
}

function fmtCompactCurrency(value) {
  const num = Number(value || 0);
  return new Intl.NumberFormat('es-CO', {
    notation: 'compact',
    maximumFractionDigits: 1,
  }).format(num);
}

function fmtNumber(value) {
  return new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 }).format(Number(value || 0));
}

function fmtPercent(value, digits = 1) {
  return `${(Number(value || 0) * 100).toFixed(digits)}%`;
}

function ratio(num, den) {
  return den ? Number(num || 0) / den : 0;
}

function normalizeProgress(value) {
  const num = Number(value || 0);
  if (!Number.isFinite(num)) return 0;
  if (num > 1) return Math.max(0, Math.min(num / 100, 1));
  return Math.max(0, Math.min(num, 1));
}

function isUsefulCell(value) {
  const text = cleanText(value);
  if (!text) return false;
  return !/^\d+$/.test(text) && !/^unnamed/i.test(text) && !/^col[_ -]?\d+$/i.test(text) && !/^type:/i.test(text) && !/^title:/i.test(text) && !/^subtitle:/i.test(text);
}

function findStructuredSheet(workbook, matcher) {
  const name = workbook.SheetNames.find((sheetName) => matcher(sheetName.toLowerCase()));
  if (!name) return null;

  const ws = workbook.Sheets[name];
  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const nonEmptyRows = rawRows.filter((row) => Array.isArray(row) && row.some((cell) => cleanText(cell) || typeof cell === 'number'));
  if (!nonEmptyRows.length) return null;

  let headerIndex = 0;
  let bestScore = -1;

  for (let i = 0; i < Math.min(nonEmptyRows.length, 20); i += 1) {
    const row = nonEmptyRows[i] || [];
    const score = row.filter(isUsefulCell).length;
    if (score > bestScore) {
      bestScore = score;
      headerIndex = i;
    }
  }

  const headerRow = nonEmptyRows[headerIndex] || [];
  const candidateHeaders = ensureUniqueHeaders(headerRow.map((cell, idx) => {
    const text = cleanText(cell);
    return text || `Col_${idx + 1}`;
  }));

  const keepIndexes = candidateHeaders
    .map((header, idx) => {
      const hasBodyData = nonEmptyRows
        .slice(headerIndex + 1)
        .some((row) => cleanText(row[idx]) || typeof row[idx] === 'number');
      return isUsefulCell(header) && hasBodyData ? idx : -1;
    })
    .filter((idx) => idx >= 0);

  const headers = keepIndexes.map((idx) => candidateHeaders[idx]);
  const rows = nonEmptyRows
    .slice(headerIndex + 1)
    .map((row) => {
      const obj = {};
      keepIndexes.forEach((idx, pos) => {
        obj[headers[pos]] = row[idx] ?? null;
      });
      return obj;
    })
    .filter((row) => Object.values(row).some((value) => cleanText(value) || typeof value === 'number'));

  return { name, headers, rows };
}

function isMeaningfulValue(value) {
  return cleanText(value) || typeof value === 'number' || value instanceof Date;
}

function isGhostText(value) {
  const text = cleanText(value).toLowerCase();
  if (!text) return true;
  return ['—', '-', 'n/a', 'na', 'null', 'undefined', 'nan', '#n/a', 's/d', 'sin dato', 'sin datos'].includes(text);
}

function isRenderableValue(value) {
  if (value instanceof Date) return true;
  if (typeof value === 'number') return Number.isFinite(value);
  return !isGhostText(value);
}

function countMeaningfulCells(row) {
  return (row || []).filter((cell) => isMeaningfulValue(cell)).length;
}

function isBlankishRow(row) {
  return countMeaningfulCells(row) === 0;
}

function rowTextValues(row) {
  return (row || [])
    .filter((cell) => isMeaningfulValue(cell))
    .map((cell) => cleanText(cell))
    .filter(Boolean);
}

function normalizeNarrativeText(text) {
  return cleanText(text)
    .replace(/^(title|titulo|subtitle|subtitulo|note|nota|observacion|observaciones)\s*:\s*/i, '')
    .trim();
}

function classifyNarrativeText(text, pendingMeta = {}) {
  const normalized = cleanText(text);
  if (!normalized) return null;
  if (/^(title|titulo)\s*:/i.test(normalized)) return 'title';
  if (/^(subtitle|subtitulo)\s*:/i.test(normalized)) return 'subtitle';
  if (/^(note|nota|observacion|observaciones)\s*:/i.test(normalized)) return 'note';
  if (/^(resumen ejecutivo|conclusiones|observaciones|hallazgos clave|comentarios|consideraciones finales)\b/i.test(normalized)) return 'title';
  if (/^(detalle|contexto|alcance|metodologia|antecedentes|analisis)\b/i.test(normalized) && normalized.length <= 90) return 'subtitle';
  if (normalized.length <= 70 && !pendingMeta.title && !/^\d+$/.test(normalized)) return 'title';
  if (normalized.length <= 140 && pendingMeta.title && !pendingMeta.subtitle) return 'subtitle';
  if (normalized.length >= 45) return 'note';
  return null;
}

function isNarrativeRow(row, pendingMeta = {}) {
  const texts = rowTextValues(row);
  if (!texts.length || texts.length > 2) return false;
  if ((row || []).some((cell) => typeof cell === 'number' || cell instanceof Date)) return false;
  const merged = texts.join(' ');
  return !!classifyNarrativeText(merged, pendingMeta);
}

function scoreHeaderRowCandidate(row, followingRows = []) {
  const usefulCount = (row || []).filter(isUsefulCell).length;
  if (usefulCount < 2) return -1;

  const textLengths = rowTextValues(row).map((value) => value.length);
  const avgLength = textLengths.length ? textLengths.reduce((acc, value) => acc + value, 0) / textLengths.length : 0;
  const futureSupport = followingRows.reduce((acc, nextRow) => acc + Math.min(usefulCount, countMeaningfulCells(nextRow || [])), 0);

  return usefulCount * 3 + futureSupport - (avgLength > 40 ? 4 : 0);
}

function isStrongStackedHeaderRow(row, followingRows = [], currentHeaders = []) {
  const texts = rowTextValues(row);
  if (texts.length < 2) return false;
  if ((row || []).some((cell) => typeof cell === 'number' || cell instanceof Date)) return false;

  const avgLength = texts.reduce((acc, value) => acc + value.length, 0) / texts.length;
  if (avgLength > 26) return false;
  if (texts.some((value) => value.length > 36)) return false;

  const normalizedTexts = texts.map((value) => normalizeHeaderKey(value));
  const currentHeaderKeys = currentHeaders.map((value) => normalizeHeaderKey(value));
  const identicalToCurrent = currentHeaderKeys.length
    && normalizedTexts.length === currentHeaderKeys.length
    && normalizedTexts.every((value, index) => value === currentHeaderKeys[index]);
  if (identicalToCurrent) return false;

  const supportRows = followingRows.slice(0, 3);
  const supportScore = supportRows.filter((nextRow) => countMeaningfulCells(nextRow || []) >= Math.max(2, texts.length - 1)).length;
  return supportScore >= 2 && scoreHeaderRowCandidate(row, supportRows) >= Math.max(8, texts.length * 4);
}

function buildDisplayBlockName(sheetName, meta, blockIndex, totalBlocks) {
  const explicitTitle = cleanText(meta.title, 42);
  if (explicitTitle) return explicitTitle;
  const explicitSubtitle = cleanText(meta.subtitle, 38);
  if (explicitSubtitle) return explicitSubtitle;
  const genericSheet = isGenericSheetLabel(sheetName);
  const headerFocus = Array.isArray(meta.headers)
    ? meta.headers.slice(0, 2).map((header) => humanizeLabel(header, 16)).filter(Boolean).join(' / ')
    : '';
  if (headerFocus) return genericSheet ? headerFocus : `${humanizeLabel(sheetName, 18)} · ${headerFocus}`;
  if (totalBlocks > 1) return genericSheet ? `Bloque ${blockIndex + 1}` : `${humanizeLabel(sheetName, 28)} · Bloque ${blockIndex + 1}`;
  return humanizeLabel(sheetName, 36);
}

function collectNarrativeRows(rawRows, startIndex, pendingMeta) {
  const meta = {
    title: pendingMeta?.title || '',
    subtitle: pendingMeta?.subtitle || '',
    notes: [...(pendingMeta?.notes || [])],
  };
  let index = startIndex;

  while (index < rawRows.length) {
    const row = rawRows[index] || [];
    if (isBlankishRow(row)) {
      index += 1;
      break;
    }
    if (!isNarrativeRow(row, meta)) break;

    const rawNarrative = rowTextValues(row).join(' ');
    const type = classifyNarrativeText(rawNarrative, meta);
    const text = normalizeNarrativeText(rawNarrative);
    if (type === 'title' && !meta.title) meta.title = text;
    else if (type === 'subtitle' && !meta.subtitle) meta.subtitle = text;
    else if (text) {
      const lastNote = meta.notes[meta.notes.length - 1];
      if (lastNote && (lastNote.length + text.length) <= 240) {
        meta.notes[meta.notes.length - 1] = `${lastNote} ${text}`.trim();
      } else {
        meta.notes.push(text);
      }
    }
    index += 1;
  }

  return { meta, nextIndex: index };
}

function extractStructuredSheetModels(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) return [];

  const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  if (!rawRows.length) return [];

  const blocks = [];
  let index = 0;
  let pendingMeta = { title: '', subtitle: '', notes: [] };

  while (index < rawRows.length) {
    const row = rawRows[index] || [];

    if (isBlankishRow(row)) {
      index += 1;
      continue;
    }

    if (isNarrativeRow(row, pendingMeta)) {
      const narrative = collectNarrativeRows(rawRows, index, pendingMeta);
      pendingMeta = narrative.meta;
      index = narrative.nextIndex;
      continue;
    }

    const candidateWindow = rawRows.slice(index, Math.min(rawRows.length, index + 4));
    let bestCandidate = null;
    candidateWindow.forEach((candidateRow, offset) => {
      const score = scoreHeaderRowCandidate(candidateRow, rawRows.slice(index + offset + 1, index + offset + 5));
      if (!bestCandidate || score > bestCandidate.score) {
        bestCandidate = { index: index + offset, score };
      }
    });

    if (!bestCandidate || bestCandidate.score < 6) {
      index += 1;
      continue;
    }

    const headerIndex = bestCandidate.index;
    const headerRow = rawRows[headerIndex] || [];
    const candidateHeaders = ensureUniqueHeaders(headerRow.map((cell, idx) => cleanText(cell) || `Col_${idx + 1}`));
    const keepIndexes = candidateHeaders
      .map((header, idx) => {
        const hasBodyData = rawRows
          .slice(headerIndex + 1, headerIndex + 8)
          .some((nextRow) => isMeaningfulValue((nextRow || [])[idx]));
        return isUsefulCell(header) && hasBodyData ? idx : -1;
      })
      .filter((idx) => idx >= 0);

    if (keepIndexes.length < 2) {
      index = headerIndex + 1;
      continue;
    }

    const headers = keepIndexes.map((idx) => candidateHeaders[idx]);
    const rows = [];
    let cursor = headerIndex + 1;
    let blankStreak = 0;

    while (cursor < rawRows.length) {
      const currentRow = rawRows[cursor] || [];
      if (isBlankishRow(currentRow)) {
        blankStreak += 1;
        if (rows.length >= 2 && blankStreak >= 1) break;
        cursor += 1;
        continue;
      }

      blankStreak = 0;

      if (rows.length >= 2 && isNarrativeRow(currentRow)) break;
      if (rows.length >= 2 && isStrongStackedHeaderRow(currentRow, rawRows.slice(cursor + 1, cursor + 4), headers)) break;

      const record = {};
      let filledValues = 0;
      keepIndexes.forEach((cellIndex, position) => {
        const value = currentRow[cellIndex] ?? null;
        record[headers[position]] = value;
        if (isMeaningfulValue(value)) filledValues += 1;
      });

      if (filledValues >= Math.max(2, Math.ceil(keepIndexes.length * 0.35))) {
        rows.push(record);
      }

      cursor += 1;
    }

    if (rows.length >= 2) {
      const trailingMeta = collectNarrativeRows(rawRows, cursor, { title: '', subtitle: '', notes: [] });
      const trailingBelongsToNextBlock = !!(trailingMeta.meta.title || trailingMeta.meta.subtitle);
      blocks.push({
        name: sheetName,
        sourceSheetName: sheetName,
        title: pendingMeta.title || '',
        subtitle: pendingMeta.subtitle || '',
        notes: [
          ...pendingMeta.notes,
          ...(trailingBelongsToNextBlock ? [] : trailingMeta.meta.notes),
        ].slice(0, 3),
        headers,
        rows,
      });
      pendingMeta = trailingBelongsToNextBlock ? trailingMeta.meta : { title: '', subtitle: '', notes: [] };
      index = trailingMeta.nextIndex;
    } else {
      index = headerIndex + 1;
    }
  }

  return blocks.map((block, blockIndex) => ({
    ...block,
    blockIndex,
    blockCount: blocks.length,
    displayName: buildDisplayBlockName(sheetName, block, blockIndex, blocks.length),
  }));
}

function sanitizeChartEntries(entries) {
  return entries.filter(([label, value]) => {
    const cleanLabel = cleanText(label);
    if (!cleanLabel || isGhostText(cleanLabel)) return false;
    return Number.isFinite(Number(value));
  });
}

function hasMeaningfulChartEntries(entries = [], seriesType = 'count') {
  const numericValues = entries
    .map(([, value]) => Number(value))
    .filter((value) => Number.isFinite(value) && value > 0);

  if (numericValues.length < 2) return false;

  const maxValue = Math.max(...numericValues);
  const minValue = Math.min(...numericValues);
  const total = sumNumericValues(numericValues);
  const spread = maxValue ? (maxValue - minValue) / maxValue : 0;
  const leaderShare = total ? maxValue / total : 0;
  const uniqueValues = new Set(numericValues.map((value) => value.toFixed(6))).size;

  if (uniqueValues < 2) return false;
  if (seriesType === 'count') {
    return spread >= 0.18 || leaderShare >= 0.3;
  }
  return spread >= 0.08 || leaderShare >= 0.26;
}

function selectBestTableHeaders(sheet, profiles, maxHeaders = 5) {
  const scored = sheet.headers.map((header) => {
    const values = sheet.rows.map((row) => row[header]);
    const filledCount = values.filter((value) => isRenderableValue(value)).length;
    const profile = profiles.find((item) => item.header === header);
    const preferredBonus = /estado|categoria|tipo|canal|sede|area|ciudad|region|mes|fecha|producto|linea|valor|venta|ingreso|costo|gasto|precio|monto|total/.test(normalizeHeaderKey(header)) ? 2 : 0;
    return {
      header,
      profile,
      fillRate: sheet.rows.length ? filledCount / sheet.rows.length : 0,
      score: filledCount + preferredBonus + ((profile?.type === 'currency' || profile?.type === 'number' || profile?.type === 'percent') ? 1 : 0),
    };
  });

  return scored
    .filter((item) => item.fillRate >= 0.6)
    .sort((a, b) => b.score - a.score)
    .slice(0, maxHeaders)
    .map((item) => item.header);
}

function isExecutiveNoiseSheet(summary) {
  const label = normalizeHeaderKey(summary.rawName || summary.name || '');
  if (!label) return true;
  if (/^(hoja|sheet)\d*$/.test(label)) return true;
  if (/posible generalidad|control de cambios|cronograma|notas|anex|soporte|cambios/.test(label)) return true;
  if (summary.rowCount < 3 && !summary.blockTitle) return true;
  if (!hasRenderableChartSummary(summary) && !hasRenderableTableSummary(summary) && (summary.insights?.length || 0) < 2) return true;
  return false;
}

function prioritizeExecutiveSheets(summaries = []) {
  const preferred = summaries.filter((sheet) => !isExecutiveNoiseSheet(sheet));
  return preferred.length ? preferred : summaries;
}

function buildCleanTablePreview(sheet, headers, profiles, maxRows = 8) {
  const candidateRows = sheet.rows
    .map((row) => ({
      row,
      filled: headers.filter((header) => isRenderableValue(row[header])).length,
    }))
    .filter((item) => item.filled >= Math.max(2, headers.length - 1))
    .sort((a, b) => b.filled - a.filled)
    .slice(0, maxRows)
    .map((item) => item.row);

  const cleanedRows = candidateRows.map((row) =>
    headers.map((header) => formatGenericCell(row[header], profiles.find((profile) => profile.header === header)?.type || 'text'))
  );

  const nonGhostColumnIndexes = headers
    .map((header, index) => ({
      header,
      index,
      useful: cleanedRows.some((row) => !isGhostText(row[index])),
    }))
    .filter((item) => item.useful);

  const finalHeaders = nonGhostColumnIndexes.map((item) => item.header);
  const finalRows = cleanedRows
    .map((row) => nonGhostColumnIndexes.map((item) => row[item.index]))
    .filter((row) => row.some((cell) => !isGhostText(cell)));

  return { headers: finalHeaders, rows: finalRows };
}

function countRenderablePreviewCells(row = []) {
  return row.filter((cell) => {
    if (cell && typeof cell === 'object' && Object.prototype.hasOwnProperty.call(cell, 'text')) {
      return !isGhostText(cell.text);
    }
    return !isGhostText(cell);
  }).length;
}

function isGenericSheetLabel(name) {
  const normalized = normalizeHeaderKey(name);
  return /^(hoja|sheet)\d*$/.test(normalized) || /^(tabla|table)\d*$/.test(normalized);
}

function findKey(headers, patterns) {
  const normalized = headers.map((header) => ({
    raw: header,
    norm: normalizeHeaderKey(header),
  }));

  for (const pattern of patterns) {
    const match = normalized.find((item) => item.norm.includes(pattern));
    if (match) return match.raw;
  }
  return null;
}

function countBy(rows, key) {
  const counts = new Map();
  rows.forEach((row) => {
    const raw = cleanText(row[key]);
    if (!raw) return;
    counts.set(raw, (counts.get(raw) || 0) + 1);
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1]);
}

function sumBy(rows, groupKey, valueKey) {
  const sums = new Map();
  rows.forEach((row) => {
    const group = cleanText(row[groupKey]);
    if (!group) return;
    const value = Number(row[valueKey] || 0);
    sums.set(group, (sums.get(group) || 0) + value);
  });
  return [...sums.entries()].sort((a, b) => b[1] - a[1]);
}

function getStatusStyle(status) {
  const key = safeUpper(status);
  if (key.includes('COMPLET')) return { fill: COLORS.greenSoft, color: '2F6B18' };
  if (key.includes('EJECUC') || key.includes('PROCES')) return { fill: COLORS.amberSoft, color: '8A5A00' };
  if (key.includes('PENDIENT')) return { fill: COLORS.redSoft, color: '8F2F2F' };
  if (key.includes('RECHAZ')) return { fill: COLORS.redSoft, color: '8F2F2F' };
  return { fill: COLORS.aquaSoft, color: '0C5A67' };
}

function getStateChartColor(label, index) {
  const key = safeUpper(label);
  if (key.includes('CONTABILIZ') || key.includes('LEGALIZ')) return COLORS.green;
  if (key.includes('RECHAZ')) return COLORS.red;
  if (key.includes('ESPERA') || key.includes('SOLICIT')) return COLORS.amber;
  return CHART_COLORS[index % CHART_COLORS.length];
}

function splitActionLines(text) {
  return splitTextForDisplay(text, 92).slice(0, 8);
}

function splitTextForDisplay(text, maxChars = 90) {
  const normalized = cleanText(text);
  if (!normalized) return [];

  const chunks = [];
  normalized
    .split(/\s{2,}|[.;](?=\s+[A-ZÁÉÍÓÚÑ])/)
    .map((part) => cleanText(part))
    .filter(Boolean)
    .forEach((part) => {
      if (part.length <= maxChars) {
        chunks.push(part);
        return;
      }

      const words = part.split(/\s+/);
      let current = '';
      words.forEach((word) => {
        const candidate = current ? `${current} ${word}` : word;
        if (candidate.length > maxChars && current) {
          chunks.push(current);
          current = word;
        } else {
          current = candidate;
        }
      });
      if (current) chunks.push(current);
    });

  return chunks;
}

function chunkArray(items, chunkSize) {
  const result = [];
  for (let i = 0; i < items.length; i += chunkSize) {
    result.push(items.slice(i, i + chunkSize));
  }
  return result;
}

function estimateWrappedLines(text, colWidth, fontSize = 8) {
  const content = cleanText(text);
  if (!content) return 1;
  const approxCharsPerLine = Math.max(10, Math.floor(colWidth * 13 * (8.2 / fontSize)));
  const words = content.split(/\s+/);
  let lines = 1;
  let current = 0;

  words.forEach((word) => {
    const wordLength = word.length + (current ? 1 : 0);
    if (current + wordLength > approxCharsPerLine) {
      lines += 1;
      current = word.length;
    } else {
      current += wordLength;
    }
  });

  return Math.max(lines, Math.ceil(content.length / Math.max(approxCharsPerLine, 1)));
}

function getAdaptiveRowHeights(bodyRows, colWidths, fontSize, minHeight, maxHeight) {
  return bodyRows.map((row) => {
    const maxLines = row.reduce((acc, cell, index) => {
      const text = typeof cell === 'object' ? cell.text : cell;
      return Math.max(acc, estimateWrappedLines(text, colWidths[index] || 1, fontSize));
    }, 1);
    return Math.min(maxHeight, Math.max(minHeight, 0.12 + maxLines * 0.15));
  });
}

function buildBodyCell(text, options = {}) {
  return {
    text: cleanText(text),
    options: {
      color: COLORS.ink,
      fontFace: 'Calibri',
      fontSize: 8.3,
      margin: 0.03,
      valign: 'mid',
      breakLine: false,
      fit: 'shrink',
      ...options,
    },
  };
}

function renderManualTable(slide, config) {
  const {
    x,
    y,
    colWidths,
    headerHeight,
    rowHeight,
    headers,
    rows,
    rowHeights,
    defaultFontSize,
    minFontSize,
  } = config;
  const safeMinFontSize = Math.max(8.2, minFontSize || 8.2);
  const safeDefaultFontSize = Math.max(safeMinFontSize + 0.35, defaultFontSize || 8.6);

  let currentX = x;
  headers.forEach((header, colIndex) => {
    slide.addShape(pptx.ShapeType.rect, {
      x: currentX,
      y,
      w: colWidths[colIndex],
      h: headerHeight,
      line: { color: COLORS.line, pt: 0.7 },
      fill: { color: COLORS.navy2 },
    });
    slide.addText(header.text, {
      x: currentX + 0.02,
      y: y + 0.02,
      w: colWidths[colIndex] - 0.04,
      h: headerHeight - 0.04,
      fontSize: Math.max(8.8, header.fontSize || 9),
      bold: true,
      color: COLORS.white,
      align: 'center',
      valign: 'mid',
      margin: 0.03,
      fontFace: 'Calibri',
      fit: 'shrink',
    });
    currentX += colWidths[colIndex];
  });

  rows.forEach((row, rowIndex) => {
    const currentRowHeight = Array.isArray(rowHeights) ? rowHeights[rowIndex] : rowHeight;
    const currentY = y + headerHeight + (Array.isArray(rowHeights)
      ? rowHeights.slice(0, rowIndex).reduce((acc, value) => acc + value, 0)
      : rowIndex * rowHeight);
    let cellX = x;
    row.forEach((cell, colIndex) => {
      const cellFill = cell.fill || (rowIndex % 2 === 0 ? 'F8FAFC' : COLORS.white);
      slide.addShape(pptx.ShapeType.rect, {
        x: cellX,
        y: currentY,
        w: colWidths[colIndex],
        h: currentRowHeight,
        line: { color: COLORS.line, pt: 0.6 },
        fill: { color: cellFill },
      });
      const baseFontSize = Math.max(safeDefaultFontSize, cell.fontSize || safeDefaultFontSize);
      const maxLines = estimateWrappedLines(cell.text, colWidths[colIndex], baseFontSize);
      const adaptiveFontSize = Math.max(safeMinFontSize, Math.min(baseFontSize, maxLines > 4 ? baseFontSize - 0.8 : baseFontSize));
      slide.addText(cleanText(cell.text), {
        x: cellX + 0.03,
        y: currentY + 0.03,
        w: colWidths[colIndex] - 0.06,
        h: currentRowHeight - 0.06,
        fontSize: adaptiveFontSize,
        bold: !!cell.bold,
        color: cell.color || COLORS.ink,
        align: cell.align || 'left',
        valign: cell.valign || (currentRowHeight >= 0.86 ? 'top' : 'mid'),
        margin: 0.04,
        fontFace: 'Calibri',
        breakLine: true,
        fit: 'shrink',
      });
      cellX += colWidths[colIndex];
    });
  });
}

function computeManualRowHeights(rows, colWidths, baseFontSize = 8.6, minHeight = 0.82, maxHeight = 1.75) {
  return rows.map((row) => {
    const estimatedLines = row.reduce((maxValue, cell, index) => {
      const fontSize = Math.max(8.2, cell.fontSize || baseFontSize);
      return Math.max(maxValue, estimateWrappedLines(cell.text, colWidths[index], fontSize));
    }, 1);
    return Math.max(minHeight, Math.min(maxHeight, 0.22 + estimatedLines * 0.22));
  });
}

function sumNumericValues(values = []) {
  return values.reduce((acc, value) => acc + Number(value || 0), 0);
}

function projectManualRows(rows, columnIndexes) {
  return rows.map((row) =>
    columnIndexes.map((columnIndex) => ({ ...row[columnIndex] }))
  );
}

function projectColumnWidths(colWidths, columnIndexes, targetWidth) {
  const selected = columnIndexes.map((columnIndex) => colWidths[columnIndex]);
  const selectedTotal = sumNumericValues(selected) || 1;
  const resolvedTargetWidth = targetWidth || selectedTotal;
  return selected.map((width) => resolvedTargetWidth * (width / selectedTotal));
}

function evaluateManualTableDensity(rows, colWidths, options = {}) {
  const {
    baseFontSize = 8.4,
    minRowHeight = 0.82,
    maxRowHeight = 1.35,
    headerHeight = 0.42,
    availableHeight = 4.2,
    maxLinesPerCell = 5.2,
    maxAverageLines = 3.1,
  } = options;

  const safeRows = rows.map((row) =>
    row.map((cell) => (typeof cell === 'object' ? cell : { text: cleanText(cell) }))
  );
  const rowHeights = computeManualRowHeights(safeRows, colWidths, baseFontSize, minRowHeight, maxRowHeight);
  const lineCounts = safeRows.flatMap((row) =>
    row.map((cell, index) => estimateWrappedLines(cell.text, colWidths[index], Math.max(8.2, cell.fontSize || baseFontSize)))
  );
  const totalHeight = headerHeight + sumNumericValues(rowHeights);
  const maxLines = lineCounts.length ? Math.max(...lineCounts) : 1;
  const avgLines = lineCounts.length ? sumNumericValues(lineCounts) / lineCounts.length : 1;
  const crampedCells = lineCounts.filter((value) => value >= Math.max(5, maxLinesPerCell)).length;
  const fits = totalHeight <= availableHeight && maxLines <= maxLinesPerCell && avgLines <= maxAverageLines;

  return {
    rowHeights,
    totalHeight,
    maxLines,
    avgLines,
    crampedCells,
    fits,
  };
}

function findFittingManualRowLimit(rows, colWidths, options = {}) {
  if (!rows.length) {
    return {
      rowLimit: 0,
      metrics: evaluateManualTableDensity([], colWidths, options),
    };
  }

  const preferredLimit = Math.min(rows.length, Math.max(1, options.preferredLimit || rows.length));
  const minLimit = Math.min(preferredLimit, Math.max(1, options.minLimit || 1));

  for (let limit = preferredLimit; limit >= minLimit; limit -= 1) {
    const metrics = evaluateManualTableDensity(rows.slice(0, limit), colWidths, options);
    if (metrics.fits) {
      return { rowLimit: limit, metrics };
    }
  }

  return {
    rowLimit: minLimit,
    metrics: evaluateManualTableDensity(rows.slice(0, minLimit), colWidths, options),
  };
}

function formatSectionPageNumber(baseNumber, pageIndex = 0) {
  if (!pageIndex) return baseNumber;
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const letter = alphabet[(pageIndex - 1) % alphabet.length];
  return `${baseNumber}${letter}`;
}

function loadOrganizerData() {
  const raw = execFileSync('python', ['-X', 'utf8', 'organizer.py', INPUT_FILE], {
    cwd: __dirname,
    encoding: 'utf8',
    env: { ...process.env, PYTHONUTF8: '1' },
  });
  const data = JSON.parse(raw);
  if (data.error) throw new Error(data.error);
  return data;
}

function buildDataset(sourceData) {
  const summary = sourceData.resumen_ejecutivo || {};
  const mainHeaders = sourceData.muestra_tabla?.encabezados || [];
  const mainRowsRaw = sourceData.muestra_tabla?.filas || [];
  const mainRows = mainRowsRaw.map((row) => ({
    id: row[0],
    solicitante: cleanText(row[1]),
    ciudad: cleanText(row[2]),
    valor: Number(row[3] || 0),
    estado: cleanText(row[4]),
    centro: cleanText(row[5]),
  }));

  const statusCounts = (sourceData.grafica_estados?.labels || []).map((label, index) => [
    cleanText(label),
    Number(sourceData.grafica_estados.valores[index] || 0),
  ]);

  const expenseSummary = (sourceData.grafica_valores?.labels || []).map((label, index) => ({
    label: cleanText(label),
    value: Number(sourceData.grafica_valores.valores[index] || 0),
  }));

  const topCities = (sourceData.grafica_ciudades?.labels || []).map((label, index) => [
    cleanText(label),
    Number(sourceData.grafica_ciudades.valores[index] || 0),
  ]);

  const topCostCenters = (sourceData.centros_costos?.labels || []).map((label, index) => [
    cleanText(label),
    Number(sourceData.centros_costos.valores[index] || 0),
  ]);

  const topRequesters = (sourceData.top_solicitantes?.labels || []).map((label, index) => ({
    name: cleanText(label),
    value: Number(sourceData.top_solicitantes.valores[index] || 0),
    count: Number(sourceData.top_solicitantes.conteos[index] || 0),
  }));

  const sampleRows = [...mainRows]
    .sort((a, b) => b.valor - a.valor)
    .slice(0, 10)
    .map((row) => ({
      id: row.id,
      solicitante: cleanText(row.solicitante, 28),
      ciudad: cleanText(row.ciudad, 28),
      valor: fmtCurrency(row.valor),
      estado: cleanText(row.estado, 24),
      centro: cleanText(row.centro, 14),
    }));

  const otherTables = sourceData.otras_tablas || {};
  const opportunityKey = Object.keys(otherTables).find((key) => key.toLowerCase().includes('opo')) || Object.keys(otherTables).find((key) => key.toLowerCase().includes('consolidado'));
  const hallazgosKey = Object.keys(otherTables).find((key) => key.toLowerCase().includes('hallazgo'));
  const opportunities = opportunityKey ? otherTables[opportunityKey] : null;
  const hallazgos = hallazgosKey ? otherTables[hallazgosKey] : null;

  const opportunityRows = opportunities
    ? opportunities.filas.map((row) => ({
        area: cleanText(row[0], 34) || 'Proceso transversal',
        opportunity: cleanText(row[1]),
        status: cleanText(row[2], 22),
        progress: normalizeProgress(row[3]),
        observation: cleanText(row[4]),
      }))
      .filter((row) => row.opportunity)
      .sort((a, b) => a.progress - b.progress || a.opportunity.localeCompare(b.opportunity))
    : [];

  const opportunityStatus = (() => {
    const map = new Map();
    opportunityRows.forEach((row) => {
      const key = row.status || 'Sin estado';
      map.set(key, (map.get(key) || 0) + 1);
    });
    return [...map.entries()].sort((a, b) => b[1] - a[1]);
  })();

  const opportunityMetrics = {
    total: opportunityRows.length,
    completed: opportunityRows.filter((row) => safeUpper(row.status).includes('COMPLET')).length,
    inProgress: opportunityRows.filter((row) => safeUpper(row.status).includes('EJECUC') || safeUpper(row.status).includes('PROCES')).length,
    pending: opportunityRows.filter((row) => safeUpper(row.status).includes('PENDIENT') || row.progress === 0).length,
    average: opportunityRows.length
      ? opportunityRows.reduce((acc, row) => acc + row.progress, 0) / opportunityRows.length
      : 0,
  };

  const hallazgoRows = hallazgos
    ? hallazgos.filas
      .map((row) => ({
        finding: cleanText(row[0]),
        risk: cleanText(row[1]),
        action: cleanText(row[3]),
        status: cleanText(row[4], 20),
      }))
      .filter((row) => row.finding)
    : [];

  const cosoRows = (sourceData.coso?.filas || []).map((row) => {
    const statusText = cleanText(row[2]);
    return {
      component: cleanText(row[0], 28),
      items: cleanText(row[1], 115),
      status: statusText,
      opportunities: (statusText.match(/oportunidad de mejora/gi) || []).length,
      effective: (statusText.match(/control efectivo/gi) || []).length,
      ineffective: (statusText.match(/control no efectivo/gi) || []).length,
    };
  });

  const contabilizado = statusCounts.find(([label]) => safeUpper(label).includes('CONTABILIZ'));
  const rechazo = statusCounts.find(([label]) => safeUpper(label).includes('RECHAZ'));
  const topRequester = topRequesters[0];
  const topCity = topCities[0];
  const topCenter = topCostCenters[0];
  const topExpense = [...expenseSummary].sort((a, b) => b.value - a.value)[0];

  return {
    generatedAt: new Date(),
    mainSheetName: cleanText(sourceData.metadatos?.archivo) || mainHeaders.join(', '),
    kpis: {
      totalRows: Number(summary.total_comisiones || mainRows.length || 0),
      totalValue: Number(summary.valor_total || 0),
      uniqueRequesters: Number(summary.unique_solicitantes || 0),
      uniqueCities: Number(summary.unique_ciudades || 0),
      uniqueCenters: Number(summary.unique_centros || 0),
      avgValue: Number(summary.promedio_comision || 0),
      maxValue: Number(summary.valor_max_comision || 0),
    },
    charts: {
      statusCounts,
      topCities,
      topCostCenters,
      topRequesters,
      expenseSummary,
    },
    tables: {
      sampleRows,
      priorityOpportunities: opportunityRows.slice(0, 8),
      hallazgos: hallazgoRows.slice(0, 6),
      coso: cosoRows,
    },
    opportunityStatus,
    opportunityMetrics,
    insights: [
      contabilizado
        ? `${cleanText(contabilizado[0])} concentra ${fmtPercent(ratio(contabilizado[1], Number(summary.total_comisiones || 0)))} de las comisiones registradas.`
        : null,
      topExpense
        ? `${topExpense.label} es el rubro de gasto más alto con ${fmtCompactCurrency(topExpense.value)}.`
        : null,
      topRequester
        ? `${cleanText(topRequester.name, 30)} lidera el valor gestionado con ${fmtCompactCurrency(topRequester.value)}.`
        : null,
      topCity
        ? `${cleanText(topCity[0], 30)} aparece como el destino más frecuente con ${fmtNumber(topCity[1])} registros.`
        : null,
      rechazo
        ? `Los estados de rechazo reúnen ${fmtNumber(rechazo[1])} casos y requieren control preventivo focalizado.`
        : null,
      topCenter
        ? `El centro de costos ${cleanText(topCenter[0])} concentra ${fmtCompactCurrency(topCenter[1])} del valor analizado.`
        : null,
    ].filter(Boolean),
  };
}

function hasMeaningfulCosoContent(dataset) {
  return dataset.tables.coso.some((item) =>
    cleanText(item.component)
    && (
      Number(item.opportunities || 0) > 0
      || Number(item.effective || 0) > 0
      || Number(item.ineffective || 0) > 0
      || cleanText(item.items, 80)
    )
  );
}

function hasMeaningfulFindings(dataset) {
  return dataset.tables.hallazgos.some((item) =>
    cleanText(item.finding, 80) && (cleanText(item.risk, 80) || cleanText(item.action, 80))
  );
}

function hasMeaningfulOpportunities(dataset) {
  return dataset.opportunityMetrics.total > 0
    && dataset.tables.priorityOpportunities.length > 0
    && dataset.opportunityStatus.length > 0;
}

function isSpecializedCommissionsDataset(sourceData) {
  const headers = sourceData?.muestra_tabla?.encabezados || [];
  const normalized = headers.map((header) =>
    cleanText(header).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase()
  );
  return normalized.includes('solicitante') && normalized.some((header) => header.includes('valor total'));
}

function extractStructuredSheetModel(workbook, sheetName) {
  return extractStructuredSheetModels(workbook, sheetName)[0] || null;
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (value instanceof Date) return null;

  let text = String(value).trim();
  if (!text) return null;
  const hasPercent = text.includes('%');
  const isNegativeByParens = /^\(.*\)$/.test(text);
  text = text.replace(/[^\d,.\-]/g, '');
  if (!text) return null;

  const lastComma = text.lastIndexOf(',');
  const lastDot = text.lastIndexOf('.');
  if (lastComma > -1 && lastDot > -1) {
    if (lastComma > lastDot) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (lastComma > -1) {
    const decimalLength = text.length - lastComma - 1;
    text = decimalLength <= 2 ? text.replace(',', '.') : text.replace(/,/g, '');
  }

  const num = Number(text);
  if (!Number.isFinite(num)) return null;
  const signed = isNegativeByParens ? -Math.abs(num) : num;
  return hasPercent ? signed / 100 : signed;
}

function toDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'string') {
    const text = cleanText(value);
    if (!text) return null;
    const iso = new Date(text);
    if (!Number.isNaN(iso.getTime())) return iso;
    const match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (match) {
      const [, dd, mm, yyyy] = match;
      const year = yyyy.length === 2 ? `20${yyyy}` : yyyy;
      const parsed = new Date(`${year}-${mm.padStart(2, '0')}-${dd.padStart(2, '0')}`);
      if (!Number.isNaN(parsed.getTime())) return parsed;
    }
  }
  return null;
}

function formatGenericCell(value, type) {
  if (value === null || value === undefined || value === '') return '—';
  if (type === 'currency') {
    const num = toNumber(value);
    return num === null ? cleanText(value, 40) : fmtCurrency(num);
  }
  if (type === 'percent') {
    const num = toNumber(value);
    return num === null ? cleanText(value, 40) : fmtPercent(num, 1);
  }
  if (type === 'number') {
    const num = toNumber(value);
    return num === null ? cleanText(value, 40) : fmtNumber(num);
  }
  if (type === 'date') {
    const date = toDateValue(value);
    return date ? date.toLocaleDateString('es-CO') : cleanText(value, 40);
  }
  return cleanText(value, 40);
}

function inferColumnProfile(header, values) {
  const nonEmpty = values.filter((value) => !(value === null || value === undefined || cleanText(value) === ''));
  if (!nonEmpty.length) {
    return { header, type: 'text', count: 0, uniqueCount: 0, total: 0, avg: 0 };
  }

  const numericValues = nonEmpty.map(toNumber).filter((value) => value !== null);
  const dateValues = nonEmpty.map(toDateValue).filter((value) => value);
  const numericRatio = numericValues.length / nonEmpty.length;
  const dateRatio = dateValues.length / nonEmpty.length;
  const normalizedHeader = normalizeHeaderKey(header);
  const uniqueCount = new Set(nonEmpty.map((value) => cleanText(value))).size;
  const idLikeHeader = /(^id\b|^id_|codigo|cod\b|folio|consecutivo|radicado|ticket|numero de|nro\b|comision\b)/.test(normalizedHeader);

  let type = 'text';
  if (idLikeHeader && numericRatio >= 0.7 && uniqueCount >= Math.max(4, Math.floor(nonEmpty.length * 0.6))) type = 'id';
  if (numericRatio >= 0.7) type = 'number';
  if (dateRatio >= 0.7) type = 'date';
  if (type === 'number' && /%|porcentaje|ratio|tasa|participacion/.test(normalizedHeader)) type = 'percent';
  if (type === 'number' && /valor|venta|ingreso|costo|gasto|precio|importe|monto|total|facturacion|presupuesto/.test(normalizedHeader)) type = 'currency';
  if (idLikeHeader && type === 'number' && uniqueCount >= Math.max(4, Math.floor(nonEmpty.length * 0.6))) type = 'id';

  return {
    header,
    type,
    count: nonEmpty.length,
    uniqueCount,
    total: numericValues.reduce((acc, value) => acc + value, 0),
    avg: numericValues.length ? numericValues.reduce((acc, value) => acc + value, 0) / numericValues.length : 0,
    max: numericValues.length ? Math.max(...numericValues) : 0,
    sample: nonEmpty.slice(0, 3).map((value) => cleanText(value, 24)),
  };
}

function groupAndAggregate(rows, categoryHeader, valueHeader) {
  const grouped = new Map();
  rows.forEach((row) => {
    const key = cleanText(row[categoryHeader]);
    if (!key || isGhostText(key)) return;
    const amount = valueHeader ? toNumber(row[valueHeader]) : 1;
    grouped.set(key, (grouped.get(key) || 0) + (amount === null ? 0 : amount));
  });
  return [...grouped.entries()].sort((a, b) => b[1] - a[1]);
}

function sortChartEntries(labels, values) {
  const monthIndexes = labels.map((label) => monthSortIndex(label));
  if (monthIndexes.every((value) => value !== null)) {
    return labels
      .map((label, index) => ({ label, value: values[index], order: monthIndexes[index] }))
      .sort((a, b) => a.order - b.order);
  }
  return labels.map((label, index) => ({ label, value: values[index], order: index }));
}

function findProfileByPattern(profiles, pattern, allowedTypes = null) {
  return profiles.find((profile) => {
    if (allowedTypes && !allowedTypes.includes(profile.type)) return false;
    return pattern.test(normalizeHeaderKey(profile.header));
  }) || null;
}

function describeMetricValue(metric, value) {
  if (!metric) return fmtNumber(value);
  if (metric.type === 'currency') return fmtCompactCurrency(value);
  if (metric.type === 'percent') return fmtPercent(value);
  return fmtNumber(value);
}

function inferSheetDomain(sheet, profiles) {
  const signalText = [
    sheet.name,
    ...profiles.map((profile) => profile.header),
  ].map((value) => normalizeHeaderKey(value)).join(' ');

  const scores = {
    sales: 0,
    inventory: 0,
    hr: 0,
    finance: 0,
    operations: 0,
  };

  if (/venta|facturacion|ingreso|cliente|comercial|pedido|canal|ticket|recaudo/.test(signalText)) scores.sales += 3;
  if (/stock|inventario|agotado|existencia|sku|bodega/.test(signalText)) scores.inventory += 5;
  if (/producto|referencia|categoria/.test(signalText)) scores.inventory += 1;
  if (/empleado|colaborador|nomina|cargo|area|sede|equipo|rrhh|talento/.test(signalText)) scores.hr += 3;
  if (/presupuesto|costo|gasto|margen|utilidad|ebitda|financ|cuenta|egreso/.test(signalText)) scores.finance += 3;
  if (/estado|avance|cumplimiento|pendiente|proceso|operacion|tiempo|turno|servicio/.test(signalText)) scores.operations += 2;

  const topDomain = Object.entries(scores).sort((a, b) => b[1] - a[1])[0];
  return topDomain && topDomain[1] > 0 ? topDomain[0] : 'general';
}

function pickExecutiveSignal(sheet) {
  const insights = Array.isArray(sheet?.insights) ? sheet.insights.filter(Boolean) : [];
  const humanSignal = insights.find((line) => !/^\d{1,2}\/\d{4}\b/.test(cleanText(line)));
  return humanSignal || insights[0] || `La hoja ${sheet.name} aporta evidencia util para la lectura ejecutiva.`;
}

function buildMetricPairInsight(sheet, profiles) {
  const budgetProfile = findProfileByPattern(profiles, /presupuesto|meta|objetivo|plan/, ['number', 'currency', 'percent']);
  const actualProfile = findProfileByPattern(profiles, /ejecutado|real|avance|cumplido|venta|ingreso|facturacion/, ['number', 'currency', 'percent']);
  if (budgetProfile && actualProfile && budgetProfile.header !== actualProfile.header) {
    const budgetTotal = Math.abs(budgetProfile.total || 0);
    const actualTotal = Math.abs(actualProfile.total || 0);
    if (budgetTotal > 0 && actualTotal > 0) {
      const execution = actualTotal / budgetTotal;
      const variance = actualTotal - budgetTotal;
      return `La ejecucion de ${humanizeLabel(actualProfile.header).toLowerCase()} alcanza ${fmtPercent(execution)} frente a ${humanizeLabel(budgetProfile.header).toLowerCase()}, con una variacion de ${describeMetricValue(actualProfile, variance)}.`;
    }
  }

  const revenueProfile = findProfileByPattern(profiles, /venta|ingreso|facturacion|recaudo|cobro/, ['number', 'currency']);
  const costProfile = findProfileByPattern(profiles, /costo|gasto|egreso|compra/, ['number', 'currency']);
  if (revenueProfile && costProfile && revenueProfile.header !== costProfile.header) {
    const revenueTotal = Math.abs(revenueProfile.total || 0);
    const costTotal = Math.abs(costProfile.total || 0);
    if (revenueTotal > 0 && costTotal > 0) {
      return `Los costos visibles equivalen a ${fmtPercent(costTotal / revenueTotal)} de ${humanizeLabel(revenueProfile.header).toLowerCase()}, dejando una brecha bruta cercana a ${fmtCompactCurrency(revenueTotal - costTotal)}.`;
    }
  }

  return null;
}

function buildDomainSpecificInsight(sheet, profiles, primaryMetric, domain, statusProfile, categoryProfile) {
  if (domain === 'inventory') {
    const stockMetric = findProfileByPattern(profiles, /stock|inventario|existencia|unidades|cantidad/, ['number', 'currency']);
    if (stockMetric) {
      const zeroOrAlertRows = sheet.rows.filter((row) => {
        const amount = toNumber(row[stockMetric.header]);
        return amount !== null && amount <= 0;
      }).length;
      if (zeroOrAlertRows > 0) {
        return `${fmtNumber(zeroOrAlertRows)} referencias aparecen sin disponibilidad en ${humanizeLabel(stockMetric.header).toLowerCase()}, senal directa de reposicion prioritaria.`;
      }
      if (stockMetric.max > 0) {
        return `${humanizeLabel(stockMetric.header)} alcanza un maximo de ${fmtNumber(stockMetric.max)} y conviene balancear cobertura frente a referencias con menor respaldo.`;
      }
    }
    if (statusProfile) {
      const positiveEntry = groupAndAggregate(sheet.rows, statusProfile.header, null)
        .find(([label]) => /disponible|activo|ok|vigente|en stock/.test(normalizeHeaderKey(label)));
      if (positiveEntry) {
        return `${humanizeLabel(positiveEntry[0])} sigue siendo el estado dominante dentro del inventario revisado, pero requiere vigilancia junto a los estados de alerta.`;
      }
    }
  }

  if (domain === 'sales') {
    const channelProfile = findProfileByPattern(profiles, /canal|cliente|vendedor|asesor|producto|categoria|linea|ciudad/, ['text']);
    const salesMetric = primaryMetric || findProfileByPattern(profiles, /venta|ingreso|facturacion|recaudo/, ['number', 'currency']);
    if (channelProfile && salesMetric) {
      const grouped = groupAndAggregate(sheet.rows, channelProfile.header, salesMetric.header).slice(0, 3);
      if (grouped.length >= 2) {
        return `${humanizeLabel(grouped[0][0])} encabeza ${humanizeLabel(salesMetric.header).toLowerCase()} con ${describeMetricValue(salesMetric, grouped[0][1])}, por encima de ${humanizeLabel(grouped[1][0])}.`;
      }
    }
  }

  if (domain === 'hr') {
    const peopleProfile = findProfileByPattern(profiles, /empleado|colaborador|persona|documento/, ['text']);
    const areaProfile = findProfileByPattern(profiles, /area|cargo|sede|equipo|gerencia/, ['text']);
    const payrollMetric = findProfileByPattern(profiles, /nomina|salario|costo|ingreso/, ['number', 'currency']);
    if (areaProfile && payrollMetric) {
      const grouped = groupAndAggregate(sheet.rows, areaProfile.header, payrollMetric.header).slice(0, 3);
      if (grouped.length) {
        return `${humanizeLabel(grouped[0][0])} concentra la mayor carga economica del talento con ${describeMetricValue(payrollMetric, grouped[0][1])}.`;
      }
    }
    if (areaProfile && peopleProfile) {
      const grouped = groupAndAggregate(sheet.rows, areaProfile.header, null).slice(0, 3);
      if (grouped.length) {
        return `${humanizeLabel(grouped[0][0])} concentra la mayor dotacion visible con ${fmtNumber(grouped[0][1])} registros de personal.`;
      }
    }
  }

  if (domain === 'finance') {
    const pairInsight = buildMetricPairInsight(sheet, profiles);
    if (pairInsight) return pairInsight;
    const financeMetric = primaryMetric || findProfileByPattern(profiles, /costo|gasto|presupuesto|utilidad|margen/, ['number', 'currency', 'percent']);
    if (financeMetric) {
      return `${humanizeLabel(financeMetric.header)} define la conversacion financiera principal, con acumulado de ${describeMetricValue(financeMetric, financeMetric.total)} y maximos de ${describeMetricValue(financeMetric, financeMetric.max)}.`;
    }
  }

  if (domain === 'operations') {
    if (statusProfile) {
      const entries = groupAndAggregate(sheet.rows, statusProfile.header, null).slice(0, 3);
      if (entries.length >= 2) {
        return `${humanizeLabel(entries[0][0])} es el estado operativo dominante, mientras ${humanizeLabel(entries[1][0])} concentra el siguiente volumen de seguimiento.`;
      }
    }
    if (categoryProfile && primaryMetric) {
      const grouped = groupAndAggregate(sheet.rows, categoryProfile.header, primaryMetric.header).slice(0, 3);
      if (grouped.length) {
        return `${humanizeLabel(grouped[0][0])} lidera el frente operativo medido por ${humanizeLabel(primaryMetric.header).toLowerCase()} con ${describeMetricValue(primaryMetric, grouped[0][1])}.`;
      }
    }
  }

  return buildMetricPairInsight(sheet, profiles);
}

function buildBusinessFocusLabel(sheet) {
  switch (sheet.domain) {
    case 'sales':
      return 'Ventas y mezcla comercial';
    case 'inventory':
      return 'Inventario y disponibilidad';
    case 'hr':
      return 'Talento y distribucion';
    case 'finance':
      return 'Finanzas y ejecucion';
    case 'operations':
      return 'Operacion y cumplimiento';
    default:
      return sheet.primaryMetric ? `Seguimiento de ${humanizeLabel(sheet.primaryMetric.header)}` : 'Lectura ejecutiva';
  }
}

function buildWorkbookPriorityRows(sheets) {
  return sheets.slice(0, 6).map((sheet) => ({
    sheet: sheet.blockTitle || sheet.name,
    focus: buildBusinessFocusLabel(sheet),
    metric: sheet.headlineMetrics[0] ? `${sheet.headlineMetrics[0].label}: ${sheet.headlineMetrics[0].value}` : 'Indicador principal no identificado',
    signal: pickExecutiveSignal(sheet),
  }));
}

function buildSheetChart(sheet, profiles) {
  const numericProfiles = profiles.filter((profile) => ['number', 'currency', 'percent'].includes(profile.type));
  const categoryProfiles = profiles.filter((profile) => profile.type === 'text' && profile.uniqueCount >= 2 && profile.uniqueCount <= Math.min(12, Math.max(6, Math.floor(sheet.rows.length * 0.6))));
  const dateProfiles = profiles.filter((profile) => profile.type === 'date');

  if (dateProfiles.length && numericProfiles.length) {
    const dateHeader = dateProfiles[0].header;
    const valueHeader = numericProfiles[0].header;
    const grouped = new Map();
    sheet.rows.forEach((row) => {
      const date = toDateValue(row[dateHeader]);
      const amount = toNumber(row[valueHeader]);
      if (!date || amount === null) return;
      const key = `${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
      grouped.set(key, (grouped.get(key) || 0) + amount);
    });
    const entries = sanitizeChartEntries([...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0])).slice(0, 8));
    if (entries.length >= 2) {
      return {
        title: `${cleanText(valueHeader, 22)} por periodo`,
        labels: entries.map(([label]) => label),
        values: entries.map(([, value]) => value),
      };
    }
  }

  if (categoryProfiles.length && numericProfiles.length) {
    const preferredCategory =
      categoryProfiles.find((profile) => /estado|categoria|tipo|canal|sede|area|ciudad|region|mes|semana|producto|linea/.test(normalizeHeaderKey(profile.header))) ||
      categoryProfiles[0];
    const preferredMetric =
      numericProfiles.find((profile) => profile.type === 'currency') ||
      numericProfiles.find((profile) => profile.type === 'number') ||
      numericProfiles[0];
    const categoryHeader = preferredCategory.header;
    const valueHeader = preferredMetric.header;
    let entries = sanitizeChartEntries(groupAndAggregate(sheet.rows, categoryHeader, valueHeader).slice(0, 6));
    const sortedEntries = sortChartEntries(entries.map(([label]) => label), entries.map(([, value]) => value));
    entries = sanitizeChartEntries(sortedEntries.map((entry) => [entry.label, entry.value]));
    if (entries.length >= 2 && hasMeaningfulChartEntries(entries, preferredMetric.type)) {
      return {
        title: `${humanizeLabel(valueHeader, 22)} por ${humanizeLabel(categoryHeader, 18).toLowerCase()}`,
        labels: entries.map(([label]) => humanizeLabel(label, 26)),
        values: entries.map(([, value]) => value),
        seriesType: preferredMetric.type,
      };
    }
  }

  if (categoryProfiles.length) {
    const preferredCategory =
      categoryProfiles.find((profile) => /estado|categoria|tipo|canal|sede|area|ciudad|region|mes|semana|producto|linea/.test(normalizeHeaderKey(profile.header))) ||
      categoryProfiles[0];
    const categoryHeader = preferredCategory.header;
    let entries = sanitizeChartEntries(groupAndAggregate(sheet.rows, categoryHeader, null).slice(0, 6));
    const sortedEntries = sortChartEntries(entries.map(([label]) => label), entries.map(([, value]) => value));
    entries = sanitizeChartEntries(sortedEntries.map((entry) => [entry.label, entry.value]));
    if (entries.length >= 2 && hasMeaningfulChartEntries(entries, 'count')) {
      return {
        title: `Distribucion por ${humanizeLabel(categoryHeader, 22).toLowerCase()}`,
        labels: entries.map(([label]) => humanizeLabel(label, 26)),
        values: entries.map(([, value]) => value),
        seriesType: 'count',
      };
    }
  }

  return null;
}

function averageTextLengthForSheet(sheet) {
  const textValues = sheet.rows
    .flatMap((row) => Object.values(row))
    .map((value) => cleanText(value))
    .filter(Boolean);
  if (!textValues.length) return 0;
  return textValues.reduce((acc, value) => acc + value.length, 0) / textValues.length;
}

function buildPrimaryTablePreview(sheet, profiles, dominantCategoryEntries, primaryMetric, dominantCategoryProfile) {
  if (dominantCategoryEntries.length && primaryMetric) {
    const metricProfile = profiles.find((profile) => profile.header === primaryMetric.header);
    const rows = dominantCategoryEntries.slice(0, 5)
      .filter(([label, value]) => !isGhostText(label) && Number.isFinite(Number(value)))
      .map(([label, value], index) => [
        {
          text: humanizeLabel(label, 26),
          fill: index % 2 === 0 ? 'F8FBF4' : COLORS.white,
          fontSize: 7.7,
        },
        {
          text: formatGenericCell(value, metricProfile?.type || primaryMetric.type || 'number'),
          align: 'right',
          fill: index % 2 === 0 ? 'F8FBF4' : COLORS.white,
          fontSize: 7.7,
        },
      ]);
    return {
      title: `${humanizeLabel(primaryMetric.header, 22)} por foco`,
      headers: [humanizeLabel(dominantCategoryProfile?.header || 'Categoria', 18), humanizeLabel(primaryMetric.header, 18)],
      rows,
    };
  }

  const headers = selectBestTableHeaders(sheet, profiles, 4);
  const fallbackHeaders = headers.length ? headers : sheet.headers.slice(0, Math.min(4, sheet.headers.length));
  const cleanedPreview = buildCleanTablePreview(sheet, fallbackHeaders, profiles, 5);
  return {
    title: 'Resumen tabular',
    headers: cleanedPreview.headers.map((header) => humanizeLabel(header, 18)),
    rows: cleanedPreview.rows.map((row, rowIndex) =>
      row.map((value) => ({
        text: cleanText(value) || 'N/D',
        fill: rowIndex % 2 === 0 ? 'F8FBF4' : COLORS.white,
        fontSize: 7.3,
      }))
    ),
  };
}

function choosePrimaryVisual(sheet, profiles, chart, dominantCategoryEntries, primaryMetric, dominantCategoryProfile) {
  const avgTextLength = averageTextLengthForSheet(sheet);
  const textHeavyColumns = profiles.filter((profile) => profile.type === 'text' && profile.sample.some((sample) => sample.length >= 24)).length;
  const chartCandidateStrength = chart
    ? (chart.labels.length >= 2 && chart.labels.length <= 8 ? 2 : 0) + (chart.seriesType === 'currency' || chart.seriesType === 'number' ? 2 : 1)
    : 0;
  const tableComplexity = (sheet.headers.length >= 6 ? 2 : 0) + (avgTextLength >= 18 ? 2 : 0) + (textHeavyColumns >= 2 ? 2 : 0);
  const tablePreview = buildPrimaryTablePreview(sheet, profiles, dominantCategoryEntries, primaryMetric, dominantCategoryProfile);
  const previewColWidths = Array(Math.max(1, tablePreview.headers.length)).fill(Math.max(1.4, 6 / Math.max(1, tablePreview.headers.length)));
  const previewDensity = evaluateManualTableDensity(tablePreview.rows, previewColWidths, {
    baseFontSize: 8.2,
    minRowHeight: 0.54,
    maxRowHeight: 0.9,
    headerHeight: 0.34,
    availableHeight: 2.1,
    maxLinesPerCell: 4.8,
    maxAverageLines: 2.85,
  });

  if (PRESENTATION_VISUAL_MODE === 'charts') {
    return chart && chartCandidateStrength >= 3
      ? { type: 'chart', title: chart.title }
      : null;
  }

  if (PRESENTATION_VISUAL_MODE === 'tables') {
    return {
      type: 'table',
      title: tablePreview.title,
      headers: tablePreview.headers,
      rows: tablePreview.rows,
    };
  }

  if (chart && chartCandidateStrength >= 3 && (tableComplexity <= 3 || !previewDensity.fits || previewDensity.crampedCells >= 3)) {
    return { type: 'chart', title: chart.title };
  }

  return {
    type: 'table',
    title: tablePreview.title,
    headers: tablePreview.headers,
    rows: tablePreview.rows,
  };
}

function buildBusinessInsightsForSheet(sheet, profiles, chart, primaryMetric) {
  const insights = [];
  const textProfiles = profiles.filter((profile) => profile.type === 'text' && profile.uniqueCount >= 2);
  const categoryProfile = textProfiles.find((profile) => /estado|categoria|tipo|canal|sede|area|ciudad|region|mes|semana|producto|linea/.test(normalizeHeaderKey(profile.header))) || textProfiles[0];
  const statusProfile = textProfiles.find((profile) => /estado|status|situacion|stock|agotado|cumplimiento|avance/.test(normalizeHeaderKey(profile.header)));
  const domain = inferSheetDomain(sheet, profiles);

  if (statusProfile) {
    const statusEntries = groupAndAggregate(sheet.rows, statusProfile.header, null).slice(0, 4);
    const totalStatuses = statusEntries.reduce((acc, [, value]) => acc + value, 0);
    const alertEntry = statusEntries.find(([label]) => /agotado|pendiente|rechaz|vencid|stock bajo|critico|alerta|inactivo|bajo/.test(normalizeHeaderKey(label)));
    if (alertEntry && totalStatuses > 0) {
      insights.push(`${humanizeLabel(alertEntry[0])} representa ${fmtPercent(alertEntry[1] / totalStatuses)} del universo principal en ${humanizeLabel(statusProfile.header).toLowerCase()}.`);
    }
  }

  if (chart && chart.labels.length >= 2) {
    const total = chart.values.reduce((acc, value) => acc + value, 0);
    const topValue = chart.values[0] || 0;
    const topLabel = chart.labels[0];
    const share = total ? topValue / total : 0;
    insights.push(primaryMetric?.type === 'currency'
      ? `${topLabel} lidera ${humanizeLabel(primaryMetric.header).toLowerCase()} con ${fmtCompactCurrency(topValue)} y una participación cercana a ${fmtPercent(share)}.`
      : `${topLabel} concentra el mayor peso en ${humanizeLabel(chart.title).toLowerCase()} con una participación cercana a ${fmtPercent(share)}.`);

    if (chart.values.length >= 3) {
      const maxValue = Math.max(...chart.values);
      const minValue = Math.min(...chart.values);
      if (maxValue > 0) insights.push(`La brecha entre el mejor y el menor resultado visible es de ${fmtPercent((maxValue - minValue) / maxValue)}.`);
    }

    const monthIndexes = chart.labels.map((label) => monthSortIndex(label));
    if (monthIndexes.every((value) => value !== null) && chart.values.length >= 2) {
      const firstValue = chart.values[0];
      const lastValue = chart.values[chart.values.length - 1];
      const peakIndex = chart.values.indexOf(Math.max(...chart.values));
      if (firstValue > 0) insights.push(`${chart.labels[peakIndex]} marca el punto más alto del periodo y el cierre muestra una variación de ${fmtPercent((lastValue - firstValue) / firstValue)} frente al inicio.`);
    }
  }

  const pairInsight = buildMetricPairInsight(sheet, profiles);
  if (pairInsight) insights.push(pairInsight);

  if (categoryProfile && primaryMetric) {
    const grouped = groupAndAggregate(sheet.rows, categoryProfile.header, primaryMetric.header).slice(0, 5);
    if (grouped.length >= 2) {
      const total = grouped.reduce((acc, [, value]) => acc + value, 0);
      const [label, value] = grouped[0];
      insights.push(primaryMetric.type === 'currency'
        ? `${label} sobresale en ${humanizeLabel(primaryMetric.header).toLowerCase()} con ${fmtCompactCurrency(value)} sobre los grupos principales.`
        : `${label} sobresale en ${humanizeLabel(primaryMetric.header).toLowerCase()} con ${fmtNumber(value)} dentro de los grupos principales.`);
      if (total > 0) insights.push(`${humanizeLabel(categoryProfile.header)} muestra concentración visible en los primeros grupos, con ${fmtPercent(value / total)} para el líder.`);
    }
  }

  const domainInsight = buildDomainSpecificInsight(sheet, profiles, primaryMetric, domain, statusProfile, categoryProfile);
  if (domainInsight) insights.push(domainInsight);

  if (primaryMetric && primaryMetric.type === 'percent') {
    insights.push(`El indicador ${humanizeLabel(primaryMetric.header).toLowerCase()} promedia ${fmtPercent(primaryMetric.avg)} y alcanza máximos de ${fmtPercent(primaryMetric.max)}.`);
  } else if (primaryMetric && primaryMetric.type === 'currency') {
    insights.push(`${humanizeLabel(primaryMetric.header)} acumula ${fmtCompactCurrency(primaryMetric.total)} y alcanza picos de ${fmtCompactCurrency(primaryMetric.max)}.`);
  } else if (primaryMetric && primaryMetric.type === 'number') {
    insights.push(`${humanizeLabel(primaryMetric.header)} alcanza un máximo de ${fmtNumber(primaryMetric.max)} y mantiene un promedio de ${fmtNumber(primaryMetric.avg)}.`);
  }

  return [...new Set(insights.filter(Boolean))].slice(0, 4);
}

function summarizeGenericSheet(sheet) {
  const profiles = sheet.headers.map((header) => inferColumnProfile(header, sheet.rows.map((row) => row[header])));
  const numericProfiles = profiles.filter((profile) => ['number', 'currency', 'percent'].includes(profile.type));
  const categoryProfiles = profiles.filter((profile) => profile.type === 'text' && profile.uniqueCount >= 2);
  const dateProfiles = profiles.filter((profile) => profile.type === 'date');
  const primaryMetric = [...numericProfiles].sort((a, b) => Math.abs(b.total) - Math.abs(a.total))[0] || null;
  const secondaryMetric = [...numericProfiles].sort((a, b) => Math.abs(b.avg) - Math.abs(a.avg))[1] || numericProfiles[1] || null;
  const chart = buildSheetChart(sheet, profiles);
  const domain = inferSheetDomain(sheet, profiles);
  const dominantCategoryProfile = categoryProfiles.find((profile) => /estado|categoria|tipo|canal|sede|area|ciudad|region|mes|semana|producto|linea/.test(normalizeHeaderKey(profile.header))) || categoryProfiles[0] || null;
  const dominantCategoryEntries = dominantCategoryProfile
    ? groupAndAggregate(sheet.rows, dominantCategoryProfile.header, primaryMetric?.header).slice(0, 5)
    : [];
  const previewHeaders = selectBestTableHeaders(sheet, profiles, 4);
  const previewFallbackHeaders = previewHeaders.length ? previewHeaders : sheet.headers.slice(0, Math.min(4, sheet.headers.length));
  const cleanedPreview = buildCleanTablePreview(sheet, previewFallbackHeaders, profiles, 6);
  const groupedPreview = dominantCategoryEntries.length && primaryMetric
    ? {
        headers: [
          humanizeLabel(dominantCategoryProfile?.header || 'Categoria', 18),
          humanizeLabel(primaryMetric.header, 18),
        ],
        rows: dominantCategoryEntries.slice(0, 5).map(([label, value]) => [
          humanizeLabel(label, 22) || 'N/D',
          formatGenericCell(value, primaryMetric.type || 'number') || 'N/D',
        ]),
      }
    : { headers: [], rows: [] };
  const useCleanPreview = cleanedPreview.headers.length >= 2 && cleanedPreview.rows.length >= 2;
  const tableHeaders = useCleanPreview
    ? cleanedPreview.headers.map((header) => humanizeLabel(header, 18))
    : groupedPreview.headers;
  const tableRows = (useCleanPreview ? cleanedPreview.rows : groupedPreview.rows).map((row) =>
    row.map((value) => cleanText(value) || 'N/D')
  );
  const tableCompleteness = tableRows.length
    ? tableRows.reduce((acc, row) => acc + (countRenderablePreviewCells(row) / Math.max(1, row.length)), 0) / tableRows.length
    : 0;
  const businessInsights = buildBusinessInsightsForSheet(sheet, profiles, chart, primaryMetric);
  const primaryVisual = choosePrimaryVisual(sheet, profiles, chart, dominantCategoryEntries, primaryMetric, dominantCategoryProfile);
  const contextParts = [
    sheet.title ? `Titulo detectado: ${cleanText(sheet.title, 72)}.` : null,
    sheet.subtitle ? `Subtitulo detectado: ${cleanText(sheet.subtitle, 90)}.` : null,
    sheet.notes?.[0] ? `Nota detectada: ${cleanText(sheet.notes[0], 100)}.` : null,
  ].filter(Boolean);
  const headlineMetrics = [
    primaryMetric
      ? {
          label: humanizeLabel(primaryMetric.header),
          value: primaryMetric.type === 'currency'
            ? fmtCompactCurrency(primaryMetric.total)
            : primaryMetric.type === 'percent'
              ? fmtPercent(primaryMetric.avg)
              : fmtNumber(primaryMetric.total || primaryMetric.max),
          note: primaryMetric.type === 'currency' ? 'Indicador principal acumulado' : 'Indicador principal identificado',
        }
      : null,
    secondaryMetric
      ? {
          label: humanizeLabel(secondaryMetric.header),
          value: secondaryMetric.type === 'currency'
            ? fmtCompactCurrency(secondaryMetric.avg)
            : secondaryMetric.type === 'percent'
              ? fmtPercent(secondaryMetric.avg)
              : fmtNumber(secondaryMetric.avg || secondaryMetric.max),
          note: secondaryMetric.type === 'currency' ? 'Promedio del segundo indicador' : 'Promedio / referencia',
        }
      : null,
    dominantCategoryEntries.length
      ? {
          label: humanizeLabel(dominantCategoryProfile.header),
          value: humanizeLabel(dominantCategoryEntries[0][0], 22),
          note: primaryMetric ? 'Mayor peso dentro de la hoja' : 'Categoria más representativa',
        }
      : null,
  ].filter(Boolean);
  const layoutProfile = buildGenericLayoutProfile({
    insights: businessInsights,
    notes: sheet.notes || [],
    contextSummary: contextParts.join(' '),
    headlineMetrics,
    chart,
    primaryVisual,
    rowCount: sheet.rows.length,
    colCount: sheet.headers.length,
    numericCount: numericProfiles.length,
    categoryCount: categoryProfiles.length,
    dateCount: dateProfiles.length,
  });

  return {
    name: cleanText(sheet.displayName || humanizeLabel(sheet.name, 36), 42),
    rawName: cleanText(sheet.name, 42),
    sourceSheetName: sheet.sourceSheetName || sheet.name,
    blockTitle: cleanText(sheet.title, 90),
    blockSubtitle: cleanText(sheet.subtitle, 120),
    blockNotes: sheet.notes || [],
    blockCount: sheet.blockCount || 1,
    blockIndex: sheet.blockIndex || 0,
    domain,
    rowCount: sheet.rows.length,
    colCount: sheet.headers.length,
    numericCount: numericProfiles.length,
    categoryCount: categoryProfiles.length,
    dateCount: dateProfiles.length,
    primaryMetric,
    secondaryMetric,
    chart,
    dominantCategoryProfile,
    dominantCategoryEntries,
    headlineMetrics,
    primaryVisual,
    layoutProfile,
    contextSummary: cleanText(contextParts.join(' '), 135),
    tableHeaders,
    tableRows,
    tableCompleteness,
    insights: businessInsights.length
      ? businessInsights
      : [
          primaryMetric ? `${humanizeLabel(primaryMetric.header)} es el indicador con mayor peso dentro de la hoja.` : null,
          chart ? `${chart.labels[0]} aparece como el frente con mayor relevancia en la visualización principal.` : null,
          dateProfiles.length ? `La hoja permite lectura temporal y comparación por periodo.` : null,
        ].filter(Boolean),
  };
}

function scoreGenericSheet(summary) {
  const label = `${summary.blockTitle || ''} ${summary.rawName || summary.name || ''}`;
  const keywordBonus = /venta|ingreso|stock|inventario|producto|empleado|personal|rrhh|nomina|costo|gasto|categoria|resumen|dashboard|auditoria|riesgo|control|plan|evidencia|hallazgo|procedimiento|cumplimiento|proceso|seguimiento|reporte|informe/i.test(label) ? 70 : 0;
  const chartBonus = summary.chart ? 55 : 0;
  const numericBonus = summary.numericCount * 18;
  const categoryBonus = summary.categoryCount * 12;
  const sizeBonus = Math.min(summary.rowCount * summary.colCount, 180);
  const titleBonus = summary.blockTitle ? 24 : 0;
  const tableBonus = Math.round((summary.tableCompleteness || 0) * 120) + Math.min((summary.tableRows?.length || 0) * 10, 40);
  const genericPenalty = isGenericSheetLabel(summary.rawName || summary.name) && !summary.blockTitle ? 95 : 0;
  const weakPenalty = summary.rowCount < 8 && summary.numericCount === 0 && summary.categoryCount <= 1 ? 70 : 0;
  const supportPenalty = /posible generalidad|control de cambios|notas|anex|hoja1|hoja\d+/.test(normalizeHeaderKey(label)) ? 45 : 0;
  return keywordBonus + chartBonus + numericBonus + categoryBonus + sizeBonus + titleBonus + tableBonus - genericPenalty - weakPenalty - supportPenalty;
}

function buildGenericLayoutProfile(options = {}) {
  const {
    insights = [],
    notes = [],
    contextSummary = '',
    headlineMetrics = [],
    chart = null,
    primaryVisual = null,
    rowCount = 0,
    colCount = 0,
    numericCount = 0,
    categoryCount = 0,
    dateCount = 0,
  } = options;
  const insightWeight = insights.reduce((acc, line) => acc + cleanText(line).length, 0);
  const noteWeight = notes.reduce((acc, line) => acc + cleanText(line).length, 0);
  const contextWeight = cleanText(contextSummary).length;
  const chartReady = !!(chart && chart.labels?.length >= 2 && chart.labels.length <= 8);
  const narrativeRich = insightWeight + noteWeight + contextWeight >= 170 || notes.length >= 2;
  const metricRich = headlineMetrics.length >= 3 || numericCount >= 3;
  const tableHeavy = primaryVisual?.type === 'table' || colCount >= 6 || rowCount >= 12;
  const dense = rowCount * colCount >= 80 || insightWeight + noteWeight >= 180 || (tableHeavy && narrativeRich);
  const dateDriven = dateCount >= 1;

  return {
    chartReady,
    narrativeRich,
    metricRich,
    tableHeavy,
    dense,
    dateDriven,
    preferredVisual: chartReady && !tableHeavy ? 'chart' : primaryVisual?.type || (tableHeavy ? 'table' : 'mixed'),
  };
}

function selectUniversalLayoutFamily(scope, key, profile = {}) {
  const {
    chartReady = false,
    narrativeRich = false,
    metricRich = false,
    tableHeavy = false,
    dense = false,
    preferredVisual = 'mixed',
  } = profile;

  let families;
  if (scope === 'overview') {
    if (PRESENTATION_VISUAL_MODE === 'tables') families = ['table-overview', 'summary-grid'];
    else if (dense && chartReady) families = ['chart-hero', 'editorial-band', 'summary-grid'];
    else if (narrativeRich) families = ['editorial-band', 'summary-grid', 'chart-hero'];
    else if (metricRich && chartReady) families = ['summary-grid', 'chart-hero', 'split-dashboard'];
    else if (tableHeavy) families = ['summary-grid', 'editorial-band', 'split-dashboard'];
    else families = ['split-dashboard', 'summary-grid', 'chart-hero'];
  } else if (scope === 'focus') {
    if (PRESENTATION_VISUAL_MODE === 'tables') families = ['table-led', 'editorial-split', 'visual-left'];
    else if (PRESENTATION_VISUAL_MODE === 'charts') families = ['visual-right', 'visual-left', 'visual-top', 'editorial-split'];
    else if (narrativeRich && dense) families = ['editorial-split', 'visual-top', 'visual-left'];
    else if (chartReady && !dense) families = ['visual-right', 'visual-left', 'visual-top'];
    else if (tableHeavy) families = ['table-led', 'visual-top', 'editorial-split'];
    else families = ['visual-left', 'visual-right', 'editorial-split', 'visual-top'];
  } else if (scope === 'map') {
    if (PRESENTATION_VISUAL_MODE === 'tables') families = ['summary-table'];
    else if (PRESENTATION_VISUAL_MODE === 'charts') families = ['cards-grid', 'editorial-strip'];
    else if (dense) families = ['summary-table', 'editorial-strip'];
    else if (narrativeRich) families = ['editorial-strip', 'cards-grid', 'summary-table'];
    else families = ['cards-grid', 'editorial-strip', 'summary-table'];
  } else {
    families = ['default'];
  }

  const variantIndex = getScopedVariant(`universal-${scope}`, `${key}|${preferredVisual}|${dense ? 'dense' : 'light'}`, families.length);
  return families[variantIndex] || families[0];
}

function hasRenderableChartSummary(summary) {
  if (!(summary?.chart && summary.chart.labels?.length >= 2 && summary.chart.values?.length >= 2)) return false;
  const labels = summary.chart.labels.map((label) => cleanText(label)).filter(Boolean);
  const values = summary.chart.values.map((value) => Number(value)).filter((value) => Number.isFinite(value) && value > 0);
  const entries = labels.map((label, index) => [label, values[index]]).filter(([, value]) => Number.isFinite(value));
  return labels.length >= 2
    && new Set(labels).size >= 2
    && values.length >= 2
    && hasMeaningfulChartEntries(entries, summary.chart.seriesType || 'count');
}

function hasRenderableTableSummary(summary) {
  return !!(
    summary?.tableHeaders?.length >= 2
    && summary?.tableRows?.length >= 2
    && summary.tableRows.some((row) => countRenderablePreviewCells(row) >= Math.max(2, Math.min(summary.tableHeaders.length, 3)))
    && (summary.tableCompleteness || 0) >= 0.55
  );
}

function buildGenericDataset(workbook) {
  const structuredSheets = workbook.SheetNames
    .flatMap((sheetName) => extractStructuredSheetModels(workbook, sheetName))
    .filter((sheet) => sheet && sheet.rows.length >= 2 && sheet.headers.length >= 2);

  if (!structuredSheets.length) {
    if (PRESENTATION_VISUAL_MODE === 'charts') {
      throw new Error('Este Excel no se puede organizar para solo gráficas porque no tiene una estructura tabular mínima ni datos graficables.');
    }
    if (PRESENTATION_VISUAL_MODE === 'tables') {
      throw new Error('Este Excel no se puede organizar para solo tablas porque no tiene una estructura tabular mínima utilizable.');
    }
    throw new Error('No se encontraron hojas tabulares utilizables en el Excel.');
  }

  const summarizedSheets = structuredSheets
    .map((sheet) => summarizeGenericSheet(sheet))
    .sort((a, b) => scoreGenericSheet(b) - scoreGenericSheet(a));

  let workingSheets = summarizedSheets;
  if (PRESENTATION_VISUAL_MODE === 'charts') {
    workingSheets = summarizedSheets.filter((sheet) => hasRenderableChartSummary(sheet));
    if (!workingSheets.length) {
      throw new Error('Este Excel no tiene suficientes datos válidos para generar una presentación solo de gráficas. Usa el modo mixto o tablas.');
    }
  } else if (PRESENTATION_VISUAL_MODE === 'tables') {
    workingSheets = summarizedSheets.filter((sheet) => hasRenderableTableSummary(sheet));
    if (!workingSheets.length) {
      throw new Error('Este Excel no tiene suficientes datos tabulares válidos para generar una presentación solo de tablas.');
    }
  }

  workingSheets = prioritizeExecutiveSheets(workingSheets);

  const featuredSheets = workingSheets.slice(0, Math.min(4, workingSheets.length));
  const totalRows = workingSheets.reduce((acc, sheet) => acc + sheet.rowCount, 0);
  const totalNumeric = workingSheets.reduce((acc, sheet) => acc + sheet.numericCount, 0);
  const totalCategories = workingSheets.reduce((acc, sheet) => acc + sheet.categoryCount, 0);
  const workbookTitle = cleanText(path.basename(INPUT_FILE, path.extname(INPUT_FILE)).replace(/[_-]+/g, ' '), 80);

  const workbookInsights = featuredSheets
    .flatMap((sheet) => sheet.insights.slice(0, 2))
    .filter(Boolean)
    .slice(0, 4);
  const headlineCards = featuredSheets
    .flatMap((sheet) => sheet.headlineMetrics.slice(0, 2).map((metric) => ({ ...metric, sheet: sheet.name })))
    .slice(0, 4);
  const priorityRows = buildWorkbookPriorityRows(workingSheets);
  const layoutProfile = buildGenericLayoutProfile({
    insights: workbookInsights,
    headlineMetrics: headlineCards,
    chart: featuredSheets[0]?.chart || null,
    primaryVisual: featuredSheets[0]?.primaryVisual || null,
    contextSummary: workbookTitle,
    rowCount: totalRows,
    colCount: Math.max(...workingSheets.map((sheet) => sheet.colCount), 0),
    numericCount: totalNumeric,
    categoryCount: totalCategories,
    dateCount: workingSheets.reduce((acc, sheet) => acc + sheet.dateCount, 0),
  });

  return {
    mode: 'generic',
    generatedAt: new Date(),
    workbookTitle,
    workbookName: cleanText(path.basename(INPUT_FILE)),
    kpis: {
      sheetCount: workbook.SheetNames.length,
      structuredSheets: workingSheets.length,
      totalRows,
      numericColumns: totalNumeric,
      categoryColumns: totalCategories,
    },
    sheets: workingSheets,
    featuredSheets,
    largestSheet: featuredSheets[0],
    headlineCards,
    priorityRows,
    layoutProfile,
    visualMode: PRESENTATION_VISUAL_MODE,
    insights: workbookInsights.length
      ? workbookInsights
      : [`El archivo contiene senales suficientes para construir una lectura ejecutiva con foco de negocio.`],
  };
}

function normalizeOrganizerHeaders(headers = []) {
  const seen = new Map();
  return headers.map((header, index) => {
    const base = cleanText(header) || `Columna ${index + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count ? `${base} ${count + 1}` : base;
  });
}

function buildStructuredRowsFromOrganizer(headers, rows = []) {
  return rows
    .map((row) => {
      const currentRow = Array.isArray(row) ? row : [];
      const mapped = {};
      headers.forEach((header, index) => {
        mapped[header] = currentRow[index] ?? '';
      });
      return mapped;
    })
    .filter((row) => Object.values(row).some((value) => cleanText(value) || toNumber(value) !== null || toDateValue(value)));
}

function organizerTableToStructuredSheet(sheetName, table, options = {}) {
  const rawHeaders = Array.isArray(table?.encabezados) ? table.encabezados : [];
  const headers = normalizeOrganizerHeaders(rawHeaders).slice(0, 8);
  const rawRows = Array.isArray(table?.filas) ? table.filas : [];
  const rows = buildStructuredRowsFromOrganizer(headers, rawRows.map((row) => row.slice(0, headers.length))).slice(0, 60);

  if (headers.length < 2 || rows.length < 2) return null;

  return {
    name: sheetName,
    sourceSheetName: options.sourceSheetName || sheetName,
    displayName: options.displayName || sheetName,
    title: options.title || sheetName,
    subtitle: options.subtitle || '',
    notes: (options.notes || []).filter(Boolean).map((note) => cleanText(note, 120)),
    blockCount: 1,
    blockIndex: 0,
    headers,
    rows,
  };
}

function collectOrganizerStructuredSheets(sourceData) {
  const structured = [];
  const workbookName = cleanText(sourceData?.metadatos?.archivo) || SOURCE_FILE_LABEL;
  const mainSheetLabel = cleanText(sourceData?.resumen_generico?.hoja_principal)
    || cleanText(sourceData?.metadatos?.hojas_encontradas?.[0])
    || 'Hoja principal';

  if (sourceData?.muestra_tabla) {
    structured.push(organizerTableToStructuredSheet(mainSheetLabel, sourceData.muestra_tabla, {
      sourceSheetName: mainSheetLabel,
      displayName: mainSheetLabel,
      title: mainSheetLabel,
      subtitle: `Tabla principal detectada en ${workbookName}`,
      notes: [
        sourceData?.resumen_generico?.columnas?.length
          ? `Columnas visibles: ${sourceData.resumen_generico.columnas.slice(0, 6).join(', ')}`
          : null,
      ],
    }));
  }

  Object.entries(sourceData?.otras_tablas || {}).forEach(([name, table]) => {
    structured.push(organizerTableToStructuredSheet(name, table, {
      sourceSheetName: name,
      displayName: name,
      title: name,
      subtitle: `Tabla organizada desde la hoja ${name}`,
    }));
  });

  Object.entries(sourceData?.genericas || {}).forEach(([name, table]) => {
    structured.push(organizerTableToStructuredSheet(name, table, {
      sourceSheetName: name,
      displayName: name,
      title: name,
      subtitle: `Vista estructurada de la hoja ${name}`,
    }));
  });

  if (sourceData?.coso) {
    structured.push(organizerTableToStructuredSheet('COSO', sourceData.coso, {
      sourceSheetName: 'COSO',
      displayName: 'COSO',
      title: 'Evaluación COSO',
      subtitle: 'Matriz de control interno identificada en el archivo',
    }));
  }

  if (sourceData?.distribucion_mes) {
    structured.push(organizerTableToStructuredSheet('Distribución', sourceData.distribucion_mes, {
      sourceSheetName: 'Distribución',
      displayName: 'Distribución',
      title: 'Distribución detectada',
      subtitle: 'Resumen complementario detectado en el archivo',
    }));
  }

  return structured.filter(Boolean);
}

function buildGenericDatasetFromOrganizer(sourceData, workbook) {
  const structuredSheets = collectOrganizerStructuredSheets(sourceData);
  if (!structuredSheets.length) {
    return buildGenericDataset(workbook);
  }

  const summarizedSheets = structuredSheets
    .map((sheet) => summarizeGenericSheet(sheet))
    .sort((a, b) => scoreGenericSheet(b) - scoreGenericSheet(a));

  let workingSheets = summarizedSheets;
  if (PRESENTATION_VISUAL_MODE === 'charts') {
    workingSheets = summarizedSheets.filter((sheet) => hasRenderableChartSummary(sheet));
    if (!workingSheets.length) {
      throw new Error('Este Excel no tiene suficientes datos organizados para generar una presentación solo de gráficas.');
    }
  } else if (PRESENTATION_VISUAL_MODE === 'tables') {
    workingSheets = summarizedSheets.filter((sheet) => hasRenderableTableSummary(sheet));
    if (!workingSheets.length) {
      throw new Error('Este Excel no tiene suficientes datos organizados para generar una presentación solo de tablas.');
    }
  }

  workingSheets = prioritizeExecutiveSheets(workingSheets);

  const featuredSheets = workingSheets.slice(0, Math.min(4, workingSheets.length));
  const totalRows = workingSheets.reduce((acc, sheet) => acc + sheet.rowCount, 0);
  const totalNumeric = workingSheets.reduce((acc, sheet) => acc + sheet.numericCount, 0);
  const totalCategories = workingSheets.reduce((acc, sheet) => acc + sheet.categoryCount, 0);
  const sheetCount = Array.isArray(sourceData?.metadatos?.hojas_encontradas) ? sourceData.metadatos.hojas_encontradas.length : workbook.SheetNames.length;
  const focusSheet = featuredSheets[0]?.blockTitle || featuredSheets[0]?.name || cleanText(sourceData?.resumen_generico?.hoja_principal) || 'Hoja principal';
  const workbookTitle = cleanText(
    sourceData?.metadatos?.archivo
      ? path.basename(sourceData.metadatos.archivo, path.extname(sourceData.metadatos.archivo)).replace(/[_-]+/g, ' ')
      : path.basename(INPUT_FILE, path.extname(INPUT_FILE)).replace(/[_-]+/g, ' '),
    80
  );

  const organizerInsights = [
    sourceData?.resumen_generico?.hoja_principal
      ? `La hoja ${cleanText(sourceData.resumen_generico.hoja_principal, 34)} se toma como base principal para estructurar la narrativa.`
      : null,
    Array.isArray(sourceData?.metadatos?.hojas_encontradas) && sourceData.metadatos.hojas_encontradas.length
      ? `Se detectan ${fmtNumber(sourceData.metadatos.hojas_encontradas.length)} hojas útiles en el archivo.`
      : null,
    sourceData?.resumen_generico?.columnas_numericas?.length
      ? `Las columnas numéricas priorizadas son ${sourceData.resumen_generico.columnas_numericas.slice(0, 3).map((value) => cleanText(value, 22)).join(', ')}.`
      : null,
  ].filter(Boolean);

  const workbookInsights = [
    ...organizerInsights,
    ...featuredSheets.flatMap((sheet) => sheet.insights.slice(0, 2)),
  ].filter(Boolean).slice(0, 4);

  const headlineCards = [
    { label: 'Hoja foco', value: focusSheet, note: 'Base principal organizada' },
    { label: 'Hojas útiles', value: fmtNumber(sheetCount), note: 'Estructuras detectadas' },
    sourceData?.resumen_generico?.total_filas
      ? { label: 'Filas base', value: fmtNumber(sourceData.resumen_generico.total_filas), note: 'Hoja principal organizada' }
      : null,
    sourceData?.resumen_generico?.total_columnas
      ? { label: 'Columnas base', value: fmtNumber(sourceData.resumen_generico.total_columnas), note: 'Campos visibles' }
      : null,
    featuredSheets[0]?.headlineMetrics?.[0]
      ? { ...featuredSheets[0].headlineMetrics[0], sheet: featuredSheets[0].name }
      : null,
  ].filter(Boolean).slice(0, 4);
  const priorityRows = buildWorkbookPriorityRows(workingSheets);
  const layoutProfile = buildGenericLayoutProfile({
    insights: workbookInsights,
    headlineMetrics: headlineCards,
    chart: featuredSheets[0]?.chart || null,
    primaryVisual: featuredSheets[0]?.primaryVisual || null,
    contextSummary: workbookTitle,
    rowCount: totalRows,
    colCount: Math.max(...workingSheets.map((sheet) => sheet.colCount), 0),
    numericCount: totalNumeric,
    categoryCount: totalCategories,
    dateCount: workingSheets.reduce((acc, sheet) => acc + sheet.dateCount, 0),
  });

  return {
    mode: 'generic',
    generatedAt: new Date(),
    workbookTitle,
    workbookName: cleanText(sourceData?.metadatos?.archivo || path.basename(INPUT_FILE)),
    kpis: {
      sheetCount,
      structuredSheets: workingSheets.length,
      totalRows,
      numericColumns: totalNumeric,
      categoryColumns: totalCategories,
    },
    sheets: workingSheets,
    featuredSheets,
    largestSheet: featuredSheets[0],
    headlineCards,
    priorityRows,
    layoutProfile,
    visualMode: PRESENTATION_VISUAL_MODE,
    insights: workbookInsights.length
      ? workbookInsights
      : [`El archivo se organizó desde las hojas detectadas para construir una lectura ejecutiva fiel al Excel.`],
  };
}

function applyTransition(slide, type = 'fade', speed = 'med') {
  slide.transition = { type, speed };
}

function addNatureDecor(slide, mode = 'header') {
  const lineColor = 'FFFFFF';
  const lineTransparency = mode === 'cover' || mode === 'closing' ? 80 : 86;
  const leafTransparency = mode === 'cover' || mode === 'closing' ? 83 : 90;
  const decorVariant = getDeckVariant('decor');
  const leafSets = decorVariant === 1
    ? [
        { x: 10.05, y: 1.02, rotate: 28, w: 0.46, h: 0.2 },
        { x: 10.48, y: 1.3, rotate: -24, w: 0.44, h: 0.18 },
        { x: 10.88, y: 0.95, rotate: 18, w: 0.42, h: 0.18 },
        { x: 11.22, y: 5.28, rotate: -22, w: 0.42, h: 0.18 },
        { x: 11.66, y: 5.58, rotate: 30, w: 0.46, h: 0.2 },
      ]
    : decorVariant === 2
      ? [
          { x: 10.22, y: 1.08, rotate: 42, w: 0.48, h: 0.2 },
          { x: 10.62, y: 1.42, rotate: -28, w: 0.44, h: 0.18 },
          { x: 11.06, y: 1.0, rotate: 16, w: 0.4, h: 0.17 },
          { x: 11.28, y: 5.3, rotate: -36, w: 0.44, h: 0.18 },
          { x: 11.72, y: 5.62, rotate: 20, w: 0.44, h: 0.18 },
        ]
      : [
          { x: 10.0, y: 1.08, rotate: 38, w: 0.44, h: 0.2 },
          { x: 10.34, y: 1.42, rotate: -28, w: 0.46, h: 0.2 },
          { x: 10.78, y: 1.0, rotate: 24, w: 0.42, h: 0.18 },
          { x: 11.24, y: 5.42, rotate: -32, w: 0.44, h: 0.18 },
          { x: 11.62, y: 5.7, rotate: 22, w: 0.46, h: 0.2 },
        ];

  const lineSets = decorVariant === 1
    ? [
        { x: 9.62, y: 0.9, w: 2.5, h: 1.46, pt: 0.9, transparency: lineTransparency },
        { x: 10.2, y: 1.26, w: 1.92, h: 0.04, pt: 0.9, transparency: lineTransparency + 4 },
        { x: 10.7, y: 5.2, w: 1.72, h: 1.02, pt: 0.9, transparency: lineTransparency + 8 },
      ]
    : decorVariant === 2
      ? [
          { x: 9.4, y: 1.02, w: 3.0, h: 1.02, pt: 0.9, transparency: lineTransparency },
          { x: 9.95, y: 0.9, w: 2.35, h: 1.46, pt: 1.0, transparency: lineTransparency + 2 },
          { x: 10.92, y: 5.08, w: 1.5, h: 1.1, pt: 0.9, transparency: lineTransparency + 8 },
        ]
      : [
          { x: 9.5, y: 0.92, w: 2.72, h: 1.28, pt: 1.0, transparency: lineTransparency },
          { x: 10.02, y: 1.28, w: 2.16, h: 0.02, pt: 0.9, transparency: lineTransparency + 5 },
          { x: 10.84, y: 5.34, w: 1.56, h: 0.88, pt: 0.9, transparency: lineTransparency + 8 },
        ];

  lineSets.forEach((item) => {
    slide.addShape('line', {
      x: item.x,
      y: item.y,
      w: item.w,
      h: item.h,
      line: { color: lineColor, pt: item.pt, transparency: item.transparency },
    });
  });

  leafSets.forEach((leaf) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: leaf.x,
      y: leaf.y,
      w: leaf.w,
      h: leaf.h,
      rotate: leaf.rotate,
      line: { color: lineColor, pt: 1, transparency: leafTransparency },
      fill: { color: lineColor, transparency: 100 },
    });
  });
}

function addFooter(slide, pageLabel) {
  slide.addText(pageLabel, {
    x: 11.75,
    y: 7.08,
    w: 0.9,
    h: 0.18,
    fontSize: 8.5,
    color: COLORS.navy2,
    align: 'right',
    fontFace: 'Calibri',
  });

  slide.addText(`Fuente: ${SOURCE_FILE_LABEL} · Fundación Socya`, {
    x: 0.55,
    y: 7.05,
    w: 6.5,
    h: 0.2,
    fontSize: 8,
    color: COLORS.navy2,
    fontFace: 'Calibri',
  });
}

function addHeader(slide, pageLabel, section, title, subtitle) {
  const headerVariant = getDeckVariant('header');
  const headingFont = getDeckFont('head');
  const bodyFont = getDeckFont('body');
  const layout = getHeaderLayoutMetrics(headerVariant, title, subtitle);
  slide.background = { color: COLORS.bg };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: headerVariant === 1 ? 0.94 : 1.08,
    line: { color: DECK_STYLE.topBandAccent, transparency: 100 },
    fill: { color: DECK_STYLE.topBandColor },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: headerVariant === 2 ? 9.55 : 9.2,
    y: 0,
    w: headerVariant === 1 ? 3.65 : headerVariant === 2 ? 3.78 : 4.13,
    h: 7.5,
    line: { color: DECK_STYLE.sidePanelColor, transparency: 100 },
    fill: { color: DECK_STYLE.sidePanelColor, transparency: DECK_STYLE.sidePanelTransparency },
  });
  addNatureDecor(slide, 'header');
  slide.addShape(pptx.ShapeType.roundRect, {
    x: headerVariant === 1 ? 9.68 : 0.55,
    y: headerVariant === 2 ? 0.3 : 0.34,
    w: headerVariant === 1 ? 2.05 : 2.15,
    h: 0.32,
    rectRadius: 0.04,
    line: { color: headerVariant === 2 ? getDeckAccent(0) : COLORS.white, transparency: headerVariant === 2 ? 15 : 100, pt: headerVariant === 2 ? 1 : 0.6 },
    fill: { color: headerVariant === 2 ? COLORS.white : getSoftAccentColor(getDeckAccent(2)), transparency: headerVariant === 2 ? 3 : 0 },
  });
  slide.addText(section, {
    x: headerVariant === 1 ? 9.83 : 0.7,
    y: headerVariant === 2 ? 0.39 : 0.4,
    w: 1.85,
    h: 0.18,
    fontSize: 9,
    bold: true,
    color: headerVariant === 2 ? COLORS.navy3 : COLORS.navy,
    fontFace: bodyFont,
    align: 'center',
  });
  if (headerVariant === 2) {
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.55,
      y: 1.12,
      w: 0.08,
      h: subtitle ? 0.92 : 0.58,
      line: { color: getDeckAccent(0), transparency: 100 },
      fill: { color: getDeckAccent(0) },
    });
  }
  slide.addText(title, {
    x: layout.titleX,
    y: layout.titleY,
    w: layout.titleW,
    h: layout.titleH,
    fontSize: 28,
    bold: true,
    color: COLORS.navy3,
    fontFace: headingFont,
    fit: layout.titleLines >= 3 ? 'shrink' : undefined,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: headerVariant === 2 ? 0.82 : 0.55,
      y: layout.subtitleY,
      w: layout.subtitleW,
      h: layout.subtitleH,
      fontSize: 10.5,
      color: COLORS.muted,
      fontFace: bodyFont,
      italic: true,
      fit: layout.subtitleLines >= 3 ? 'shrink' : undefined,
    });
  }
  slide.addShape(pptx.ShapeType.roundRect, {
    x: headerVariant === 1 ? 0.55 : 10.65,
    y: 0.24,
    w: 2.12,
    h: 0.42,
    rectRadius: 0.06,
    line: { color: COLORS.white, pt: 1 },
    fill: { color: DECK_STYLE.topBandAccent },
  });
  slide.addText('SOCYA', {
    x: headerVariant === 1 ? 0.85 : 10.95,
    y: 0.34,
    w: 0.95,
    h: 0.16,
    fontSize: 11,
    bold: true,
    color: COLORS.white,
    fontFace: headingFont,
    align: 'center',
  });
  slide.addText(pageLabel, {
    x: headerVariant === 1 ? 1.85 : 11.95,
    y: 0.34,
    w: 0.55,
    h: 0.16,
    fontSize: 10,
    bold: true,
    color: COLORS.white,
    fontFace: headingFont,
    align: 'center',
  });
  addFooter(slide, pageLabel);
}

function addMetricCard(slide, x, y, w, h, label, value, accent, note) {
  const variant = getDeckVariant('metricCard');
  const softAccent = getSoftAccentColor(accent);
  const headingFont = getDeckFont('head');
  const bodyFont = getDeckFont('body');
  const textX = variant === 1 ? x + 0.28 : x + 0.22;
  const textW = w - (variant === 1 ? 0.42 : 0.35);
  const labelY = y + 0.16;
  const valueY = variant === 2 ? y + 0.4 : y + 0.34;
  const noteY = y + h - 0.24;

  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    line: { color: variant === 2 ? accent : COLORS.borderDark, pt: variant === 1 ? 1.0 : 1.2, transparency: variant === 2 ? 30 : 0 },
    fill: { color: variant === 1 ? softAccent : 'FFFFFF', transparency: variant === 1 ? 0 : 5 },
  });

  if (variant === 0) {
    slide.addShape(pptx.ShapeType.rect, {
      x,
      y,
      w,
      h: 0.08,
      line: { color: accent, transparency: 100 },
      fill: { color: accent },
    });
  } else if (variant === 1) {
    slide.addShape(pptx.ShapeType.rect, {
      x: x + 0.12,
      y: y + 0.12,
      w: 0.08,
      h: h - 0.24,
      line: { color: accent, transparency: 100 },
      fill: { color: accent },
    });
  } else {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.16,
      y: y + 0.12,
      w: Math.min(0.42, Math.max(0.3, w * 0.18)),
      h: 0.18,
      rectRadius: 0.05,
      line: { color: accent, transparency: 100 },
      fill: { color: accent },
    });
  }

  slide.addText(label, {
    x: variant === 2 ? x + 0.64 : textX,
    y: labelY,
    w: variant === 2 ? Math.max(0.95, w - 0.78) : textW,
    h: 0.16,
    fontSize: 9,
    bold: true,
    color: variant === 2 ? accent : COLORS.muted,
    fontFace: bodyFont,
    fit: 'shrink',
  });
  slide.addText(value, {
    x: textX,
    y: valueY,
    w: textW,
    h: 0.34,
    fontSize: variant === 2 ? 19.5 : 22,
    bold: true,
    color: COLORS.navy3,
    fontFace: headingFont,
    align: 'left',
    fit: 'shrink',
  });
  if (note) {
    slide.addText(note, {
      x: textX,
      y: noteY,
      w: textW,
      h: 0.14,
      fontSize: 8.6,
      color: COLORS.muted,
      fontFace: bodyFont,
      italic: true,
      fit: 'shrink',
    });
  }
}

function addCompactStatCard(slide, x, y, w, h, label, value, accent) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.06,
    line: { color: COLORS.borderDark, pt: 1 },
    fill: { color: COLORS.greenSoft },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.08,
    y: y + 0.12,
    w: 0.06,
    h: h - 0.24,
    line: { color: accent, transparency: 100 },
    fill: { color: accent },
  });
  slide.addText(value, {
    x: x + 0.22,
    y: y + 0.16,
    w: w - 0.32,
    h: 0.3,
    fontSize: 21,
    bold: true,
    color: COLORS.navy3,
    fontFace: getDeckFont('head'),
    fit: 'shrink',
  });
  slide.addText(label, {
    x: x + 0.22,
    y: y + h - 0.24,
    w: w - 0.32,
    h: 0.14,
    fontSize: 8.8,
    color: COLORS.muted,
    fontFace: getDeckFont('body'),
    fit: 'shrink',
  });
}

function addInsightBox(slide, x, y, w, h, title, lines, fillColor = COLORS.card) {
  const variant = getDeckVariant('insightBox');
  const panelColor = fillColor === COLORS.card ? COLORS.greenSoft : fillColor;
  const headlineAccent = isSoftAccentColor(panelColor) ? getDeckAccent(0) : panelColor;
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    line: { color: variant === 2 ? headlineAccent : COLORS.borderDark, pt: 1.2, transparency: variant === 2 ? 28 : 0 },
    fill: { color: variant === 1 ? getSoftAccentColor(headlineAccent) : panelColor, transparency: 0 },
  });
  if (variant === 0) {
    slide.addShape(pptx.ShapeType.rect, {
      x: x + 0.16,
      y: y + 0.12,
      w: 0.06,
      h: h - 0.24,
      line: { color: headlineAccent, transparency: 100 },
      fill: { color: headlineAccent },
    });
  } else if (variant === 1) {
    slide.addShape(pptx.ShapeType.rect, {
      x: x + 0.14,
      y: y + 0.14,
      w: w - 0.28,
      h: 0.08,
      line: { color: headlineAccent, transparency: 100 },
      fill: { color: headlineAccent },
    });
  } else {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: x + 0.18,
      y: y + 0.12,
      w: 1.05,
      h: 0.24,
      rectRadius: 0.04,
      line: { color: headlineAccent, transparency: 100 },
      fill: { color: headlineAccent },
    });
  }
  slide.addText(title, {
    x: variant === 0 ? x + 0.32 : x + 0.24,
    y: y + 0.12,
    w: w - 0.4,
    h: 0.2,
    fontSize: 11,
    bold: true,
    color: variant === 2 ? COLORS.white : COLORS.navy3,
    fontFace: getDeckFont('head'),
  });

  // Render without truncation. Wrap lines and adapt spacing to avoid overlaps.
  const contentX = variant === 0 ? x + 0.32 : x + 0.24;
  const contentW = w - (variant === 0 ? 0.48 : 0.4);
  const topY = y + 0.42;
  const bottomY = y + h - 0.14;
  let cursorY = topY;
  const baseFontSize = 9;
  const minFontSize = 7.2;

  lines
    .map((line) => cleanText(line))
    .filter(Boolean)
    .slice(0, 4)
    .forEach((line, index) => {
      const prefix = `${index + 1}. `;
      let fontSize = baseFontSize;
      let neededLines = estimateWrappedLines(prefix + line, contentW, fontSize);
      let boxHeight = 0.06 + neededLines * 0.18;

      while (cursorY + boxHeight > bottomY && fontSize > minFontSize) {
        fontSize -= 0.3;
        neededLines = estimateWrappedLines(prefix + line, contentW, fontSize);
        boxHeight = 0.06 + neededLines * 0.18;
      }

      // If still tight, allow PowerPoint to shrink within the box instead of truncating.
      const safeHeight = Math.max(0.22, Math.min(boxHeight, bottomY - cursorY));
      slide.addText(prefix + line, {
        x: contentX,
        y: cursorY,
        w: contentW,
        h: safeHeight,
        fontSize,
        color: COLORS.ink,
        fontFace: getDeckFont('body'),
        breakLine: true,
        valign: 'top',
        margin: 0.02,
        fit: 'shrink',
      });

      cursorY += safeHeight + 0.06;
    });
}

function addCover(prs, dataset) {
  const slide = prs.addSlide();
  slide.background = { color: COLORS.navy };
  applyTransition(slide, 'fade', 'slow');
  const coverVariant = getDeckVariant('commissionsCover');
  const headingFont = getDeckFont('head');
  const bodyFont = getDeckFont('body');
  const accentSet = getDeckChartPalette(3);

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: coverVariant === 1 ? 13.33 : 4.6,
    h: coverVariant === 1 ? 1.08 : 7.5,
    line: { color: DECK_STYLE.topBandAccent, transparency: 100 },
    fill: { color: coverVariant === 1 ? DECK_STYLE.topBandColor : DECK_STYLE.topBandAccent },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: coverVariant === 2 ? 8.7 : 9.35,
    y: 0,
    w: coverVariant === 1 ? 4.15 : coverVariant === 2 ? 4.63 : 3.98,
    h: 7.5,
    line: { color: DECK_STYLE.sidePanelColor, transparency: 100 },
    fill: { color: DECK_STYLE.sidePanelColor, transparency: DECK_STYLE.sidePanelTransparency },
  });
  addNatureDecor(slide, 'cover');
  if (coverVariant !== 1) {
    slide.addShape(pptx.ShapeType.rect, {
      x: 4.6,
      y: 0,
      w: 0.08,
      h: 7.5,
      line: { color: COLORS.white, transparency: 18 },
      fill: { color: COLORS.white, transparency: 18 },
    });
  }
  slide.addShape(pptx.ShapeType.roundRect, {
    x: coverVariant === 1 ? 0.92 : 4.7,
    y: coverVariant === 1 ? 1.38 : 0.88,
    w: 2.6,
    h: 0.38,
    rectRadius: 0.06,
    line: { color: COLORS.white, transparency: 100 },
    fill: { color: getSoftAccentColor(getDeckAccent(2)) },
  });
  slide.addText('COMITÉ DIRECTIVO', {
    x: coverVariant === 1 ? 1.2 : 4.98,
    y: coverVariant === 1 ? 1.48 : 0.98,
    w: 2.05,
    h: 0.15,
    fontSize: 10,
    bold: true,
    color: COLORS.navy3,
    fontFace: bodyFont,
    align: 'center',
  });
  slide.addText('Auditoría de Comisiones y Control Interno', {
    x: coverVariant === 1 ? 0.92 : 4.7,
    y: coverVariant === 1 ? 2.02 : 1.6,
    w: coverVariant === 2 ? 3.75 : 4.55,
    h: coverVariant === 1 ? 1.05 : 1.4,
    fontSize: 28,
    bold: true,
    color: COLORS.white,
    fontFace: headingFont,
  });
  slide.addText(`Presentación ejecutiva generada a partir de datos reales del archivo ${SOURCE_FILE_LABEL}, con foco en volumen, gasto, hallazgos, oportunidades y madurez de control.`, {
    x: coverVariant === 1 ? 0.92 : 4.7,
    y: coverVariant === 1 ? 3.2 : 3.0,
    w: coverVariant === 2 ? 3.55 : 5.9,
    h: coverVariant === 1 ? 0.72 : 0.95,
    fontSize: 13.5,
    color: COLORS.textLight,
    fontFace: bodyFont,
    italic: true,
    fit: 'shrink',
  });

  const dateText = dataset.generatedAt.toLocaleDateString('es-CO', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
  });
  const coverCards = [
    ['Comisiones analizadas', fmtNumber(dataset.kpis.totalRows)],
    ['Valor total', fmtCompactCurrency(dataset.kpis.totalValue)],
    ['Solicitantes únicos', fmtNumber(dataset.kpis.uniqueRequesters)],
  ];

  if (coverVariant === 1) {
    coverCards.forEach(([label, value], index) => {
      addMetricCard(slide, 0.95 + index * 2.7, 4.45, 2.45, 1.05, label, value, accentSet[index], 'Indicador clave');
    });
    addInsightBox(slide, 9.15, 1.6, 3.35, 4.55, 'Mensajes clave', dataset.insights.slice(0, 3), getSoftAccentColor(getDeckAccent(1)));
  } else if (coverVariant === 2) {
    coverCards.forEach(([label, value], index) => {
      const y = 1.65 + index * 0.92;
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 9.1,
        y,
        w: 3.15,
        h: 0.68,
        rectRadius: 0.07,
        line: { color: COLORS.white, pt: 1, transparency: 24 },
        fill: { color: DECK_STYLE.topBandAccent, transparency: 28 },
      });
      slide.addText(label, {
        x: 9.34,
        y: y + 0.13,
        w: 1.5,
        h: 0.16,
        fontSize: 8.8,
        bold: true,
        color: COLORS.textSoft,
        fontFace: bodyFont,
      });
      slide.addText(value, {
        x: 10.55,
        y: y + 0.11,
        w: 1.45,
        h: 0.2,
        fontSize: 15.5,
        bold: true,
        color: COLORS.white,
        align: 'right',
        fontFace: headingFont,
      });
    });
    addInsightBox(slide, 4.7, 4.35, 3.5, 1.75, 'Mensajes clave', dataset.insights.slice(0, 2), getSoftAccentColor(getDeckAccent(0)));
  } else {
    coverCards.forEach(([label, value], index) => {
      const y = 4.45 + index * 0.78;
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 4.7,
        y,
        w: 3.45,
        h: 0.54,
        rectRadius: 0.06,
        line: { color: COLORS.white, pt: 1, transparency: 30 },
        fill: { color: DECK_STYLE.topBandAccent, transparency: 22 },
      });
      slide.addShape(pptx.ShapeType.rect, {
        x: 4.7,
        y,
        w: 0.08,
        h: 0.54,
        line: { color: COLORS.white, transparency: 35 },
        fill: { color: COLORS.white, transparency: 35 },
      });
      slide.addText(label, {
        x: 4.92,
        y: y + 0.1,
        w: 1.95,
        h: 0.16,
        fontSize: 9,
        bold: true,
        color: COLORS.textLight,
        fontFace: bodyFont,
      });
      slide.addText(value, {
        x: 6.35,
        y: y + 0.12,
        w: 1.55,
        h: 0.2,
        fontSize: 16,
        bold: true,
        color: COLORS.textLight,
        align: 'right',
        fontFace: headingFont,
      });
    });

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 9.85,
      y: 1.22,
      w: 2.55,
      h: 4.7,
      rectRadius: 0.08,
      line: { color: COLORS.white, pt: 1, transparency: 28 },
      fill: { color: DECK_STYLE.topBandAccent, transparency: 24 },
    });
    slide.addText('Mensajes clave', {
      x: 10.12,
      y: 1.48,
      w: 1.8,
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: COLORS.white,
      fontFace: headingFont,
    });
    dataset.insights.slice(0, 3).forEach((line, index) => {
      slide.addText(`${index + 1}. ${line}`, {
        x: 10.12,
        y: 1.95 + index * 1.02,
        w: 1.95,
        h: 0.72,
        fontSize: 10.2,
        color: COLORS.textSoft,
        fontFace: bodyFont,
      });
    });
  }

  slide.addText(`Fundación Socya · ${dateText}`, {
    x: coverVariant === 1 ? 0.92 : 4.7,
    y: 6.75,
    w: 4.8,
    h: 0.2,
    fontSize: 10,
    color: COLORS.textSoft,
    fontFace: bodyFont,
  });
}

function addExecutiveSummary(prs, dataset) {
  const slide = prs.addSlide();
  applyTransition(slide, 'push', 'med');
  const sectionTitle = 'Panorama general del universo auditado';
  const sectionSubtitle = 'La narrativa inicia con escala, valor económico y concentración operativa del proceso de comisiones.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 2.0, 0.26);
  addHeader(
    slide,
    '01',
    'RESUMEN EJECUTIVO',
    sectionTitle,
    sectionSubtitle
  );

  const metrics = [
    ['Comisiones', fmtNumber(dataset.kpis.totalRows), 'Base total revisada'],
    ['Valor total', fmtCompactCurrency(dataset.kpis.totalValue), 'Monto solicitado acumulado'],
    ['Promedio', fmtCompactCurrency(dataset.kpis.avgValue), 'Valor medio por comisión'],
    ['Solicitantes', fmtNumber(dataset.kpis.uniqueRequesters), 'Colaboradores diferentes'],
    ['Ciudades', fmtNumber(dataset.kpis.uniqueCities), 'Cobertura territorial'],
  ];

  metrics.forEach((item, index) => {
    addMetricCard(
      slide,
      0.55 + index * 2.52,
      topY,
      2.32,
      1.18,
      item[0],
      item[1],
      [COLORS.green, COLORS.aqua, COLORS.amber, '5B8DEF', '8B5CF6'][index],
      item[2]
    );
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: 3.3,
    w: 6.1,
    h: 3.1,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Distribución por estado', {
    x: 0.8,
    y: 3.55,
    w: 2.6,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: 'Calibri',
  });

  slide.addChart(
    prs.ChartType.doughnut,
    [
      {
        name: 'Estados',
        labels: dataset.charts.statusCounts.map(([label]) => cleanText(label, 22)),
        values: dataset.charts.statusCounts.map(([, value]) => value),
      },
    ],
    {
      x: 0.8,
      y: 3.95,
      w: 2.55,
      h: 2.05,
      holeSize: 64,
      showLegend: true,
      legendPos: 'r',
      legendFontSize: 8.5,
      legendFontColor: COLORS.muted,
      chartColors: dataset.charts.statusCounts.map(([label], index) => getStateChartColor(label, index)),
      showValue: true,
      dataLabelPosition: 'bestFit',
      dataLabelColor: COLORS.ink,
      dataLabelFontSize: 8,
    }
  );

  addInsightBox(
    slide,
    3.55,
    3.82,
    2.8,
    2.28,
    'Lectura ejecutiva',
    dataset.insights.slice(0, 3),
    COLORS.greenSoft
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.9,
    y: 3.3,
    w: 5.88,
    h: 3.1,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Concentración territorial y de centros de costo', {
    x: 7.15,
    y: 3.55,
    w: 4.8,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: 'Calibri',
  });

  dataset.charts.topCities.slice(0, 4).forEach(([label, value], index) => {
    const y = 3.95 + index * 0.52;
    const width = Math.max(0.5, (value / dataset.charts.topCities[0][1]) * 2.45);
    slide.addText(cleanText(label, 28), {
      x: 7.15,
      y,
      w: 2.2,
      h: 0.15,
      fontSize: 8.5,
      color: COLORS.ink,
      fontFace: 'Calibri',
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 9.3,
      y: y + 0.02,
      w: 2.6,
      h: 0.16,
      rectRadius: 0.04,
      line: { color: COLORS.line, transparency: 100 },
      fill: { color: 'E8EEF6' },
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 9.3,
      y: y + 0.02,
      w: width,
      h: 0.16,
      rectRadius: 0.04,
      line: { color: COLORS.aqua, transparency: 100 },
      fill: { color: COLORS.aqua },
    });
    slide.addText(fmtNumber(value), {
      x: 11.95,
      y: y - 0.02,
      w: 0.5,
      h: 0.18,
      fontSize: 8.5,
      bold: true,
      color: COLORS.ink,
      align: 'right',
      fontFace: 'Calibri',
    });
  });

  dataset.charts.topCostCenters.slice(0, 3).forEach(([label, value], index) => {
    slide.addText(`${cleanText(label)} · ${fmtCompactCurrency(value)}`, {
      x: 7.15,
      y: 6.0 + index * 0.18,
      w: 5.1,
      h: 0.16,
      fontSize: 8.5,
      color: COLORS.muted,
      fontFace: 'Calibri',
    });
  });
}

function addFinancialSlide(prs, dataset) {
  const slide = prs.addSlide();
  applyTransition(slide, 'fade', 'med');
  const sectionTitle = 'Composición del gasto solicitado';
  const sectionSubtitle = 'Se contrastan montos por rubro con el comportamiento económico agregado del portafolio de comisiones.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.3);
  addHeader(
    slide,
    '02',
    'ANÁLISIS FINANCIERO',
    sectionTitle,
    sectionSubtitle
  );

  [
    ['Valor total', fmtCurrency(dataset.kpis.totalValue), COLORS.green],
    ['Promedio por comisión', fmtCurrency(dataset.kpis.avgValue), COLORS.aqua],
    ['Comisión máxima', fmtCurrency(dataset.kpis.maxValue), COLORS.amber],
  ].forEach(([label, value, accent], index) => {
    addMetricCard(slide, 0.55 + index * 4.1, topY, 3.8, 1.08, label, value, accent);
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: Math.max(3.2, topY + 1.42),
    w: 7.25,
    h: 3.1,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Montos acumulados por tipo de gasto', {
    x: 0.82,
    y: Math.max(3.45, topY + 1.67),
    w: 3.6,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: 'Calibri',
  });
  slide.addChart(
    prs.ChartType.bar,
    [
      {
        name: 'Valor',
        labels: dataset.charts.expenseSummary.map((item) => item.label),
        values: dataset.charts.expenseSummary.map((item) => item.value),
      },
    ],
    {
      x: 0.82,
      y: Math.max(3.82, topY + 2.04),
      w: 6.6,
      h: 2.08,
      barDir: 'col',
      chartColors: [COLORS.aqua, COLORS.green, COLORS.amber, '5B8DEF'],
      showValue: true,
      dataLabelPosition: 'outEnd',
      dataLabelColor: COLORS.ink,
      valAxisLabelColor: COLORS.muted,
      catAxisLabelColor: COLORS.muted,
      valAxisNumFmt: '$#,##0',
      gridLineColor: 'D9E3EE',
      showLegend: false,
    }
  );

  addInsightBox(
    slide,
    8.05,
    Math.max(3.2, topY + 1.42),
    4.73,
    3.1,
    'Mensajes de negocio',
    [
      dataset.insights[1],
      `La base mueve ${fmtCompactCurrency(dataset.kpis.totalValue)} y combina cobertura geográfica extensa con un ticket promedio cercano a ${fmtCompactCurrency(dataset.kpis.avgValue)}.`,
      `El valor máximo alcanza ${fmtCompactCurrency(dataset.kpis.maxValue)}, útil para definir umbrales de revisión reforzada.`,
      `La mezcla de gasto ayuda a priorizar controles sobre soportes, topes y criterios de aprobación.`,
    ],
    COLORS.amberSoft
  );
}

function addOperationalSlide(prs, dataset) {
  const slide = prs.addSlide();
  applyTransition(slide, 'push', 'fast');
  const sectionTitle = 'Destinos y centros con mayor concentración';
  const sectionSubtitle = 'La dispersión territorial y presupuestal muestra dónde se concentra la demanda y dónde conviene reforzar monitoreo.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.3);
  addHeader(
    slide,
    '03',
    'HUELLA OPERATIVA',
    sectionTitle,
    sectionSubtitle
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: topY,
    w: 6.2,
    h: 4.7,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Top ciudades de destino', {
    x: 0.82,
    y: topY + 0.22,
    w: 2.5,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: 'Calibri',
  });
  slide.addChart(
    prs.ChartType.bar,
    [
      {
        name: 'Comisiones',
        labels: dataset.charts.topCities.map(([label]) => cleanText(label, 28)),
        values: dataset.charts.topCities.map(([, value]) => value),
      },
    ],
    {
      x: 0.85,
      y: topY + 0.55,
      w: 5.55,
      h: 3.7,
      barDir: 'bar',
      chartColors: [COLORS.aqua],
      showValue: true,
      dataLabelColor: COLORS.ink,
      valAxisLabelColor: COLORS.muted,
      catAxisLabelColor: COLORS.muted,
      showLegend: false,
      valGridLine: { color: 'D9E3EE' },
    }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.95,
    y: topY,
    w: 5.83,
    h: 4.7,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Top centros de costos por valor', {
    x: 7.22,
    y: topY + 0.22,
    w: 3.1,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: 'Calibri',
  });
  slide.addChart(
    prs.ChartType.bar,
    [
      {
        name: 'Valor',
        labels: dataset.charts.topCostCenters.map(([label]) => cleanText(label, 18)),
        values: dataset.charts.topCostCenters.map(([, value]) => value),
      },
    ],
    {
      x: 7.25,
      y: topY + 0.55,
      w: 5.2,
      h: 3.7,
      barDir: 'bar',
      chartColors: [COLORS.green],
      showValue: true,
      dataLabelColor: COLORS.ink,
      valAxisLabelColor: COLORS.muted,
      catAxisLabelColor: COLORS.muted,
      valAxisNumFmt: '$#,##0',
      showLegend: false,
      valGridLine: { color: 'D9E3EE' },
    }
  );
}

function addTopRequestersSlide(prs, dataset) {
  const slide = prs.addSlide();
  applyTransition(slide, 'fade', 'fast');
  const sectionTitle = 'Colaboradores con mayor valor gestionado';
  const sectionSubtitle = 'La combinación entre monto total y número de comisiones permite enfocar seguimiento en usuarios de mayor exposición.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.3);
  addHeader(
    slide,
    '04',
    'CONCENTRACIÓN POR SOLICITANTE',
    sectionTitle,
    sectionSubtitle
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: topY,
    w: 7.25,
    h: 4.65,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Top 8 por valor total solicitado', {
    x: 0.82,
    y: topY + 0.23,
    w: 3.4,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: 'Calibri',
  });
  slide.addChart(
    prs.ChartType.bar,
    [
      {
        name: 'Valor',
        labels: dataset.charts.topRequesters.map((item) => cleanText(item.name, 26)),
        values: dataset.charts.topRequesters.map((item) => item.value),
      },
    ],
    {
      x: 0.82,
      y: topY + 0.53,
      w: 6.55,
      h: 3.82,
      barDir: 'bar',
      chartColors: [COLORS.navy2],
      showValue: true,
      dataLabelColor: COLORS.ink,
      valAxisNumFmt: '$#,##0',
      valAxisLabelColor: COLORS.muted,
      catAxisLabelColor: COLORS.muted,
      showLegend: false,
      valGridLine: { color: 'D9E3EE' },
    }
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.02,
    y: topY,
    w: 4.76,
    h: 4.65,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Detalle de foco', {
    x: 8.28,
    y: topY + 0.23,
    w: 1.8,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: 'Calibri',
  });

  dataset.charts.topRequesters.slice(0, 5).forEach((item, index) => {
    const y = topY + 0.6 + index * 0.73;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 8.25,
      y,
      w: 4.28,
      h: 0.56,
      rectRadius: 0.05,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: index === 0 ? COLORS.greenSoft : 'F9FBFD' },
    });
    slide.addText(cleanText(item.name, 30), {
      x: 8.45,
      y: y + 0.08,
      w: 2.1,
      h: 0.18,
      fontSize: 8.8,
      bold: true,
      color: COLORS.ink,
      fontFace: 'Calibri',
    });
    slide.addText(`${fmtCompactCurrency(item.value)} · ${fmtNumber(item.count)} comisiones`, {
      x: 8.45,
      y: y + 0.28,
      w: 3.6,
      h: 0.16,
      fontSize: 8.3,
      color: COLORS.muted,
      fontFace: 'Calibri',
    });
  });

  addInsightBox(
    slide,
    8.25,
    Math.max(6.0, topY + 4.05),
    4.28,
    0.6,
    'Claves',
    [
      `${cleanText(dataset.charts.topRequesters[0].name, 25)} ocupa el primer lugar por valor agregado.`,
      'La lectura conjunta de valor y frecuencia ayuda a diferenciar casos intensivos de casos recurrentes.',
    ],
    COLORS.aquaSoft
  );
}

function addSampleTableSlide(prs, dataset) {
  const baseTitle = 'Comisiones con mayor valor solicitado';
  const columns = [
    { header: 'ID', width: 0.8, buildCell: (row) => ({ text: String(row.id), align: 'center', fontSize: 9.0 }) },
    { header: 'Solicitante', width: 2.5, buildCell: (row) => ({ text: row.solicitante, fontSize: 9.0 }) },
    { header: 'Ciudad destino', width: 2.65, buildCell: (row) => ({ text: row.ciudad, fontSize: 9.0 }) },
    { header: 'Valor total', width: 1.55, buildCell: (row) => ({ text: row.valor, align: 'right', fontSize: 9.0 }) },
    {
      header: 'Estado',
      width: 2.35,
      buildCell: (row) => {
        const statusStyle = getStatusStyle(row.estado);
        return { text: row.estado, align: 'center', fill: statusStyle.fill, color: statusStyle.color, bold: true, fontSize: 8.7 };
      },
    },
    { header: 'Centro', width: 1.25, buildCell: (row) => ({ text: row.centro, align: 'center', fontSize: 9.0 }) },
  ];
  const tableWidth = sumNumericValues(columns.map((column) => column.width));
  const fullRows = dataset.tables.sampleRows.map((row, rowIndex) =>
    columns.map((column) => ({
      fill: rowIndex % 2 === 0 ? 'F8FAFC' : COLORS.white,
      color: COLORS.ink,
      ...column.buildCell(row, rowIndex),
    }))
  );
  const fullWidths = columns.map((column) => column.width);
  const fullFit = findFittingManualRowLimit(fullRows, fullWidths, {
    preferredLimit: 5,
    minLimit: 2,
    availableHeight: 4.18,
    baseFontSize: 8.9,
    minRowHeight: 0.76,
    maxRowHeight: 1.12,
    headerHeight: 0.42,
    maxLinesPerCell: 4.8,
    maxAverageLines: 2.9,
  });
  const summaryIndexes = [0, 1, 3, 4];
  const summaryRows = projectManualRows(fullRows, summaryIndexes);
  const summaryWidths = projectColumnWidths(fullWidths, summaryIndexes, tableWidth);
  const summaryFit = findFittingManualRowLimit(summaryRows, summaryWidths, {
    preferredLimit: 6,
    minLimit: 3,
    availableHeight: 4.18,
    baseFontSize: 9.0,
    minRowHeight: 0.72,
    maxRowHeight: 1.02,
    headerHeight: 0.42,
    maxLinesPerCell: 4.4,
    maxAverageLines: 2.6,
  });
  const useSummaryDetail = fullFit.rowLimit < Math.min(5, fullRows.length) || fullFit.metrics.maxLines >= 5 || fullFit.metrics.avgLines >= 3;
  let pageCounter = 0;

  function renderSampleTablePage(chunkRows, widths, headers, rowHeights, sectionTitle, sectionSubtitle, densityLabel) {
    const slide = prs.addSlide();
    applyTransition(slide, 'wipe', 'med');
    const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.28);
    addHeader(
      slide,
      formatSectionPageNumber('05', pageCounter),
      'MUESTRA DE REGISTROS RELEVANTES',
      sectionTitle,
      sectionSubtitle
    );
    pageCounter += 1;

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.55,
      y: topY,
      w: 12.23,
      h: 4.95,
      rectRadius: 0.08,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: COLORS.card },
    });
    slide.addText(densityLabel, {
      x: 9.15,
      y: topY + 0.09,
      w: 3.1,
      h: 0.18,
      fontSize: 8.4,
      bold: true,
      color: COLORS.muted,
      align: 'right',
      fontFace: getDeckFont('body'),
    });
    renderManualTable(slide, {
      x: 0.78,
      y: topY + 0.29,
      colWidths: widths,
      headerHeight: 0.42,
      rowHeight: 0.78,
      rowHeights,
      defaultFontSize: 8.9,
      minFontSize: 8.2,
      headers: headers.map((header) => ({ text: header, fontSize: 9.2 })),
      rows: chunkRows,
    });
  }

  if (useSummaryDetail) {
    const summaryTitle = `${baseTitle} | vista resumida`;
    const summarySubtitle = 'La densidad obligó a priorizar lectura ejecutiva: primero se muestran las columnas decisivas y luego el detalle completo.';
    const summaryChunks = chunkArray(summaryRows, Math.max(3, summaryFit.rowLimit || 3));
    summaryChunks.forEach((chunk) => {
      const metrics = evaluateManualTableDensity(chunk, summaryWidths, {
        availableHeight: 4.18,
        baseFontSize: 9.0,
        minRowHeight: 0.72,
        maxRowHeight: 1.02,
        headerHeight: 0.42,
        maxLinesPerCell: 4.4,
        maxAverageLines: 2.6,
      });
      renderSampleTablePage(
        chunk,
        summaryWidths,
        summaryIndexes.map((columnIndex) => columns[columnIndex].header),
        metrics.rowHeights,
        summaryTitle,
        summarySubtitle,
        'Resumen automatico'
      );
    });

    const detailTitle = `${baseTitle} | detalle completo`;
    const detailSubtitle = 'Se libera una segunda capa para conservar todas las columnas sin microtexto ni compresion visual.';
    const detailChunks = chunkArray(fullRows, Math.max(2, fullFit.rowLimit || 2));
    detailChunks.forEach((chunk) => {
      const metrics = evaluateManualTableDensity(chunk, fullWidths, {
        availableHeight: 4.18,
        baseFontSize: 8.9,
        minRowHeight: 0.76,
        maxRowHeight: 1.12,
        headerHeight: 0.42,
        maxLinesPerCell: 4.8,
        maxAverageLines: 2.9,
      });
      renderSampleTablePage(
        chunk,
        fullWidths,
        columns.map((column) => column.header),
        metrics.rowHeights,
        detailTitle,
        detailSubtitle,
        'Detalle automatico'
      );
    });
    return;
  }

  const sectionSubtitle = 'La tabla mantiene proporciones homogéneas y prioriza filas representativas para lectura ejecutiva sin perder contexto.';
  const fullChunks = chunkArray(fullRows, Math.max(2, fullFit.rowLimit || 2));
  fullChunks.forEach((chunk) => {
    const metrics = evaluateManualTableDensity(chunk, fullWidths, {
      availableHeight: 4.18,
      baseFontSize: 8.9,
      minRowHeight: 0.76,
      maxRowHeight: 1.12,
      headerHeight: 0.42,
      maxLinesPerCell: 4.8,
      maxAverageLines: 2.9,
    });
    renderSampleTablePage(
      chunk,
      fullWidths,
      columns.map((column) => column.header),
      metrics.rowHeights,
      baseTitle,
      sectionSubtitle,
      'Densidad estable'
    );
  });
}

function addCosoSlide(prs, dataset) {
  const slide = prs.addSlide();
  applyTransition(slide, 'fade', 'med');
  const sectionTitle = 'Lectura del modelo COSO';
  const sectionSubtitle = 'El modelo evidencia fortalezas puntuales, pero mantiene frentes claros de oportunidad en ambiente, información y monitoreo.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.3);
  addHeader(
    slide,
    '06',
    'MADUREZ DE CONTROL INTERNO',
    sectionTitle,
    sectionSubtitle
  );

  dataset.tables.coso.slice(0, 4).forEach((item, index) => {
    const x = 0.55 + (index % 2) * 6.15;
    const y = topY + Math.floor(index / 2) * 2.18;
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w: 5.95,
      h: 2.02,
      rectRadius: 0.08,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: COLORS.card },
    });
    slide.addText(item.component, {
      x: x + 0.22,
      y: y + 0.18,
      w: 5.1,
      h: 0.24,
      fontSize: 12.5,
      bold: true,
      color: COLORS.ink,
      fontFace: getDeckFont('head'),
      fit: 'shrink',
    });
    addCompactStatCard(slide, x + 0.22, y + 0.54, 1.45, 0.92, 'Oportunidades', fmtNumber(item.opportunities), COLORS.amber);
    addCompactStatCard(slide, x + 1.82, y + 0.54, 1.45, 0.92, 'Efectivos', fmtNumber(item.effective), COLORS.green);
    addCompactStatCard(slide, x + 3.42, y + 0.54, 1.45, 0.92, 'No efectivos', fmtNumber(item.ineffective), COLORS.red);
    slide.addText(cleanText(item.items, 180), {
      x: x + 0.22,
      y: y + 1.56,
      w: 5.35,
      h: 0.28,
      fontSize: 9.2,
      color: COLORS.muted,
      fontFace: getDeckFont('body'),
      italic: true,
      fit: 'shrink',
    });
  });
}

function addFindingsSlide(prs, dataset) {
  const chunks = chunkArray(dataset.tables.hallazgos.slice(0, 6), 1);

  chunks.forEach((chunk, chunkIndex) => {
    const slide = prs.addSlide();
    applyTransition(slide, 'push', 'med');
    const sectionTitle = `Riesgos observados y foco correctivo ${chunks.length > 1 ? `(${chunkIndex + 1}/${chunks.length})` : ''}`.trim();
    const sectionSubtitle = 'Los hallazgos se conectan con trazabilidad, topes, soportes, tiempos de legalización y suficiencia documental.';
    const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.32);
    addHeader(
      slide,
      chunkIndex === 0 ? '07' : `07${String.fromCharCode(65 + chunkIndex)}`,
      'HALLAZGOS RELEVANTES',
      sectionTitle,
      sectionSubtitle
    );

    chunk.forEach((item, index) => {
      const x = 0.55;
      const y = topY + index * 2.35;
      const statusStyle = getStatusStyle(item.status);
      slide.addShape(pptx.ShapeType.roundRect, {
        x,
        y,
        w: 12.2,
        h: 4.05,
        rectRadius: 0.08,
        line: { color: COLORS.line, pt: 1 },
        fill: { color: COLORS.card },
      });
      slide.addShape(pptx.ShapeType.rect, {
        x,
        y,
        w: 0.12,
        h: 4.05,
        line: { color: COLORS.red, transparency: 100 },
        fill: { color: COLORS.red },
      });
      slide.addText(cleanText(item.finding), {
        x: x + 0.24,
        y: y + 0.18,
        w: 9.15,
        h: 0.42,
        fontSize: 13.2,
        bold: true,
        color: COLORS.ink,
        fontFace: getDeckFont('head'),
        fit: 'shrink',
      });
      slide.addShape(pptx.ShapeType.roundRect, {
        x: x + 10.18,
        y: y + 0.16,
        w: 1.92,
        h: 0.36,
        rectRadius: 0.05,
        line: { color: statusStyle.fill, transparency: 100 },
        fill: { color: statusStyle.fill },
      });
      slide.addText(item.status || 'Sin estado', {
        x: x + 10.34,
        y: y + 0.26,
        w: 1.58,
        h: 0.14,
        fontSize: 9,
        bold: true,
        color: statusStyle.color,
        align: 'center',
        fontFace: getDeckFont('body'),
        fit: 'shrink',
      });
      slide.addText(cleanText(item.risk), {
        x: x + 0.24,
        y: y + 0.78,
        w: 11.55,
        h: 0.72,
        fontSize: 10.8,
        color: COLORS.muted,
        fontFace: getDeckFont('body'),
        fit: 'shrink',
      });
      const actions = splitActionLines(item.action);
      actions.slice(0, 4).forEach((action, actionIndex) => {
        slide.addText(`• ${action}`, {
          x: x + 0.26,
          y: y + 1.72 + actionIndex * 0.42,
          w: 11.4,
          h: 0.3,
          fontSize: 9.6,
          color: COLORS.ink,
          fontFace: getDeckFont('body'),
          fit: 'shrink',
        });
      });
    });
  });
}

function addOpportunityDashboardSlide(prs, dataset) {
  const slide = prs.addSlide();
  applyTransition(slide, 'fade', 'med');
  const sectionTitle = 'Estado de avance de oportunidades correctivas';
  const sectionSubtitle = 'Se consolida el progreso de cierre para identificar pendientes estructurales y acciones ya implementadas.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.85, 0.26);
  addHeader(
    slide,
    '08',
    'PLAN DE MEJORA',
    sectionTitle,
    sectionSubtitle
  );

  [
    ['Total acciones', fmtNumber(dataset.opportunityMetrics.total), COLORS.aqua],
    ['Completadas', fmtNumber(dataset.opportunityMetrics.completed), COLORS.green],
    ['En ejecución', fmtNumber(dataset.opportunityMetrics.inProgress), COLORS.amber],
    ['Pendientes', fmtNumber(dataset.opportunityMetrics.pending), COLORS.red],
    ['Avance promedio', fmtPercent(dataset.opportunityMetrics.average), COLORS.navy2],
  ].forEach(([label, value, accent], index) => {
    addMetricCard(slide, 0.55 + index * 2.52, topY, 2.32, 1.08, label, value, accent);
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: Math.max(3.2, topY + 1.35),
    w: 5.4,
    h: 3.1,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Distribución por estado', {
    x: 0.82,
    y: Math.max(3.45, topY + 1.6),
    w: 2.6,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: 'Calibri',
  });
  slide.addChart(
    prs.ChartType.doughnut,
    [
      {
        name: 'Estado',
        labels: dataset.opportunityStatus.map(([label]) => cleanText(label, 18)),
        values: dataset.opportunityStatus.map(([, value]) => value),
      },
    ],
    {
      x: 0.82,
      y: Math.max(3.82, topY + 1.97),
      w: 2.2,
      h: 2.05,
      holeSize: 62,
      showLegend: true,
      legendPos: 'r',
      legendFontSize: 8.5,
      legendFontColor: COLORS.muted,
      chartColors: [COLORS.green, COLORS.amber, COLORS.red, COLORS.aqua],
      showValue: true,
      dataLabelColor: COLORS.ink,
      dataLabelFontSize: 8,
    }
  );

  addInsightBox(
    slide,
    3.2,
    Math.max(3.48, topY + 1.63),
    2.3,
    2.1,
    'Lectura',
    [
      `El avance promedio de cierre es ${fmtPercent(dataset.opportunityMetrics.average)}.`,
      dataset.opportunityMetrics.pending
        ? `${fmtNumber(dataset.opportunityMetrics.pending)} acciones siguen pendientes y requieren fecha compromiso.`
        : 'No se observan acciones pendientes.',
    ],
    COLORS.greenSoft
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.2,
    y: Math.max(3.2, topY + 1.35),
    w: 6.58,
    h: 3.1,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Prioridades inmediatas', {
    x: 6.47,
    y: Math.max(3.45, topY + 1.6),
    w: 3.1,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: getDeckFont('head'),
  });

  dataset.tables.priorityOpportunities.slice(0, 3).forEach((item, index) => {
    const y = Math.max(3.9, topY + 2.05) + index * 0.78;
    const style = getStatusStyle(item.status);
    slide.addText(cleanText(item.opportunity), {
      x: 6.47,
      y,
      w: 4.35,
      h: 0.34,
      fontSize: 8.9,
      color: COLORS.ink,
      fontFace: getDeckFont('body'),
      fit: 'shrink',
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 11.0,
      y: y + 0.07,
      w: 1.1,
      h: 0.18,
      rectRadius: 0.03,
      line: { color: 'DDE5EF', transparency: 100 },
      fill: { color: 'EDF2F8' },
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 11.0,
      y: y + 0.07,
      w: Math.max(0.08, item.progress) * 1.1,
      h: 0.18,
      rectRadius: 0.03,
      line: { color: style.color, transparency: 100 },
      fill: { color: style.color },
    });
    slide.addText(fmtPercent(item.progress, 0), {
      x: 12.18,
      y: y + 0.03,
      w: 0.38,
      h: 0.18,
      fontSize: 8.4,
      bold: true,
      color: COLORS.ink,
      align: 'right',
      fontFace: getDeckFont('body'),
    });
  });
}

function addOpportunityTableSlide(prs, dataset) {
  const columns = [
    { header: 'Proceso / área', width: 1.45, buildCell: (item) => ({ text: item.area, fontSize: 8.3 }) },
    { header: 'Oportunidad de mejora', width: 4.45, buildCell: (item) => ({ text: item.opportunity, fontSize: 8.3 }) },
    {
      header: 'Estado',
      width: 1.15,
      buildCell: (item) => {
        const style = getStatusStyle(item.status);
        return { text: item.status || 'Sin estado', align: 'center', bold: true, color: style.color, fill: style.fill, fontSize: 8.1 };
      },
    },
    { header: 'Avance', width: 0.85, buildCell: (item) => ({ text: fmtPercent(item.progress, 0), align: 'center', fontSize: 8.1 }) },
    { header: 'Observación clave', width: 3.88, buildCell: (item) => ({ text: item.observation, fontSize: 8.3 }) },
  ];
  const fullWidths = columns.map((column) => column.width);
  const tableWidth = sumNumericValues(fullWidths);
  const fullRows = dataset.tables.priorityOpportunities.map((item, rowIndex) =>
    columns.map((column) => ({
      fill: rowIndex % 2 === 0 ? 'F8FAFC' : COLORS.white,
      color: COLORS.ink,
      ...column.buildCell(item, rowIndex),
    }))
  );
  const fullFit = findFittingManualRowLimit(fullRows, fullWidths, {
    preferredLimit: 2,
    minLimit: 1,
    availableHeight: 4.18,
    baseFontSize: 8.2,
    minRowHeight: 0.94,
    maxRowHeight: 1.32,
    headerHeight: 0.42,
    maxLinesPerCell: 5.1,
    maxAverageLines: 3.1,
  });
  const summaryIndexes = [0, 1, 2, 3];
  const summaryRows = projectManualRows(fullRows, summaryIndexes);
  const summaryWidths = projectColumnWidths(fullWidths, summaryIndexes, tableWidth);
  const summaryFit = findFittingManualRowLimit(summaryRows, summaryWidths, {
    preferredLimit: 3,
    minLimit: 2,
    availableHeight: 4.18,
    baseFontSize: 8.3,
    minRowHeight: 0.82,
    maxRowHeight: 1.04,
    headerHeight: 0.42,
    maxLinesPerCell: 4.6,
    maxAverageLines: 2.8,
  });
  const useSummaryDetail = fullFit.rowLimit < Math.min(2, fullRows.length) || fullFit.metrics.maxLines >= 5 || fullFit.metrics.avgLines >= 3.15;
  let pageCounter = 0;

  function renderOpportunityPage(chunkRows, widths, headers, rowHeights, sectionTitle, sectionSubtitle, densityLabel) {
    const slide = prs.addSlide();
    applyTransition(slide, 'push', 'fast');
    const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.26);
    addHeader(
      slide,
      formatSectionPageNumber('09', pageCounter),
      'DETALLE DE OPORTUNIDADES',
      sectionTitle,
      sectionSubtitle
    );
    pageCounter += 1;

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.55,
      y: topY,
      w: 12.23,
      h: 4.95,
      rectRadius: 0.08,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: COLORS.card },
    });
    slide.addText(densityLabel, {
      x: 9.05,
      y: topY + 0.09,
      w: 3.25,
      h: 0.18,
      fontSize: 8.4,
      bold: true,
      color: COLORS.muted,
      align: 'right',
      fontFace: getDeckFont('body'),
    });
    renderManualTable(slide, {
      x: 0.78,
      y: topY + 0.29,
      colWidths: widths,
      headerHeight: 0.42,
      rowHeight: 1.0,
      rowHeights,
      defaultFontSize: 8.2,
      minFontSize: 8.0,
      headers: headers.map((header) => ({ text: header, fontSize: 9.2 })),
      rows: chunkRows,
    });
  }

  if (useSummaryDetail) {
    const summaryTitle = 'Acciones con menor avance o mayor criticidad | resumen';
    const summarySubtitle = 'Cuando la observación vuelve pesada la tabla, el motor compacta la primera vista y reserva el detalle completo en slides posteriores.';
    const summaryChunks = chunkArray(summaryRows, Math.max(2, summaryFit.rowLimit || 2));
    summaryChunks.forEach((chunk) => {
      const metrics = evaluateManualTableDensity(chunk, summaryWidths, {
        availableHeight: 4.18,
        baseFontSize: 8.3,
        minRowHeight: 0.82,
        maxRowHeight: 1.04,
        headerHeight: 0.42,
        maxLinesPerCell: 4.6,
        maxAverageLines: 2.8,
      });
      renderOpportunityPage(
        chunk,
        summaryWidths,
        summaryIndexes.map((columnIndex) => columns[columnIndex].header),
        metrics.rowHeights,
        summaryTitle,
        summarySubtitle,
        'Resumen automatico'
      );
    });

    const detailTitle = 'Acciones con menor avance o mayor criticidad | detalle';
    const detailSubtitle = 'La observación completa pasa a una segunda capa para evitar compresión visual y preservar lectura íntegra.';
    const detailChunks = chunkArray(fullRows, Math.max(1, fullFit.rowLimit || 1));
    detailChunks.forEach((chunk) => {
      const metrics = evaluateManualTableDensity(chunk, fullWidths, {
        availableHeight: 4.18,
        baseFontSize: 8.2,
        minRowHeight: 0.94,
        maxRowHeight: 1.32,
        headerHeight: 0.42,
        maxLinesPerCell: 5.1,
        maxAverageLines: 3.1,
      });
      renderOpportunityPage(
        chunk,
        fullWidths,
        columns.map((column) => column.header),
        metrics.rowHeights,
        detailTitle,
        detailSubtitle,
        'Detalle automatico'
      );
    });
    return;
  }

  const sectionTitle = 'Acciones con menor avance o mayor criticidad';
  const sectionSubtitle = 'Se priorizan frentes de mejora con bajo porcentaje de ejecución para acelerar cierre y trazabilidad.';
  const chunks = chunkArray(fullRows, Math.max(1, fullFit.rowLimit || 1));
  chunks.forEach((chunk) => {
    const metrics = evaluateManualTableDensity(chunk, fullWidths, {
      availableHeight: 4.18,
      baseFontSize: 8.2,
      minRowHeight: 0.94,
      maxRowHeight: 1.32,
      headerHeight: 0.42,
      maxLinesPerCell: 5.1,
      maxAverageLines: 3.1,
    });
    renderOpportunityPage(
      chunk,
      fullWidths,
      columns.map((column) => column.header),
      metrics.rowHeights,
      sectionTitle,
      sectionSubtitle,
      'Densidad estable'
    );
  });
}

function addHallazgosTableSlide(prs, dataset) {
  const columns = [
    { header: 'Hallazgo', width: 2.75, buildCell: (item) => ({ text: item.finding, fontSize: 8.2 }) },
    { header: 'Riesgo asociado', width: 3.55, buildCell: (item) => ({ text: item.risk, fontSize: 8.1 }) },
    {
      header: 'Estado',
      width: 1.05,
      buildCell: (item) => {
        const style = getStatusStyle(item.status);
        return { text: item.status || 'Sin estado', align: 'center', bold: true, color: style.color, fill: style.fill, fontSize: 8.0 };
      },
    },
    { header: 'Acción correctiva', width: 4.43, buildCell: (item) => ({ text: item.action, fontSize: 8.1 }) },
  ];
  const fullWidths = columns.map((column) => column.width);
  const tableWidth = sumNumericValues(fullWidths);
  const fullRows = dataset.tables.hallazgos.map((item, rowIndex) =>
    columns.map((column) => ({
      fill: rowIndex % 2 === 0 ? 'F8FAFC' : COLORS.white,
      color: COLORS.ink,
      ...column.buildCell(item, rowIndex),
    }))
  );
  const fullFit = findFittingManualRowLimit(fullRows, fullWidths, {
    preferredLimit: 1,
    minLimit: 1,
    availableHeight: 4.18,
    baseFontSize: 8.1,
    minRowHeight: 1.18,
    maxRowHeight: 1.72,
    headerHeight: 0.42,
    maxLinesPerCell: 5.6,
    maxAverageLines: 3.35,
  });
  const summaryIndexes = [0, 2, 3];
  const summaryRows = projectManualRows(fullRows, summaryIndexes);
  const summaryWidths = projectColumnWidths(fullWidths, summaryIndexes, tableWidth);
  const summaryFit = findFittingManualRowLimit(summaryRows, summaryWidths, {
    preferredLimit: 2,
    minLimit: 1,
    availableHeight: 4.18,
    baseFontSize: 8.2,
    minRowHeight: 0.96,
    maxRowHeight: 1.28,
    headerHeight: 0.42,
    maxLinesPerCell: 4.9,
    maxAverageLines: 2.95,
  });
  const useSummaryView = fullFit.metrics.maxLines >= 6 || fullFit.metrics.avgLines >= 3.5 || fullFit.metrics.totalHeight > 4.2;
  const sectionTitle = useSummaryView ? 'Matriz resumida de cierre y estado' : 'Resumen de riesgo y acción correctiva';
  const sectionSubtitle = useSummaryView
    ? 'La densidad obliga a resumir la matriz; el detalle completo queda expandido en la sección previa de hallazgos narrativos.'
    : 'Esta vista conecta hallazgo, riesgo asociado y acción definida para facilitar seguimiento de cierre.';
  const activeRows = useSummaryView ? summaryRows : fullRows;
  const activeWidths = useSummaryView ? summaryWidths : fullWidths;
  const activeHeaders = (useSummaryView ? summaryIndexes : columns.map((_, index) => index)).map((columnIndex) => columns[columnIndex].header);
  const chunkSize = useSummaryView ? Math.max(1, summaryFit.rowLimit || 1) : 1;
  const chunks = chunkArray(activeRows, chunkSize);

  chunks.forEach((chunk, chunkIndex) => {
    const slide = prs.addSlide();
    applyTransition(slide, 'fade', 'fast');
    const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.28);
    addHeader(
      slide,
      formatSectionPageNumber('10', chunkIndex),
      'TRAZABILIDAD DE HALLAZGOS',
      sectionTitle,
      sectionSubtitle
    );

    slide.addShape(pptx.ShapeType.roundRect, {
      x: 0.55,
      y: topY,
      w: 12.23,
      h: 4.95,
      rectRadius: 0.08,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: COLORS.card },
    });
    slide.addText(useSummaryView ? 'Resumen automatico' : 'Detalle completo', {
      x: 9.05,
      y: topY + 0.09,
      w: 3.25,
      h: 0.18,
      fontSize: 8.4,
      bold: true,
      color: COLORS.muted,
      align: 'right',
      fontFace: getDeckFont('body'),
    });
    const metrics = evaluateManualTableDensity(chunk, activeWidths, {
      availableHeight: 4.18,
      baseFontSize: useSummaryView ? 8.2 : 8.1,
      minRowHeight: useSummaryView ? 0.96 : 1.18,
      maxRowHeight: useSummaryView ? 1.28 : 1.72,
      headerHeight: 0.42,
      maxLinesPerCell: useSummaryView ? 4.9 : 5.6,
      maxAverageLines: useSummaryView ? 2.95 : 3.35,
    });

    renderManualTable(slide, {
      x: 0.78,
      y: topY + 0.29,
      colWidths: activeWidths,
      headerHeight: 0.42,
      rowHeight: useSummaryView ? 1.02 : 1.34,
      rowHeights: metrics.rowHeights,
      defaultFontSize: useSummaryView ? 8.2 : 8.1,
      minFontSize: 8.0,
      headers: activeHeaders.map((header) => ({ text: header, fontSize: 9.2 })),
      rows: chunk,
    });
  });
}

function addClosing(prs, dataset) {
  const slide = prs.addSlide();
  slide.background = { color: COLORS.navy };
  applyTransition(slide, 'fade', 'slow');

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: 1.1,
    line: { color: COLORS.navy2, transparency: 100 },
    fill: { color: COLORS.navy2 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 10.15,
    y: 0,
    w: 3.18,
    h: 7.5,
    line: { color: COLORS.panelAlt, transparency: 100 },
    fill: { color: COLORS.panelAlt, transparency: 18 },
  });
  addNatureDecor(slide, 'closing');
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.82,
    y: 1.55,
    w: 0.08,
    h: 3.95,
    line: { color: COLORS.white, transparency: 25 },
    fill: { color: COLORS.white, transparency: 25 },
  });
  slide.addText('Cierre Ejecutivo', {
    x: 1.22,
    y: 1.55,
    w: 7.5,
    h: 0.6,
    fontSize: 32,
    bold: true,
    color: COLORS.textLight,
    fontFace: 'Calibri',
  });
  slide.addText('La información del archivo permite priorizar controles sobre trazabilidad, legalización, topes y cierre de acciones correctivas, manteniendo visibilidad simultánea del gasto y del riesgo.', {
    x: 1.22,
    y: 2.48,
    w: 6.8,
    h: 0.9,
    fontSize: 15,
    color: COLORS.textSoft,
    fontFace: 'Calibri',
    italic: true,
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.7,
    y: 1.7,
    w: 3.55,
    h: 2.95,
    rectRadius: 0.08,
    line: { color: COLORS.white, pt: 1, transparency: 28 },
    fill: { color: COLORS.navy2, transparency: 24 },
  });
  slide.addText('Cifras de cierre', {
    x: 8.95,
    y: 1.98,
    w: 1.85,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.white,
    fontFace: 'Calibri',
  });
  [
    [`${fmtNumber(dataset.kpis.totalRows)}`, 'comisiones auditadas'],
    [`${fmtCompactCurrency(dataset.kpis.totalValue)}`, 'valor total analizado'],
    [`${fmtNumber(dataset.opportunityMetrics.total)}`, 'acciones de mejora'],
  ].forEach(([value, label], index) => {
    slide.addText(value, {
      x: 8.98,
      y: 2.42 + index * 0.7,
      w: 1.4,
      h: 0.24,
      fontSize: 18,
      bold: true,
      color: COLORS.textLight,
      fontFace: 'Calibri',
    });
    slide.addText(label, {
      x: 10.35,
      y: 2.47 + index * 0.7,
      w: 1.7,
      h: 0.22,
      fontSize: 8.8,
      color: COLORS.textSoft,
      fontFace: 'Calibri',
    });
  });
  slide.addText(
    `Resumen base: ${fmtNumber(dataset.kpis.totalRows)} comisiones · ${fmtCompactCurrency(dataset.kpis.totalValue)} · ${fmtNumber(dataset.opportunityMetrics.total)} acciones de mejora`,
    {
      x: 1.22,
      y: 4.45,
      w: 7.3,
      h: 0.25,
      fontSize: 11,
      color: COLORS.textLight,
      fontFace: 'Calibri',
    }
  );
  slide.addText('Fundación Socya · Presentación generada automáticamente con datos reales y formato optimizado para PowerPoint.', {
    x: 1.22,
    y: 6.65,
    w: 10.0,
    h: 0.2,
    fontSize: 9.5,
    color: COLORS.textSoft,
    fontFace: 'Calibri',
  });
}

function buildCommissionRecommendations(dataset) {
  const recommendations = [];
  const topState = Array.isArray(dataset?.charts?.statusCounts) ? dataset.charts.statusCounts[0] : null;
  const topCity = Array.isArray(dataset?.charts?.cities) ? dataset.charts.cities[0] : null;
  const topRequester = Array.isArray(dataset?.charts?.topRequesters) ? dataset.charts.topRequesters[0] : null;

  if (topState && cleanText(topState.label) && Number(topState.value) > 0) {
    recommendations.push(`Intervenir prioritariamente el estado ${cleanText(topState.label, 28)} por su concentración en ${fmtNumber(topState.value)} registros del corte analizado.`);
  }
  if (topRequester && cleanText(topRequester.name) && Number(topRequester.value) > 0) {
    recommendations.push(`Revisar el comportamiento del solicitante ${cleanText(topRequester.name, 32)} y validar recurrencia, soportes y controles previos asociados.`);
  }
  if (topCity && cleanText(topCity.label) && Number(topCity.value) > 0) {
    recommendations.push(`Fortalecer seguimiento territorial en ${cleanText(topCity.label, 28)} para reducir concentración y mejorar oportunidad de revisión.`);
  }
  if (dataset.opportunityMetrics.total > 0) {
    recommendations.push('Formalizar un tablero quincenal de cierre para oportunidades de mejora con responsable, fecha compromiso y evidencia de cierre.');
  }
  if (dataset.kpis.totalValue > 0) {
    recommendations.push(`Definir umbrales de revisión reforzada para solicitudes de alto monto, tomando como base un universo analizado de ${fmtCompactCurrency(dataset.kpis.totalValue)}.`);
  }
  if (!recommendations.length) {
    recommendations.push('Consolidar acciones de control, seguimiento y cierre sobre los hallazgos materializados en las tablas priorizadas del Excel.');
  }
  return recommendations.slice(0, 4);
}

function addCommissionRecommendationsSlide(prs, dataset, sectionNumber = '09') {
  const slide = prs.addSlide();
  applyTransition(slide, 'fade', 'med');
  const sectionTitle = 'Conclusiones y recomendaciones';
  const sectionSubtitle = 'Síntesis ejecutiva construida a partir de concentraciones, hallazgos y oportunidades detectadas en el archivo analizado.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.28);
  addHeader(slide, sectionNumber, 'CIERRE EJECUTIVO', sectionTitle, sectionSubtitle);

  const recommendations = buildCommissionRecommendations(dataset);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: topY,
    w: 7.45,
    h: 4.95,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Recomendaciones priorizadas', {
    x: 0.82,
    y: topY + 0.22,
    w: 6.6,
    h: 0.2,
    fontSize: 12.5,
    bold: true,
    color: COLORS.ink,
    fontFace: getDeckFont('head'),
  });
  recommendations.forEach((rec, index) => {
    slide.addText(`${index + 1}. ${rec}`, {
      x: 0.82,
      y: topY + 0.58 + index * 0.5,
      w: 7.0,
      h: 0.38,
      fontSize: 10.6,
      color: COLORS.ink,
      fontFace: getDeckFont('body'),
      fit: 'shrink',
    });
  });

  const highlights = [
    `Total analizado: ${fmtNumber(dataset.kpis.totalRows)} comisiones`,
    `Valor identificado: ${fmtCurrency(dataset.kpis.totalValue)}`,
    dataset.opportunityMetrics.total > 0
      ? `Oportunidades activas: ${fmtNumber(dataset.opportunityMetrics.total)}`
      : 'No se detectaron oportunidades estructuradas en el corte actual.',
  ];
  addInsightBox(slide, 8.25, topY, 5.0, 2.35, 'Mensajes clave', highlights, getSoftAccentColor(getDeckAccent(0)));

  const requesters = Array.isArray(dataset?.charts?.topRequesters) ? dataset.charts.topRequesters.slice(0, 4) : [];
  const cityFocus = Array.isArray(dataset?.charts?.cities) ? dataset.charts.cities.slice(0, 4) : [];
  const focusItems = requesters.length
    ? requesters
        .filter((row) => cleanText(row.name) && Number(row.value) > 0)
        .map((row) => `• ${cleanText(row.name, 28)} · ${fmtCurrency(row.value)} · ${fmtNumber(row.count)} casos`)
    : cityFocus
        .filter((row) => cleanText(row.label) && Number(row.value) > 0)
        .map((row) => `• ${cleanText(row.label, 28)} · ${fmtNumber(row.value)} comisiones`);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.25,
    y: topY + 2.52,
    w: 5.0,
    h: 2.43,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Focos de seguimiento', {
    x: 8.48,
    y: topY + 2.72,
    w: 4.55,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: getDeckFont('head'),
  });
  focusItems.slice(0, 4).forEach((text, index) => {
    slide.addText(text, {
      x: 8.5,
      y: topY + 3.08 + index * 0.42,
      w: 4.65,
      h: 0.28,
      fontSize: 9.1,
      color: COLORS.muted,
      fontFace: getDeckFont('body'),
      fit: 'shrink',
    });
  });
}

function addGenericCover(prs, dataset) {
  const slide = prs.addSlide();
  slide.background = { color: COLORS.navy };
  applyTransition(slide, 'fade', 'slow');
  addNatureDecor(slide, 'cover');
  const coverVariant = getDeckVariant('genericCover');
  const headingFont = getDeckFont('head');
  const bodyFont = getDeckFont('body');
  const headlineCards = dataset.headlineCards.slice(0, 3).length ? dataset.headlineCards.slice(0, 3) : [
    { label: 'Indicador principal', value: dataset.featuredSheets[0]?.headlineMetrics?.[0]?.value || 'N/D' },
    { label: 'Categoria lider', value: dataset.featuredSheets[0]?.headlineMetrics?.[2]?.value || 'N/D' },
    { label: 'Hoja foco', value: dataset.featuredSheets[0]?.name || 'N/D' },
  ];

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 13.33,
    h: coverVariant === 1 ? 1.08 : 7.5,
    line: { color: DECK_STYLE.topBandAccent, transparency: 100 },
    fill: { color: DECK_STYLE.topBandColor },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: coverVariant === 2 ? 8.55 : 9.0,
    y: 0,
    w: coverVariant === 1 ? 4.33 : coverVariant === 2 ? 4.78 : 4.33,
    h: 7.5,
    line: { color: DECK_STYLE.sidePanelColor, transparency: 100 },
    fill: { color: DECK_STYLE.sidePanelColor, transparency: DECK_STYLE.sidePanelTransparency },
  });
  slide.addText('PRESENTACION EJECUTIVA DE DATOS', {
    x: 0.9,
    y: coverVariant === 1 ? 1.42 : 0.9,
    w: 3.9,
    h: 0.22,
    fontSize: 12,
    bold: true,
    color: COLORS.white,
    fontFace: bodyFont,
  });
  slide.addText(dataset.workbookTitle || 'Analisis de Excel', {
    x: 0.9,
    y: coverVariant === 1 ? 2.0 : 1.55,
    w: coverVariant === 2 ? 5.8 : 6.6,
    h: coverVariant === 1 ? 0.9 : 1.0,
    fontSize: 28,
    bold: true,
    color: COLORS.white,
    fontFace: headingFont,
  });
  slide.addText(`Motor universal de PowerPoint · archivo ${SOURCE_FILE_LABEL}`, {
    x: 0.9,
    y: coverVariant === 1 ? 2.95 : 2.7,
    w: 5.2,
    h: 0.24,
    fontSize: 13,
    italic: true,
    color: COLORS.textSoft,
    fontFace: bodyFont,
  });

  if (coverVariant === 1) {
    addInsightBox(slide, 9.25, 1.55, 3.35, 4.7, 'Mensajes clave', dataset.insights.slice(0, 3), getSoftAccentColor(getDeckAccent(1)));
    headlineCards.forEach((card, index) => {
      addMetricCard(slide, 0.95 + index * 2.8, 4.65, 2.55, 1.02, card.label, card.value, getDeckAccent(index), card.sheet || card.note);
    });
  } else if (coverVariant === 2) {
    dataset.insights.slice(0, 3).forEach((line, index) => {
      slide.addText(`${index + 1}. ${line}`, {
        x: 0.95,
        y: 3.28 + index * 0.54,
        w: 5.9,
        h: 0.28,
        fontSize: 10.8,
        color: COLORS.textSoft,
        fontFace: bodyFont,
        fit: 'shrink',
      });
    });
    headlineCards.forEach((card, index) => {
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 8.95,
        y: 1.62 + index * 1.0,
        w: 3.25,
        h: 0.72,
        rectRadius: 0.07,
        line: { color: COLORS.white, pt: 1, transparency: 28 },
        fill: { color: DECK_STYLE.topBandAccent, transparency: 24 },
      });
      slide.addText(card.label, {
        x: 9.18,
        y: 1.8 + index * 1.0,
        w: 1.45,
        h: 0.16,
        fontSize: 8.6,
        color: COLORS.textSoft,
        fontFace: bodyFont,
      });
      slide.addText(card.value, {
        x: 10.55,
        y: 1.76 + index * 1.0,
        w: 1.35,
        h: 0.18,
        fontSize: 14,
        bold: true,
        color: COLORS.white,
        align: 'right',
        fontFace: headingFont,
      });
    });
    addInsightBox(slide, 0.95, 5.0, 5.75, 1.25, 'Lectura inicial', dataset.insights.slice(0, 2), getSoftAccentColor(getDeckAccent(0)));
  } else {
    dataset.insights.slice(0, 3).forEach((line, index) => {
      slide.addText(`${index + 1}. ${line}`, {
        x: 0.95,
        y: 3.3 + index * 0.52,
        w: 6.1,
        h: 0.24,
        fontSize: 11,
        color: COLORS.textSoft,
        fontFace: bodyFont,
      });
    });
    headlineCards.forEach((card, index) => {
      slide.addShape(pptx.ShapeType.roundRect, {
        x: 9.55,
        y: 1.55 + index * 0.9,
        w: 2.7,
        h: 0.62,
        rectRadius: 0.06,
        line: { color: COLORS.white, pt: 1, transparency: 35 },
        fill: { color: DECK_STYLE.topBandAccent, transparency: 22 },
      });
      slide.addText(card.label, {
        x: 9.78,
        y: 1.72 + index * 0.9,
        w: 1.55,
        h: 0.16,
        fontSize: 8.8,
        color: COLORS.textSoft,
        fontFace: bodyFont,
      });
      slide.addText(card.value, {
        x: 10.55,
        y: 1.68 + index * 0.9,
        w: 1.48,
        h: 0.18,
        fontSize: 13.5,
        bold: true,
        color: COLORS.white,
        align: 'right',
        fontFace: headingFont,
      });
    });
  }
}

function addGenericOverviewSlide(prs, dataset) {
  const slide = prs.addSlide();
  applyTransition(slide, 'push', 'med');
  const sectionTitle = 'Resumen ejecutivo del archivo';
  const sectionSubtitle = 'La lectura se concentra en indicadores, focos de concentración y señales de negocio con mayor impacto directivo.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.28);
  addHeader(slide, '01', 'PANORAMA', sectionTitle, sectionSubtitle);

  const overviewCards = dataset.headlineCards.length ? dataset.headlineCards.slice(0, 4) : [];
  const overviewLayout = selectUniversalLayoutFamily('overview', dataset.workbookName || dataset.workbookTitle, dataset.layoutProfile);
  const accents = getDeckChartPalette(4);
  const chartLabels = dataset.featuredSheets[0]?.chart?.labels?.slice(0, 8) || dataset.sheets.slice(0, 8).map((sheet) => cleanText(sheet.name, 22));
  const chartValues = dataset.featuredSheets[0]?.chart?.values?.slice(0, 8) || dataset.sheets.slice(0, 8).map((sheet) => sheet.rowCount);
  const leadTableSheet = dataset.featuredSheets.find((sheet) => hasRenderableTableSummary(sheet)) || dataset.sheets.find((sheet) => hasRenderableTableSummary(sheet));

  function renderChartPanel(x, y, w, h, accentIndex = 1) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h,
      rectRadius: 0.08,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: COLORS.card },
    });
    slide.addText('Frentes de mayor impacto', {
      x: x + 0.27,
      y: y + 0.23,
      w: 3.2,
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: COLORS.ink,
      fontFace: getDeckFont('head'),
    });
    slide.addChart(
      prs.ChartType.bar,
      [{ name: dataset.featuredSheets[0]?.chart?.title || 'Valor', labels: chartLabels, values: chartValues }],
      {
        x: x + 0.27,
        y: y + 0.58,
        w: w - 0.52,
        h: Math.max(1.45, h - 0.88),
        barDir: 'bar',
        chartColors: [getDeckAccent(accentIndex)],
        showValue: true,
        showLegend: false,
        catAxisLabelColor: COLORS.muted,
        valAxisLabelColor: COLORS.muted,
        dataLabelColor: COLORS.ink,
        valGridLine: { color: 'DCE9D0' },
      }
    );
  }

  function renderOverviewTablePanel(x, y, w, h) {
    const preferredHeaders = leadTableSheet?.primaryVisual?.type === 'table' && leadTableSheet.primaryVisual.headers?.length
      ? leadTableSheet.primaryVisual.headers
      : (leadTableSheet?.tableHeaders || []);
    const preferredRows = leadTableSheet?.primaryVisual?.type === 'table' && leadTableSheet.primaryVisual.rows?.length
      ? leadTableSheet.primaryVisual.rows.map((row) => row.map((cell) => cleanText(cell?.text) || 'N/D'))
      : (leadTableSheet?.tableRows || []);
    const tableHeadersAll = preferredHeaders.slice(0, 4).map((header) => cleanText(header, 18));
    const tableRowsAll = preferredRows
      .slice(0, 4)
      .map((row) => row.slice(0, tableHeadersAll.length).map((value) => cleanText(value) || 'N/D'))
      .filter((row) => countRenderablePreviewCells(row) >= Math.max(2, Math.min(tableHeadersAll.length, 3)));
    const colWidthsAll = Array(Math.max(1, tableHeadersAll.length)).fill(Math.max(1.45, (w - 0.54) / Math.max(1, tableHeadersAll.length)));
    const fit = findFittingManualRowLimit(tableRowsAll, colWidthsAll, {
      preferredLimit: Math.min(4, tableRowsAll.length),
      minLimit: 2,
      availableHeight: Math.max(1.15, h - 0.95),
      baseFontSize: 8.2,
      minRowHeight: 0.56,
      maxRowHeight: 0.84,
      headerHeight: 0.34,
      maxLinesPerCell: 4.7,
      maxAverageLines: 2.8,
    });
    const rows = tableRowsAll.slice(0, Math.max(2, fit.rowLimit || 2));
    const metrics = evaluateManualTableDensity(rows, colWidthsAll, {
      availableHeight: Math.max(1.15, h - 0.95),
      baseFontSize: 8.2,
      minRowHeight: 0.56,
      maxRowHeight: 0.84,
      headerHeight: 0.34,
      maxLinesPerCell: 4.7,
      maxAverageLines: 2.8,
    });

    slide.addShape(pptx.ShapeType.roundRect, {
      x,
      y,
      w,
      h,
      rectRadius: 0.08,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: COLORS.card },
    });
    slide.addText(leadTableSheet ? `Tabla priorizada: ${leadTableSheet.blockTitle || leadTableSheet.name}` : 'Tabla priorizada', {
      x: x + 0.27,
      y: y + 0.23,
      w: w - 0.54,
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: COLORS.ink,
      fontFace: getDeckFont('head'),
      fit: 'shrink',
    });
    renderManualTable(slide, {
      x: x + 0.27,
      y: y + 0.58,
      colWidths: colWidthsAll,
      headerHeight: 0.34,
      rowHeight: 0.42,
      rowHeights: metrics.rowHeights,
      defaultFontSize: 8.2,
      minFontSize: 7.9,
      headers: tableHeadersAll.map((header) => ({ text: cleanText(header, 18), fontSize: 8.8 })),
      rows,
    });
  }

  if (PRESENTATION_VISUAL_MODE === 'tables' || overviewLayout === 'table-overview') {
    overviewCards.forEach((card, index) => {
      addMetricCard(slide, 0.55 + index * 3.05, topY, 2.85, 1.02, card.label, card.value, accents[index], card.sheet || card.note);
    });
    renderOverviewTablePanel(0.55, Math.max(3.35, topY + 1.3), 7.25, 3.0);
    addInsightBox(slide, 8.05, Math.max(3.35, topY + 1.3), 4.73, 3.0, 'Lectura ejecutiva', dataset.insights.slice(0, 2), getSoftAccentColor(getDeckAccent(0)));
  } else if (overviewLayout === 'chart-hero') {
    overviewCards.slice(0, 3).forEach((card, index) => {
      addMetricCard(slide, 0.55, topY + index * 1.14, 2.6, 1.0, card.label, card.value, accents[index], card.sheet || card.note);
    });
    renderChartPanel(3.45, topY, 9.33, 3.25, 1);
    addInsightBox(slide, 3.45, Math.max(5.36, topY + 3.45), 9.33, 1.0, 'Lectura ejecutiva', dataset.insights.slice(0, 2), getSoftAccentColor(getDeckAccent(0)));
  } else if (overviewLayout === 'editorial-band') {
    overviewCards.forEach((card, index) => {
      const x = index < 2 ? 0.55 + index * 3.05 : 0.55 + (index - 2) * 3.05;
      const y = index < 2 ? topY : topY + 1.18;
      addMetricCard(slide, x, y, 2.8, 1.04, card.label, card.value, accents[index], card.sheet || card.note);
    });
    addInsightBox(slide, 6.72, topY, 6.06, 2.22, 'Lectura ejecutiva', dataset.insights.slice(0, 3), getSoftAccentColor(getDeckAccent(0)));
    renderChartPanel(0.55, Math.max(4.42, topY + 2.58), 12.23, 1.92, 1);
  } else if (overviewLayout === 'summary-grid') {
    overviewCards.forEach((card, index) => {
      const x = 0.55 + (index % 2) * 3.05;
      const y = topY + Math.floor(index / 2) * 1.16;
      addMetricCard(slide, x, y, 2.84, 1.02, card.label, card.value, accents[index], card.sheet || card.note);
    });
    renderChartPanel(6.55, topY, 6.23, 3.0, 1);
    addInsightBox(slide, 0.55, Math.max(4.58, topY + 2.58), 5.85, 1.76, 'Lectura ejecutiva', dataset.insights.slice(0, 2), getSoftAccentColor(getDeckAccent(0)));
  } else {
    overviewCards.forEach((card, index) => {
      addMetricCard(slide, 0.55 + index * 3.05, topY, 2.85, 1.12, card.label, card.value, accents[index], card.sheet || card.note);
    });
    addInsightBox(slide, 0.55, Math.max(3.35, topY + 1.4), 4.7, 3.0, 'Lectura ejecutiva', dataset.insights, getSoftAccentColor(getDeckAccent(0)));
    renderChartPanel(5.55, Math.max(3.35, topY + 1.4), 7.23, 3.0, 1);
  }
}

function addGenericSheetFocusSlide(prs, sheet, pageNumber) {
  const slide = prs.addSlide();
  applyTransition(slide, 'fade', 'med');
  const focusLayout = selectUniversalLayoutFamily('focus', sheet.rawName || sheet.name, sheet.layoutProfile);
  const focusLabel = sheet.blockTitle || sheet.name;
  const focusTitle = `Lectura ejecutiva: ${focusLabel}`;
  const focusSubtitle = sheet.contextSummary || 'Cada hoja se resume con hallazgos ejecutivos, una visualización principal y evidencia tabular compacta.';
  const headerMetrics = getHeaderLayoutMetrics(getDeckVariant('header'), focusTitle, focusSubtitle);
  const topCardsY = Math.max(1.95, headerMetrics.contentBottom + 0.28);
  addHeader(
    slide,
    pageNumber,
    'HOJA CLAVE',
    focusTitle,
    focusSubtitle
  );

  const cardMetrics = (sheet.headlineMetrics.length ? sheet.headlineMetrics.slice(0, 3) : [
    { label: 'Indicador principal', value: sheet.primaryMetric ? humanizeLabel(sheet.primaryMetric.header) : 'N/D', note: 'Hallazgo base' },
    { label: 'Categoria lider', value: sheet.dominantCategoryEntries?.[0] ? humanizeLabel(sheet.dominantCategoryEntries[0][0], 18) : 'N/D', note: 'Mayor peso detectado' },
    { label: 'Hoja', value: focusLabel, note: 'Foco de análisis' },
  ]);

  let visualBox;
  let insightBox;
  let contextBox;

  if (focusLayout === 'editorial-split') {
    cardMetrics.forEach((metric, index) => {
      addMetricCard(slide, 0.55, topCardsY + index * 1.08, 3.2, 0.95, metric.label, metric.value, getDeckAccent(index), metric.note);
    });
    visualBox = { x: 4.02, y: topCardsY, w: 8.76, h: 2.46 };
    insightBox = { x: 4.02, y: Math.max(5.0, topCardsY + 2.7), w: 4.18, h: 1.35, lines: sheet.insights.slice(0, 2) };
    contextBox = { x: 8.5, y: Math.max(5.0, topCardsY + 2.7), w: 4.28, h: 1.35 };
  } else if (focusLayout === 'visual-top') {
    cardMetrics.forEach((metric, index) => {
      addMetricCard(slide, 0.55 + index * 4.0, topCardsY, 3.7, 1.0, metric.label, metric.value, getDeckAccent(index), metric.note);
    });
    visualBox = { x: 0.55, y: Math.max(3.28, topCardsY + 1.28), w: 12.23, h: 1.96 };
    insightBox = { x: 0.55, y: 5.38, w: 5.7, h: 0.97, lines: sheet.insights.slice(0, 2) };
    contextBox = { x: 6.55, y: 5.38, w: 6.23, h: 0.97 };
  } else if (focusLayout === 'visual-right') {
    cardMetrics.forEach((metric, index) => {
      addMetricCard(slide, 0.55 + index * 4.0, topCardsY, 3.7, 1.05, metric.label, metric.value, getDeckAccent(index), metric.note);
    });
    const primaryRowY = Math.max(3.3, topCardsY + 1.35);
    visualBox = { x: 6.62, y: primaryRowY, w: 6.16, h: 3.1 };
    insightBox = { x: 0.55, y: primaryRowY, w: 2.45, h: 3.1, lines: sheet.insights.slice(0, 3) };
    contextBox = { x: 3.2, y: primaryRowY, w: 3.18, h: 3.1 };
  } else if (focusLayout === 'table-led') {
    cardMetrics.forEach((metric, index) => {
      const x = index < 2 ? 0.55 + index * 3.1 : 6.95;
      const y = index < 2 ? topCardsY : topCardsY;
      addMetricCard(slide, x, y, index < 2 ? 2.9 : 2.45, 1.0, metric.label, metric.value, getDeckAccent(index), metric.note);
    });
    const primaryRowY = Math.max(3.28, topCardsY + 1.28);
    visualBox = { x: 0.55, y: primaryRowY, w: 7.3, h: 3.02 };
    insightBox = { x: 8.1, y: primaryRowY, w: 4.68, h: 1.45, lines: sheet.insights.slice(0, 2) };
    contextBox = { x: 8.1, y: primaryRowY + 1.57, w: 4.68, h: 1.45 };
  } else {
    cardMetrics.forEach((metric, index) => {
      addMetricCard(slide, 0.55 + index * 4.0, topCardsY, 3.7, 1.05, metric.label, metric.value, getDeckAccent(index), metric.note);
    });
    const primaryRowY = Math.max(3.3, topCardsY + 1.35);
    visualBox = { x: 0.55, y: primaryRowY, w: 7.15, h: 3.1 };
    insightBox = { x: 7.95, y: primaryRowY, w: 4.82, h: 1.46, lines: sheet.insights.slice(0, 2) };
    contextBox = { x: 7.95, y: primaryRowY + 1.64, w: 4.82, h: 1.46 };
  }

  if (sheet.primaryVisual?.type === 'chart' && sheet.chart) {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: visualBox.x,
      y: visualBox.y,
      w: visualBox.w,
      h: visualBox.h,
      rectRadius: 0.08,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: COLORS.card },
    });
    slide.addText(sheet.chart.title, {
      x: visualBox.x + 0.27,
      y: visualBox.y + 0.24,
      w: Math.max(2.6, visualBox.w - 1.2),
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: COLORS.ink,
      fontFace: getDeckFont('head'),
    });
    slide.addChart(
      prs.ChartType.bar,
      [{ name: 'Valor', labels: sheet.chart.labels, values: sheet.chart.values }],
      {
        x: visualBox.x + 0.27,
        y: visualBox.y + 0.6,
        w: visualBox.w - 0.6,
        h: visualBox.h - 0.95,
        barDir: 'bar',
        chartColors: [getDeckAccent(0)],
        showValue: true,
        showLegend: false,
        catAxisLabelColor: COLORS.muted,
        valAxisLabelColor: COLORS.muted,
        dataLabelColor: COLORS.ink,
        valGridLine: { color: 'DCE9D0' },
      }
    );
  } else if (sheet.primaryVisual?.type === 'table') {
    slide.addShape(pptx.ShapeType.roundRect, {
      x: visualBox.x,
      y: visualBox.y,
      w: visualBox.w,
      h: visualBox.h,
      rectRadius: 0.08,
      line: { color: COLORS.line, pt: 1 },
      fill: { color: COLORS.card },
    });
    slide.addText(sheet.primaryVisual.title || 'Resumen tabular', {
      x: visualBox.x + 0.27,
      y: visualBox.y + 0.24,
      w: Math.max(2.6, visualBox.w - 1.1),
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: COLORS.ink,
      fontFace: getDeckFont('head'),
    });
    const leftHeadersAll = (sheet.primaryVisual.headers || []).slice(0, 4);
    const leftRowsAll = (sheet.primaryVisual.rows || []).slice(0, 5).map((row) => row.slice(0, leftHeadersAll.length));
    const leftTableWidth = Math.max(2.3, visualBox.w - 0.7);
    const leftColWidthsAll = Array(leftHeadersAll.length).fill(leftTableWidth / Math.max(1, leftHeadersAll.length));
    const leftFit = findFittingManualRowLimit(leftRowsAll, leftColWidthsAll, {
      preferredLimit: Math.min(5, leftRowsAll.length),
      minLimit: 2,
      availableHeight: Math.max(1.25, visualBox.h - 0.95),
      baseFontSize: 8.0,
      minRowHeight: 0.54,
      maxRowHeight: 0.84,
      headerHeight: 0.34,
      maxLinesPerCell: 4.8,
      maxAverageLines: 2.85,
    });
    const useCompactPreview = (leftFit.metrics.maxLines >= 5 || leftFit.metrics.avgLines >= 2.95) && leftHeadersAll.length > 3;
    const visualColumnIndexes = useCompactPreview ? [0, 1, 2] : leftHeadersAll.map((_, index) => index);
    const leftHeaders = visualColumnIndexes.map((columnIndex) => leftHeadersAll[columnIndex]);
    const leftRows = projectManualRows(leftRowsAll, visualColumnIndexes).slice(0, Math.max(2, leftFit.rowLimit || 2));
    const leftColWidths = projectColumnWidths(leftColWidthsAll, visualColumnIndexes, leftTableWidth);
    const leftMetrics = evaluateManualTableDensity(leftRows, leftColWidths, {
      availableHeight: Math.max(1.25, visualBox.h - 0.95),
      baseFontSize: 8.0,
      minRowHeight: 0.54,
      maxRowHeight: 0.84,
      headerHeight: 0.34,
      maxLinesPerCell: 4.8,
      maxAverageLines: 2.85,
    });
    renderManualTable(slide, {
      x: visualBox.x + 0.27,
      y: visualBox.y + 0.6,
      colWidths: leftColWidths,
      headerHeight: 0.34,
      rowHeight: 0.42,
      rowHeights: leftMetrics.rowHeights,
      defaultFontSize: 8.0,
      minFontSize: 7.8,
      headers: leftHeaders.map((header) => ({ text: cleanText(header, 18), fontSize: 8.8 })),
      rows: leftRows,
    });
  }

  addInsightBox(
    slide,
    insightBox.x,
    insightBox.y,
    insightBox.w,
    insightBox.h,
    'Hallazgos',
    insightBox.lines,
    COLORS.aquaSoft
  );

  slide.addShape(pptx.ShapeType.roundRect, {
    x: contextBox.x,
    y: contextBox.y,
    w: contextBox.w,
    h: contextBox.h,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  if (sheet.primaryVisual?.type === 'table' || PRESENTATION_VISUAL_MODE === 'charts') {
    slide.addText('Contexto detectado', {
      x: contextBox.x + 0.2,
      y: contextBox.y + 0.24,
      w: contextBox.w - 0.4,
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: COLORS.ink,
      fontFace: getDeckFont('head'),
    });
    const contextLines = [
      sheet.blockTitle ? `Titulo: ${sheet.blockTitle}` : null,
      sheet.blockSubtitle ? `Subtitulo: ${sheet.blockSubtitle}` : null,
      sheet.blockNotes?.[0] ? `Nota: ${cleanText(sheet.blockNotes[0], 80)}` : null,
      PRESENTATION_VISUAL_MODE === 'charts' && sheet.chart?.labels?.[0]
        ? `Grafica priorizada: ${sheet.chart.labels[0]} lidera ${cleanText(sheet.chart.title, 44).toLowerCase()}.`
        : null,
      sheet.blockCount > 1 ? `Bloque ${sheet.blockIndex + 1} de ${sheet.blockCount} en ${sheet.sourceSheetName}.` : `Hoja origen: ${sheet.sourceSheetName}.`,
    ].filter(Boolean);
    contextLines.slice(0, focusLayout === 'visual-top' ? 2 : focusLayout === 'visual-right' ? 4 : 3).forEach((line, index) => {
      slide.addText(`• ${line}`, {
        x: contextBox.x + 0.22,
        y: contextBox.y + 0.63 + index * (focusLayout === 'visual-top' ? 0.22 : focusLayout === 'visual-right' ? 0.48 : 0.32),
        w: contextBox.w - 0.38,
        h: 0.34,
        fontSize: focusLayout === 'visual-top' ? 8.1 : focusLayout === 'visual-right' ? 8.6 : 8.4,
        color: COLORS.ink,
        fontFace: getDeckFont('body'),
        fit: 'shrink',
      });
    });
  } else {
    const previewHeadersAll = sheet.tableHeaders.slice(0, Math.min(4, sheet.tableHeaders.length));
    slide.addText('Muestra de datos', {
      x: contextBox.x + 0.2,
      y: contextBox.y + 0.24,
      w: contextBox.w - 0.4,
      h: 0.2,
      fontSize: 12,
      bold: true,
      color: COLORS.ink,
      fontFace: getDeckFont('head'),
    });
    const previewRowLimit = focusLayout === 'visual-top' ? 2 : focusLayout === 'visual-right' ? 5 : 3;
    const manualRowsAll = sheet.tableRows.slice(0, previewRowLimit).map((row, rowIndex) =>
      row.slice(0, previewHeadersAll.length).map((value) => ({
        text: cleanText(value),
        fill: rowIndex % 2 === 0 ? 'F8FBF4' : COLORS.white,
        fontSize: focusLayout === 'visual-top' ? 7.9 : 8.2,
      }))
    );
    const previewWidth = Math.max(1.9, contextBox.w - 0.38);
    const previewColWidthsAll = Array(previewHeadersAll.length).fill(previewWidth / Math.max(1, previewHeadersAll.length));
    const previewFit = findFittingManualRowLimit(manualRowsAll, previewColWidthsAll, {
      preferredLimit: Math.min(previewRowLimit, manualRowsAll.length),
      minLimit: 2,
      availableHeight: Math.max(1.2, contextBox.h - 0.95),
      baseFontSize: focusLayout === 'visual-top' ? 7.9 : 8.2,
      minRowHeight: 0.5,
      maxRowHeight: focusLayout === 'visual-top' ? 0.64 : 0.78,
      headerHeight: 0.34,
      maxLinesPerCell: 4.8,
      maxAverageLines: 2.9,
    });
    const compactPreview = previewFit.metrics.maxLines >= 5 || previewFit.metrics.avgLines >= 2.95;
    const previewIndexes = compactPreview && previewHeadersAll.length > 3
      ? [0, 1, 2]
      : previewHeadersAll.map((_, index) => index);
    const trimmedHeaders = previewIndexes.map((columnIndex) => previewHeadersAll[columnIndex]);
    const trimmedRows = projectManualRows(manualRowsAll, previewIndexes).slice(0, Math.max(2, previewFit.rowLimit || 2));
    const previewColWidths = projectColumnWidths(previewColWidthsAll, previewIndexes, previewWidth);
    const previewMetrics = evaluateManualTableDensity(trimmedRows, previewColWidths, {
      availableHeight: Math.max(1.2, contextBox.h - 0.95),
      baseFontSize: focusLayout === 'visual-top' ? 7.9 : 8.2,
      minRowHeight: 0.5,
      maxRowHeight: focusLayout === 'visual-top' ? 0.64 : 0.78,
      headerHeight: 0.34,
      maxLinesPerCell: 4.8,
      maxAverageLines: 2.9,
    });

    renderManualTable(slide, {
      x: contextBox.x + 0.18,
      y: contextBox.y + 0.6,
      colWidths: previewColWidths,
      headerHeight: 0.34,
      rowHeight: 0.42,
      rowHeights: previewMetrics.rowHeights,
      defaultFontSize: focusLayout === 'visual-top' ? 7.9 : 8.2,
      minFontSize: 7.8,
      headers: trimmedHeaders.map((header) => ({ text: cleanText(header, 18), fontSize: focusLayout === 'visual-top' ? 8.4 : 8.8 })),
      rows: trimmedRows,
    });
  }
}

function addGenericWorkbookMapSlide(prs, dataset) {
  const priorityRows = dataset.priorityRows || [];
  const sectionTitle = 'Prioridades detectadas en el libro';
  const sectionSubtitle = 'Cada frente resume el foco gerencial, su indicador rector y la senal mas accionable encontrada en la lectura.';
  const mapLayout = selectUniversalLayoutFamily('map', dataset.workbookName || dataset.workbookTitle, dataset.layoutProfile);
  const forceTableSummary = priorityRows.some((row) =>
    cleanText(row.signal).length >= 110 || cleanText(row.focus).length >= 75 || cleanText(row.metric).length >= 58
  );

  if (mapLayout === 'editorial-strip' && !forceTableSummary) {
    const chunks = chunkArray(priorityRows.slice(0, 6), 2);
    chunks.forEach((chunk, chunkIndex) => {
      const slide = prs.addSlide();
      applyTransition(slide, 'push', 'fast');
      const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.28);
      addHeader(
        slide,
        formatSectionPageNumber('05', chunkIndex),
        'MAPA EJECUTIVO',
        sectionTitle,
        sectionSubtitle
      );

      chunk.forEach((row, index) => {
        const y = topY + index * 2.32;
        slide.addShape(pptx.ShapeType.roundRect, {
          x: 0.55,
          y,
          w: 12.23,
          h: 1.95,
          rectRadius: 0.08,
          line: { color: COLORS.line, pt: 1 },
          fill: { color: COLORS.card },
        });
        slide.addShape(pptx.ShapeType.rect, {
          x: 0.55,
          y,
          w: 0.12,
          h: 1.95,
          line: { color: getDeckAccent(index), transparency: 100 },
          fill: { color: getDeckAccent(index) },
        });
        slide.addText(cleanText(row.sheet), {
          x: 0.82,
          y: y + 0.16,
          w: 3.1,
          h: 0.24,
          fontSize: 11.2,
          bold: true,
          color: COLORS.ink,
          fontFace: getDeckFont('head'),
          fit: 'shrink',
        });
        slide.addText(cleanText(row.focus), {
          x: 0.82,
          y: y + 0.5,
          w: 3.2,
          h: 0.18,
          fontSize: 9.0,
          color: COLORS.muted,
          fontFace: getDeckFont('body'),
          italic: true,
          fit: 'shrink',
        });
        slide.addText(cleanText(row.metric), {
          x: 4.28,
          y: y + 0.2,
          w: 2.15,
          h: 0.22,
          fontSize: 10.1,
          bold: true,
          color: getDeckAccent(index + 1),
          fontFace: getDeckFont('head'),
          fit: 'shrink',
        });
        slide.addText(cleanText(row.signal), {
          x: 4.28,
          y: y + 0.56,
          w: 7.95,
          h: 0.72,
          fontSize: 8.8,
          color: COLORS.ink,
          fontFace: getDeckFont('body'),
          fit: 'shrink',
        });
      });
    });
  } else if (mapLayout === 'cards-grid' && !forceTableSummary) {
    const chunks = chunkArray(priorityRows.slice(0, 6), 3);
    chunks.forEach((chunk, chunkIndex) => {
      const slide = prs.addSlide();
      applyTransition(slide, 'push', 'fast');
      const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.28);
      addHeader(
        slide,
        formatSectionPageNumber('05', chunkIndex),
        'MAPA EJECUTIVO',
        sectionTitle,
        sectionSubtitle
      );

      chunk.forEach((row, index) => {
        const x = 0.55 + (index % 2) * 6.12;
        const y = topY + Math.floor(index / 2) * 2.36;
        slide.addShape(pptx.ShapeType.roundRect, {
          x,
          y,
          w: 5.92,
          h: 2.08,
          rectRadius: 0.08,
          line: { color: COLORS.line, pt: 1 },
          fill: { color: COLORS.card },
        });
        slide.addShape(pptx.ShapeType.rect, {
          x,
          y,
          w: 0.1,
          h: 2.08,
          line: { color: getDeckAccent(index), transparency: 100 },
          fill: { color: getDeckAccent(index) },
        });
        slide.addText(cleanText(row.sheet), {
          x: x + 0.22,
          y: y + 0.16,
          w: 2.65,
          h: 0.26,
          fontSize: 11.2,
          bold: true,
          color: COLORS.ink,
          fontFace: getDeckFont('head'),
          fit: 'shrink',
        });
        slide.addText(cleanText(row.focus), {
          x: x + 0.22,
          y: y + 0.52,
          w: 3.25,
          h: 0.18,
          fontSize: 9.2,
          color: COLORS.muted,
          fontFace: getDeckFont('body'),
          italic: true,
          fit: 'shrink',
        });
        slide.addText(cleanText(row.metric), {
          x: x + 0.22,
          y: y + 0.9,
          w: 3.2,
          h: 0.22,
          fontSize: 10.2,
          bold: true,
          color: getDeckAccent(index),
          fontFace: getDeckFont('head'),
          fit: 'shrink',
        });
        slide.addText(cleanText(row.signal), {
          x: x + 0.22,
          y: y + 1.24,
          w: 5.3,
          h: 0.5,
          fontSize: 8.8,
          color: COLORS.ink,
          fontFace: getDeckFont('body'),
          fit: 'shrink',
        });
      });
    });
  } else {
    const columns = [
      { header: 'Frente', width: 2.2, buildCell: (row) => ({ text: cleanText(row.sheet), fontSize: 8.6 }) },
      { header: 'Foco directivo', width: 2.35, buildCell: (row) => ({ text: cleanText(row.focus), fontSize: 8.6 }) },
      { header: 'Indicador rector', width: 2.25, buildCell: (row) => ({ text: cleanText(row.metric), fontSize: 8.4 }) },
      { header: 'Senal principal', width: 4.93, buildCell: (row) => ({ text: cleanText(row.signal), fontSize: 8.4 }) },
    ];
    const fullWidths = columns.map((column) => column.width);
    const tableWidth = sumNumericValues(fullWidths);
    const fullRows = priorityRows.map((row, rowIndex) =>
      columns.map((column) => ({
        fill: rowIndex % 2 === 0 ? 'F8FBF4' : COLORS.white,
        color: COLORS.ink,
        ...column.buildCell(row, rowIndex),
      }))
    );
    const fullFit = findFittingManualRowLimit(fullRows, fullWidths, {
      preferredLimit: 4,
      minLimit: 2,
      availableHeight: 4.18,
      baseFontSize: 8.5,
      minRowHeight: 0.82,
      maxRowHeight: 1.22,
      headerHeight: 0.42,
      maxLinesPerCell: 5.0,
      maxAverageLines: 3.0,
    });
    const summaryIndexes = [0, 1, 2];
    const summaryRows = projectManualRows(fullRows, summaryIndexes);
    const summaryWidths = projectColumnWidths(fullWidths, summaryIndexes, tableWidth);
    const summaryFit = findFittingManualRowLimit(summaryRows, summaryWidths, {
      preferredLimit: 5,
      minLimit: 3,
      availableHeight: 4.18,
      baseFontSize: 8.6,
      minRowHeight: 0.76,
      maxRowHeight: 1.05,
      headerHeight: 0.42,
      maxLinesPerCell: 4.6,
      maxAverageLines: 2.75,
    });
    const useSummaryView = forceTableSummary || mapLayout === 'summary-table' || fullFit.rowLimit < Math.min(4, fullRows.length) || fullFit.metrics.maxLines >= 5 || fullFit.metrics.avgLines >= 3;
    const activeRows = useSummaryView ? summaryRows : fullRows;
    const activeWidths = useSummaryView ? summaryWidths : fullWidths;
    const activeHeaders = (useSummaryView ? summaryIndexes : columns.map((_, index) => index)).map((columnIndex) => columns[columnIndex].header);
    const chunkSize = useSummaryView ? Math.max(3, summaryFit.rowLimit || 3) : Math.max(2, fullFit.rowLimit || 2);
    const chunks = chunkArray(activeRows, chunkSize);

    chunks.forEach((chunk, chunkIndex) => {
      const slide = prs.addSlide();
      applyTransition(slide, 'push', 'fast');
      const dynamicTitle = useSummaryView ? `${sectionTitle} | resumen` : sectionTitle;
      const dynamicSubtitle = useSummaryView
        ? 'Cuando el mapa supera la densidad segura, se resume la matriz y el detalle permanece en las slides focales por hoja.'
        : sectionSubtitle;
      const topY = getSafeContentStartY(dynamicTitle, dynamicSubtitle, 1.95, 0.28);
      addHeader(
        slide,
        formatSectionPageNumber('05', chunkIndex),
        'MAPA EJECUTIVO',
        dynamicTitle,
        dynamicSubtitle
      );
      const metrics = evaluateManualTableDensity(chunk, activeWidths, {
        availableHeight: 4.18,
        baseFontSize: useSummaryView ? 8.6 : 8.5,
        minRowHeight: useSummaryView ? 0.76 : 0.82,
        maxRowHeight: useSummaryView ? 1.05 : 1.22,
        headerHeight: 0.42,
        maxLinesPerCell: useSummaryView ? 4.6 : 5.0,
        maxAverageLines: useSummaryView ? 2.75 : 3.0,
      });

      slide.addShape(pptx.ShapeType.roundRect, {
        x: 0.55,
        y: topY,
        w: 12.23,
        h: 4.95,
        rectRadius: 0.08,
        line: { color: COLORS.line, pt: 1 },
        fill: { color: COLORS.card },
      });
      slide.addText(useSummaryView ? 'Resumen automatico' : 'Detalle visible', {
        x: 9.1,
        y: topY + 0.09,
        w: 3.15,
        h: 0.18,
        fontSize: 8.4,
        bold: true,
        color: COLORS.muted,
        align: 'right',
        fontFace: getDeckFont('body'),
      });
      renderManualTable(slide, {
        x: 0.8,
        y: topY + 0.29,
        colWidths: activeWidths,
        headerHeight: 0.42,
        rowHeight: 0.7,
        rowHeights: metrics.rowHeights,
        defaultFontSize: useSummaryView ? 8.6 : 8.5,
        minFontSize: 7.9,
        headers: activeHeaders.map((header) => ({ text: header, fontSize: 9.2 })),
        rows: chunk,
      });
    });
  }
}

function addGenericClosing(prs, dataset) {
  const slide = prs.addSlide();
  slide.background = { color: COLORS.navy };
  applyTransition(slide, 'fade', 'slow');
  addNatureDecor(slide, 'closing');
  const closingVariant = getDeckVariant('genericClosing');
  const headingFont = getDeckFont('head');
  const bodyFont = getDeckFont('body');

  slide.addText('Cierre Ejecutivo', {
    x: closingVariant === 1 ? 0.95 : 1.0,
    y: closingVariant === 1 ? 1.0 : 1.3,
    w: closingVariant === 1 ? 7.2 : 5.5,
    h: 0.5,
    fontSize: 30,
    bold: true,
    color: COLORS.white,
    fontFace: headingFont,
  });
  slide.addText('El libro ya se traduce a una narrativa visual consistente y orientada a decisiones, manteniendo el estilo premium sin depender de un unico tipo de Excel.', {
    x: closingVariant === 1 ? 0.95 : 1.0,
    y: closingVariant === 1 ? 1.72 : 2.1,
    w: closingVariant === 2 ? 5.55 : 6.8,
    h: closingVariant === 1 ? 0.72 : 0.95,
    fontSize: 14,
    color: COLORS.textSoft,
    italic: true,
    fontFace: bodyFont,
    fit: 'shrink',
  });
  const closingCards = dataset.headlineCards.length
    ? dataset.headlineCards.slice(0, 3)
    : [
        { label: 'Hoja foco', value: dataset.featuredSheets[0]?.name || 'N/D', note: 'Frente principal' },
        { label: 'Indicador', value: dataset.featuredSheets[0]?.headlineMetrics?.[0]?.value || 'N/D', note: dataset.featuredSheets[0]?.headlineMetrics?.[0]?.label || 'Lectura clave' },
        { label: 'Categoria', value: dataset.featuredSheets[0]?.headlineMetrics?.[2]?.value || 'N/D', note: 'Mayor peso visible' },
      ];

  if (closingVariant === 1) {
    closingCards.forEach((card, index) => {
      addMetricCard(slide, 0.95 + index * 2.85, 3.1, 2.55, 1.02, card.label, card.value, getDeckAccent(index), card.sheet || card.note);
    });
    addInsightBox(slide, 9.15, 1.32, 3.45, 4.7, 'Mensajes finales', dataset.insights.slice(0, 3), getSoftAccentColor(getDeckAccent(1)));
  } else if (closingVariant === 2) {
    addInsightBox(slide, 8.02, 1.35, 4.18, 3.05, 'Mensajes finales', dataset.insights.slice(0, 3), getSoftAccentColor(getDeckAccent(2)));
    closingCards.forEach((card, index) => {
      addMetricCard(slide, 1.0 + index * 2.4, 4.15, 2.15, 1.0, card.label, card.value, getDeckAccent(index), card.sheet || card.note);
    });
  } else {
    addInsightBox(slide, 8.0, 1.35, 4.1, 3.0, 'Mensajes finales', dataset.insights.slice(0, 3), COLORS.greenSoft);
    closingCards.forEach((card, index) => {
      addMetricCard(slide, 1.0 + index * 2.4, 4.05, 2.15, 1.0, card.label, card.value, getDeckAccent(index), card.sheet || card.note);
    });
  }

  slide.addText(`Archivo procesado: ${SOURCE_FILE_LABEL}`, {
    x: closingVariant === 1 ? 0.95 : 1.0,
    y: 5.22,
    w: 4.6,
    h: 0.2,
    fontSize: 10.5,
    color: COLORS.white,
    fontFace: bodyFont,
  });
  slide.addText('Fundación Socya · Motor universal Excel a PowerPoint', {
    x: closingVariant === 1 ? 0.95 : 1.0,
    y: 6.65,
    w: 7.2,
    h: 0.2,
    fontSize: 9.5,
    color: COLORS.textSoft,
    fontFace: bodyFont,
  });
}

function buildExecutiveRecommendations(dataset) {
  const recommendations = [];
  const sheets = Array.isArray(dataset?.sheets) ? dataset.sheets : [];
  const focusSheet = dataset?.headlineCards?.find((card) => cleanText(card.label).toLowerCase().includes('hoja foco'))?.value || '';

  if (focusSheet) {
    recommendations.push(`Alinear el comité directivo alrededor del frente ${cleanText(focusSheet, 46)} como hoja foco del ciclo de seguimiento.`);
  }

  const operationSheets = sheets.filter((sheet) => sheet.domain === 'operations');
  if (operationSheets.length) {
    recommendations.push('Definir responsables y fechas compromiso para los puntos de mayor concentración operativa detectados en las hojas clave.');
  }

  const financeSheets = sheets.filter((sheet) => sheet.domain === 'finance');
  if (financeSheets.length) {
    recommendations.push('Revisar umbrales y controles preventivos sobre los rubros con mayor peso financiero para reducir exposición y reprocesos.');
  }

  const hasEvidence = sheets.some((sheet) => /evidenc|prueba|document|soporte/.test(normalizeHeaderKey(sheet.blockTitle || sheet.name)));
  if (hasEvidence) {
    recommendations.push('Estandarizar la evidencia mínima esperada por actividad y asegurar trazabilidad por proceso / subproceso.');
  }

  if (!recommendations.length) {
    recommendations.push('Priorizar dos frentes de control y seguimiento con base en las tablas más completas del archivo y formalizar plan de acción.');
  }

  return recommendations.slice(0, 4);
}

function addGenericRecommendationsSlide(prs, dataset, sectionNumber = '06') {
  const slide = prs.addSlide();
  applyTransition(slide, 'fade', 'med');
  const sectionTitle = 'Conclusiones y recomendaciones';
  const sectionSubtitle = 'Síntesis ejecutiva orientada a decisiones, derivada de las hojas priorizadas y de la estructura real del Excel.';
  const topY = getSafeContentStartY(sectionTitle, sectionSubtitle, 1.95, 0.28);
  addHeader(slide, sectionNumber, 'CIERRE EJECUTIVO', sectionTitle, sectionSubtitle);

  const recommendations = buildExecutiveRecommendations(dataset);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.55,
    y: topY,
    w: 7.45,
    h: 4.95,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Recomendaciones accionables', {
    x: 0.82,
    y: topY + 0.22,
    w: 6.6,
    h: 0.2,
    fontSize: 12.5,
    bold: true,
    color: COLORS.ink,
    fontFace: getDeckFont('head'),
  });
  recommendations.forEach((rec, index) => {
    slide.addText(`${index + 1}. ${rec}`, {
      x: 0.82,
      y: topY + 0.58 + index * 0.5,
      w: 7.0,
      h: 0.38,
      fontSize: 10.6,
      color: COLORS.ink,
      fontFace: getDeckFont('body'),
      fit: 'shrink',
    });
  });

  const highlights = Array.isArray(dataset?.insights) ? dataset.insights.slice(0, 3) : [];
  addInsightBox(slide, 8.25, topY, 5.0, 2.35, 'Mensajes clave', highlights, getSoftAccentColor(getDeckAccent(1)));

  const priorityRows = (dataset?.priorityRows || []).slice(0, 4);
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.25,
    y: topY + 2.52,
    w: 5.0,
    h: 2.43,
    rectRadius: 0.08,
    line: { color: COLORS.line, pt: 1 },
    fill: { color: COLORS.card },
  });
  slide.addText('Frentes priorizados', {
    x: 8.48,
    y: topY + 2.72,
    w: 4.55,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    fontFace: getDeckFont('head'),
  });
  priorityRows.forEach((row, index) => {
    slide.addText(`• ${cleanText(row.sheet, 32)}: ${cleanText(row.metric, 46)}`, {
      x: 8.5,
      y: topY + 3.08 + index * 0.42,
      w: 4.65,
      h: 0.28,
      fontSize: 9.2,
      color: COLORS.muted,
      fontFace: getDeckFont('body'),
      fit: 'shrink',
    });
  });
}

async function buildPresentation() {
  if (!fs.existsSync(INPUT_FILE)) {
    throw new Error(`No existe el archivo de entrada: ${INPUT_FILE}`);
  }

  assertNoUnauthorizedDataDependencies(PROCESSING_CONTEXT, [
    INPUT_FILE,
    path.join(__dirname, 'organizer.py'),
    path.join(__dirname, 'generate_excel_org_presentation.js'),
  ]);

  const workbook = XLSX.readFile(INPUT_FILE, { cellDates: true });
  let dataset;
  let presentationMode = 'generic';
  let sourceData = null;
  let safeMode = { enabled: false, reasons: [], organizerTableCount: 0 };
  const auditWarnings = [];

  try {
    sourceData = loadOrganizerData();
  } catch (error) {
    sourceData = null;
    auditWarnings.push(`No se pudo cargar organizer.py: ${error.message}`);
  }

  if (sourceData) {
    validateOrganizerConsistency(sourceData, INPUT_FILE);
  }

  const specializedCommissions = !!(sourceData && isSpecializedCommissionsDataset(sourceData));
  safeMode = decideSafeMode({
    requestedVisualMode: PRESENTATION_VISUAL_MODE,
    sourceData,
    specializedCommissions,
  });

  if (safeMode.enabled) {
    process.env.SOCYA_SAFE_MODE = '1';
  } else {
    delete process.env.SOCYA_SAFE_MODE;
  }

  configureDeckStyle(resolveDeckSeed());

  if (sourceData && PRESENTATION_VISUAL_MODE === 'mixed' && specializedCommissions) {
    dataset = buildDataset(sourceData);
    presentationMode = 'commissions';
  }

  if (!dataset && sourceData) {
    try {
      dataset = buildGenericDatasetFromOrganizer(sourceData, workbook);
      presentationMode = 'generic';
    } catch (error) {
      dataset = null;
      auditWarnings.push(`No se pudo construir el dataset organizado: ${error.message}`);
    }
  }

  if (!dataset) {
    if (safeMode.enabled) {
      throw new Error(
        `Modo seguro activado: el sistema no obtuvo una estructura confiable del Excel ${SOURCE_FILE_LABEL}. ` +
        `Se bloquea cualquier fallback no autorizado para evitar cruces de información.`
      );
    }
    dataset = buildGenericDataset(workbook);
  }

  const prs = new pptxgen();
  prs.layout = 'LAYOUT_WIDE';
  prs.author = 'OpenAI';
  prs.company = 'Fundación Socya';
  prs.subject = presentationMode === 'commissions' ? 'Auditoría de comisiones y control interno' : 'Presentación ejecutiva universal de datos';
  prs.title = presentationMode === 'commissions' ? 'Presentación de auditoría de comisiones' : `Presentación ejecutiva · ${dataset.workbookTitle || SOURCE_FILE_LABEL}`;
  prs.lang = 'es-CO';
  prs.theme = {
    headFontFace: getDeckFont('head'),
    bodyFontFace: getDeckFont('body'),
    lang: 'es-CO',
  };

  if (presentationMode === 'commissions') {
    let commissionSectionCounter = 7;
    addCover(prs, dataset);
    addExecutiveSummary(prs, dataset);
    addFinancialSlide(prs, dataset);
    addOperationalSlide(prs, dataset);
    addTopRequestersSlide(prs, dataset);
    addSampleTableSlide(prs, dataset);
    if (hasMeaningfulCosoContent(dataset)) {
      addCosoSlide(prs, dataset);
      commissionSectionCounter += 1;
    }
    if (hasMeaningfulFindings(dataset)) {
      addFindingsSlide(prs, dataset);
      addHallazgosTableSlide(prs, dataset);
      commissionSectionCounter += 2;
    }
    if (hasMeaningfulOpportunities(dataset)) {
      addOpportunityDashboardSlide(prs, dataset);
      addOpportunityTableSlide(prs, dataset);
      commissionSectionCounter += 2;
    }
    addCommissionRecommendationsSlide(prs, dataset, String(commissionSectionCounter).padStart(2, '0'));
    addClosing(prs, dataset);
  } else {
    addGenericCover(prs, dataset);
    addGenericOverviewSlide(prs, dataset);
    dataset.featuredSheets.forEach((sheet, index) => {
      addGenericSheetFocusSlide(prs, sheet, `0${index + 2}`);
    });
    addGenericWorkbookMapSlide(prs, dataset);
    addGenericRecommendationsSlide(prs, dataset, String(dataset.featuredSheets.length + 4).padStart(2, '0'));
    addGenericClosing(prs, dataset);
  }

  let finalOutputFile = OUTPUT_FILE;
  try {
    await prs.writeFile({ fileName: finalOutputFile });
  } catch (error) {
    if (error && error.code === 'EBUSY') {
      const parsed = path.parse(OUTPUT_FILE);
      finalOutputFile = path.join(parsed.dir, `${parsed.name}_verde${parsed.ext}`);
      await prs.writeFile({ fileName: finalOutputFile });
    } else {
      throw error;
    }
  }

  const auditRecord = buildAuditRecord({
    context: PROCESSING_CONTEXT,
    presentationMode,
    safeMode,
    sourceData,
    outputFile: finalOutputFile,
    warnings: auditWarnings,
  });
  writeAuditRecord(PROCESSING_CONTEXT.auditLogFile, auditRecord);

  console.log(`Presentación creada en: ${finalOutputFile}`);
  console.log(`Auditoría creada en: ${PROCESSING_CONTEXT.auditLogFile}`);
}

buildPresentation().catch((error) => {
  console.error(error);
  process.exit(1);
});
