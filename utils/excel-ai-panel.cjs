const STANDARD_TIMEOUT_MS = 6 * 60 * 1000;
const EXTENDED_TIMEOUT_MS = 10 * 60 * 1000;
const MAX_TIMEOUT_MS = 15 * 60 * 1000;
const SHALLOW_VALIDATION_THRESHOLD_BYTES = 12 * 1024 * 1024;
const EXTENDED_PROCESSING_THRESHOLD_BYTES = 18 * 1024 * 1024;

function compactNumber(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return null;
  }

  if (Math.abs(value) >= 1_000_000) {
    return `${(value / 1_000_000).toFixed(1)}M`;
  }

  if (Math.abs(value) >= 1_000) {
    return `${(value / 1_000).toFixed(1)}K`;
  }

  return String(Math.round(value));
}

function uniqueStrings(values, limit = Infinity) {
  const seen = new Set();
  const result = [];

  for (const value of values || []) {
    const text = String(value || '').replace(/\s+/g, ' ').trim();
    if (!text) continue;
    const key = text.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(text);
    if (result.length >= limit) break;
  }

  return result;
}

function collectTableCount(analysis) {
  return [
    analysis?.muestra_tabla ? 1 : 0,
    Object.keys(analysis?.otras_tablas || {}).length,
    Object.keys(analysis?.genericas || {}).length,
    analysis?.coso ? 1 : 0,
    analysis?.distribucion_mes ? 1 : 0,
  ].reduce((sum, count) => sum + count, 0);
}

function buildProcessingProfile({ fileSizeBytes = 0, userPrompt = '' } = {}) {
  const promptLength = String(userPrompt || '').trim().length;
  let timeoutMs = STANDARD_TIMEOUT_MS;
  let tier = 'estandar';

  if (fileSizeBytes >= SHALLOW_VALIDATION_THRESHOLD_BYTES || promptLength >= 240) {
    timeoutMs = EXTENDED_TIMEOUT_MS;
    tier = 'extendido';
  }

  if (fileSizeBytes >= EXTENDED_PROCESSING_THRESHOLD_BYTES || promptLength >= 700) {
    timeoutMs = MAX_TIMEOUT_MS;
    tier = 'alto-volumen';
  }

  return {
    tier,
    timeoutMs,
    timeoutMinutes: Math.round(timeoutMs / 60000),
    skipClientDeepValidation: fileSizeBytes >= SHALLOW_VALIDATION_THRESHOLD_BYTES,
    allowsExtendedProcessing: timeoutMs > STANDARD_TIMEOUT_MS,
  };
}

function buildRecommendedSlides(analysis, dataset) {
  const slides = [];
  const chartCount = Array.isArray(analysis?.graficas_automaticas) ? analysis.graficas_automaticas.length : 0;
  const kpiCount = Array.isArray(analysis?.kpis_automaticos) ? analysis.kpis_automaticos.length : 0;
  const textBlockCount = Array.isArray(analysis?.bloques_textuales) ? analysis.bloques_textuales.length : 0;
  const tableCount = collectTableCount(analysis);

  slides.push({
    type: 'title',
    title: 'Portada ejecutiva',
    reason: 'Abre con el objetivo del archivo y la lectura gerencial del conjunto de datos.',
  });

  if (kpiCount > 0) {
    slides.push({
      type: 'kpi',
      title: 'KPIs prioritarios',
      reason: `El análisis detectó ${kpiCount} KPI(s) con suficiente señal para un resumen ejecutivo.`,
    });
  }

  if (chartCount > 0) {
    slides.push({
      type: 'chart',
      title: 'Tendencias y concentraciones',
      reason: `Hay ${chartCount} gráfica(s) candidata(s) para explicar patrones y comparativas.`,
    });
  }

  if (tableCount > 0) {
    slides.push({
      type: 'table',
      title: 'Tabla de soporte',
      reason: `Se detectaron ${tableCount} bloque(s) tabulares para respaldo y detalle controlado.`,
    });
  }

  if (textBlockCount > 0 || dataset.workbookType === 'auditoria_control') {
    slides.push({
      type: 'narrative',
      title: 'Hallazgos y recomendaciones',
      reason: 'Conviene cerrar con insights, riesgos y acciones sugeridas por la IA.',
    });
  }

  return slides.slice(0, 5);
}

function buildPatterns(analysis) {
  const pareto = Array.isArray(analysis?.analisis_avanzado?.pareto)
    ? analysis.analisis_avanzado.pareto
    : [];

  const patterns = [];

  for (const item of pareto.slice(0, 3)) {
    const column = String(item?.columna || '').trim();
    const topValue = item?.top_valores?.[0];
    const label = String(topValue?.valor || '').trim();
    const pct = typeof topValue?.pct === 'number' ? `${Math.round(topValue.pct)}%` : null;

    if (column && label) {
      patterns.push(
        pct
          ? `${column}: "${label}" concentra aproximadamente ${pct} de la distribución observada.`
          : `${column}: "${label}" aparece como el patrón dominante en la distribución.`
      );
    }
  }

  return uniqueStrings(patterns, 3);
}

function buildTrends(analysis) {
  const trends = [];
  const trend = analysis?.analisis_avanzado?.tendencia || {};
  const temporalColumn = String(trend?.col_temporal || '').trim();
  const valueColumn = String(trend?.col_valor || '').trim();
  const trendType = String(trend?.tendencia || '').trim();

  if (temporalColumn && valueColumn && trendType) {
    trends.push(`La serie temporal entre ${temporalColumn} y ${valueColumn} muestra una tendencia ${trendType}.`);
  }

  const rawInsights = Array.isArray(analysis?.analisis_avanzado?.insights)
    ? analysis.analisis_avanzado.insights
    : [];

  for (const insight of rawInsights.slice(0, 2)) {
    const text = typeof insight === 'string' ? insight : insight?.texto;
    if (text) trends.push(text);
  }

  return uniqueStrings(trends, 4);
}

function buildKeyFindings(analysis) {
  const findings = [];
  const executive = analysis?.resumen_ejecutivo_ia || {};
  const briefing = analysis?.briefing_ejecutivo_ia || {};

  findings.push(executive?.vision_general);
  findings.push(executive?.alerta_principal);
  findings.push(executive?.recomendacion);

  for (const key of [
    'breve_resumen',
    'puntos_a_tratar',
    'elementos_prioritarios',
    'planeamiento',
  ]) {
    for (const item of briefing?.[key] || []) {
      findings.push(item);
    }
  }

  for (const item of analysis?.conclusiones || []) {
    findings.push(item);
  }

  return uniqueStrings(findings, 8);
}

function inferPalette(dataset) {
  if (dataset.workbookType === 'auditoria_control') {
    return {
      name: 'Comité Ejecutivo',
      colors: ['#0F172A', '#1D4ED8', '#F59E0B', '#E2E8F0'],
    };
  }

  return {
    name: 'Analítica Moderna',
    colors: ['#111827', '#2563EB', '#10B981', '#F3F4F6'],
  };
}

function buildPromptHints(report) {
  const hints = [];
  const primarySheet = report?.dataset?.primarySheet;
  const chartTitles = report?.powerPointPlan?.preferredCharts || [];
  const findings = report?.keyFindings || [];

  if (primarySheet) {
    hints.push(`Estructura la historia alrededor de la hoja "${primarySheet}" y deja las demás como soporte.`);
  }

  if (chartTitles.length > 0) {
    hints.push(`Incluye gráficas como ${chartTitles.slice(0, 2).join(' y ')} con una narrativa de hallazgos.`);
  }

  if (findings.length > 0) {
    hints.push('Pide un resumen ejecutivo que priorice alertas, tendencias y recomendaciones accionables.');
  }

  hints.push('Especifica colores, tipo de gráficos, densidad de tablas y tono narrativo en lenguaje libre.');
  return uniqueStrings(hints, 4);
}

function buildExcelIntelligenceReport({
  analysis = {},
  suggestions = [],
  fileName = '',
  fileSizeBytes = 0,
  userPrompt = '',
  model = 'nousresearch/hermes-3-llama-3.1-405b:free',
} = {}) {
  const genericSummary = analysis?.resumen_generico || {};
  const metadata = analysis?.metadatos || {};
  const processing = buildProcessingProfile({ fileSizeBytes, userPrompt });
  const tableCount = collectTableCount(analysis);
  const chartTitles = uniqueStrings((analysis?.graficas_automaticas || []).map((item) => item?.titulo), 4);
  const kpiLabels = uniqueStrings((analysis?.kpis_automaticos || []).map((item) => item?.label || item?.valor), 6);

  const dataset = {
    fileName: fileName || metadata?.archivo || 'Excel cargado',
    sizeMb: Number((fileSizeBytes / (1024 * 1024)).toFixed(1)),
    workbookType: metadata?.tipo_libro || genericSummary?.tipo_libro || 'general',
    primarySheet: genericSummary?.hoja_principal || metadata?.hojas_encontradas?.[0] || null,
    sheetCount: Array.isArray(metadata?.hojas_encontradas) ? metadata.hojas_encontradas.length : 0,
    totalRows: Number(genericSummary?.total_filas || 0),
    totalColumns: Number(genericSummary?.total_columnas || 0),
    familiesDetected: uniqueStrings(metadata?.familias_detectadas || genericSummary?.familias_detectadas || [], 6),
    kpiCount: kpiLabels.length,
    chartCount: chartTitles.length,
    tableCount,
    textBlockCount: Array.isArray(analysis?.bloques_textuales) ? analysis.bloques_textuales.length : 0,
  };

  const keyFindings = buildKeyFindings(analysis);
  const trends = buildTrends(analysis);
  const patterns = buildPatterns(analysis);
  const executiveSummary = uniqueStrings([
    analysis?.briefing_ejecutivo_ia?.de_que_trata,
    analysis?.resumen_ejecutivo_ia?.vision_general,
    keyFindings[0],
  ], 1)[0] || 'La IA completó el análisis del Excel y preparó una narrativa para la presentación.';

  const report = {
    model,
    dataset,
    processing,
    appliedPrompt: String(userPrompt || '').trim(),
    executiveSummary,
    keyFindings,
    trends,
    patterns,
    suggestions: uniqueStrings(suggestions, 5),
    powerPointPlan: {
      palette: inferPalette(dataset),
      preferredCharts: chartTitles,
      preferredTables: uniqueStrings([
        analysis?.muestra_tabla?.hoja_origen,
        ...Object.keys(analysis?.otras_tablas || {}),
        ...Object.keys(analysis?.genericas || {}),
      ], 4),
      recommendedSlides: buildRecommendedSlides(analysis, dataset),
      preferredNarrative: uniqueStrings([
        analysis?.resumen_ejecutivo_ia?.alerta_principal,
        analysis?.resumen_ejecutivo_ia?.recomendacion,
        ...(analysis?.briefing_ejecutivo_ia?.objetivos || []),
      ], 4),
      kpiLabels,
    },
  };

  report.promptHints = buildPromptHints(report);
  report.healthSignals = uniqueStrings([
    processing.skipClientDeepValidation
      ? 'El archivo entra en ruta de validación profunda en servidor para proteger el navegador.'
      : 'El archivo puede validarse parcialmente en navegador antes del envío.',
    processing.allowsExtendedProcessing
      ? `La ruta habilita hasta ${processing.timeoutMinutes} minutos de procesamiento para análisis extensos.`
      : 'El archivo entra en la ventana estándar de procesamiento.',
    dataset.totalRows
      ? `Se detectan ${compactNumber(dataset.totalRows)} filas utilizables en la hoja principal.`
      : null,
  ], 3);

  return report;
}

module.exports = {
  buildExcelIntelligenceReport,
  buildProcessingProfile,
  constants: {
    STANDARD_TIMEOUT_MS,
    EXTENDED_TIMEOUT_MS,
    MAX_TIMEOUT_MS,
    SHALLOW_VALIDATION_THRESHOLD_BYTES,
    EXTENDED_PROCESSING_THRESHOLD_BYTES,
  },
};
