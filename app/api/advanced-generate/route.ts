import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, getMaxExcelUploadSizeMb, validateExcelUpload } from '@/utils/excel-file';
import { ORGANIZER_SCRIPT_NAME, getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';

const execFileAsync = promisify(execFile);

export const runtime = 'nodejs';

type ExecFileError = Error & { code?: string; killed?: boolean };
type SlidePayload = Record<string, unknown>;
type TableData = {
  encabezados?: unknown[];
  filas?: unknown[][];
  progress?: number[];
};
type ChartData = {
  tipo?: string;
  titulo?: string;
  labels?: unknown[];
  valores?: number[];
  colores?: string[];
};
type TopSolicitantesData = {
  labels?: unknown[];
  valores?: number[];
  conteos?: unknown[];
};
type KPIAutoData = {
  label: string;
  value: string;
  importancia?: number;
};
type PandasData = {
  error?: unknown;
  es_comisiones?: boolean;
  metadatos?: {
    archivo?: string;
    hojas_encontradas?: unknown[];
  };
  resumen_ejecutivo?: {
    total_comisiones?: number;
    valor_total?: number;
    unique_solicitantes?: number;
    unique_ciudades?: number;
    unique_centros?: number;
    promedio_comision?: number;
    valor_max_comision?: number;
  };
  resumen_generico?: {
    hoja_principal?: string;
    total_filas?: number;
    total_columnas?: number;
    columnas_numericas?: unknown[];
    columnas?: unknown[];
  };
  kpis_automaticos?: KPIAutoData[];
  graficas_automaticas?: ChartData[];
  conclusiones?: string[];
  presupuesto_slides?: Record<string, number>;
  grafica_valores?: ChartData;
  grafica_estados?: ChartData;
  grafica_ciudades?: ChartData;
  top_solicitantes?: TopSolicitantesData;
  centros_costos?: ChartData;
  muestra_tabla?: TableData;
  coso?: TableData;
  distribucion_mes?: TableData;
  otras_tablas?: Record<string, TableData>;
  genericas?: Record<string, TableData>;
};

// ── HARD LIMITS ─────────────────────────────────────────────────────────────
const MAX_SLIDES = 25;
const ROWS_PER_TABLE_SLIDE = 12;
const MAX_MULTIPART_SIZE_BYTES = MAX_EXCEL_UPLOAD_BYTES + 1024 * 1024;
const PYTHON_TIMEOUT_MS = 90 * 1000;

// ── HELPERS ─────────────────────────────────────────────────────────────────

function formatCOP(val?: number): string {
  const safeValue = Number(val ?? 0);
  if (safeValue >= 1_000_000_000) return '$' + (safeValue / 1_000_000_000).toFixed(1) + 'B';
  if (safeValue >= 1_000_000) return '$' + (safeValue / 1_000_000).toFixed(1) + 'M';
  if (safeValue >= 1_000) return '$' + (safeValue / 1_000).toFixed(0) + 'K';
  return '$' + safeValue.toFixed(0);
}

function sanitizeUploadName(fileName: string): string {
  const parsed = path.parse(fileName);
  const safeBase = parsed.name.replace(/[^a-zA-Z0-9_-]+/g, '_').replace(/^_+|_+$/g, '') || 'archivo';
  const safeExt = parsed.ext.replace(/[^a-zA-Z0-9.]+/g, '').toLowerCase();
  return `${safeBase}${safeExt}`;
}

function isTimedOut(error: unknown): error is ExecFileError {
  return Boolean(
    error &&
    typeof error === 'object' &&
    ('code' in error || 'killed' in error) &&
    (((error as ExecFileError).code === 'ETIMEDOUT') || Boolean((error as ExecFileError).killed))
  );
}

function getErrorMessage(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }
  return fallback;
}

/** Validates that a chart dataset has real data (non-zero values). */
function hasValidChartData(chart: ChartData | undefined): boolean {
  if (!chart?.labels?.length || !chart?.valores?.length) return false;
  return chart.valores.some((v) => typeof v === 'number' && v > 0);
}

/** Validates that a table has real, non-empty data. */
function hasValidTableData(table: TableData | undefined): boolean {
  if (!table?.encabezados?.length || !table?.filas?.length) return false;
  // Check that at least one row has substantive data
  return table.filas.some((row) =>
    row.some((cell) => {
      const s = String(cell ?? '').trim();
      return s !== '' && s !== '—' && s !== '0' && s !== 'nan' && s !== 'None';
    })
  );
}

function isSpecializedCommissionsDataset(pandasData: PandasData): boolean {
  return pandasData.es_comisiones === true;
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE BUILDER — GENERIC PATH (ANY Excel)
// ═══════════════════════════════════════════════════════════════════════════════

function buildGenericSlides(pandasData: PandasData): SlidePayload[] {
  const slides: SlidePayload[] = [];
  const fileLabel = String(pandasData.metadatos?.archivo || 'Archivo Excel');
  const mainSheetName = String(pandasData.resumen_generico?.hoja_principal || pandasData.metadatos?.hojas_encontradas?.[0] || 'Hoja principal');

  // 1. PORTADA
  slides.push({
    type: 'title',
    title: 'Reporte Ejecutivo',
    subtitle: `Análisis inteligente: ${fileLabel}`,
  });

  // 2. RESUMEN KPIs — Automáticos
  if (pandasData.kpis_automaticos?.length) {
    slides.push({
      type: 'kpi_row',
      title: 'Indicadores Clave Detectados',
      content: pandasData.kpis_automaticos.map((kpi) => ({
        label: kpi.label,
        value: kpi.value,
      })),
    });
  } else if (pandasData.resumen_generico) {
    slides.push({
      type: 'kpi_row',
      title: 'Resumen del Archivo',
      content: [
        { label: 'Hoja base', value: mainSheetName },
        { label: 'Filas', value: String(pandasData.resumen_generico.total_filas ?? 0) },
        { label: 'Columnas', value: String(pandasData.resumen_generico.total_columnas ?? 0) },
        { label: 'Columnas numéricas', value: String((pandasData.resumen_generico.columnas_numericas ?? []).length) },
      ],
    });
  }

  // 3. GRÁFICAS AUTOMÁTICAS
  const autoCharts = pandasData.graficas_automaticas ?? [];
  for (const chart of autoCharts.slice(0, 3)) {
    if (hasValidChartData(chart)) {
      slides.push({
        type: 'chart',
        title: chart.titulo || 'Análisis Visual',
        content: {
          name: chart.titulo || 'Dato',
          labels: chart.labels,
          values: chart.valores,
          barDir: chart.tipo === 'pie' ? undefined : 'col',
        },
      });
    }
  }

  // 4. TABLA PRINCIPAL
  if (hasValidTableData(pandasData.muestra_tabla)) {
    const allRows = pandasData.muestra_tabla!.filas!;
    const headers = pandasData.muestra_tabla!.encabezados!;
    const maxPages = Math.min(Math.ceil(allRows.length / ROWS_PER_TABLE_SLIDE), 2);

    for (let page = 0; page < maxPages && slides.length < MAX_SLIDES - 3; page++) {
      const startRow = page * ROWS_PER_TABLE_SLIDE;
      const endRow = Math.min(startRow + ROWS_PER_TABLE_SLIDE, allRows.length);
      const pageRows = allRows.slice(startRow, endRow);
      const pageLabel = maxPages > 1 ? ` (${page + 1}/${maxPages})` : '';

      slides.push({
        type: 'table',
        title: `Vista principal: ${mainSheetName}${pageLabel}`,
        subtitle: `Registros ${startRow + 1} - ${endRow} de ${allRows.length}`,
        content: {
          headers: headers.slice(0, 7),
          rows: pageRows.map((row) => row.slice(0, 7)),
        },
      });
    }
  }

  // 5. OTRAS TABLAS (Hallazgos, Oportunidades de Mejora, etc.)
  for (const [name, table] of Object.entries(pandasData.otras_tablas ?? {}).slice(0, 3)) {
    if (slides.length >= MAX_SLIDES - 2) break;
    if (hasValidTableData(table)) {
      slides.push({
        type: 'table',
        title: String(name),
        subtitle: `Datos extraídos de la hoja "${name}"`,
        content: {
          headers: (table.encabezados ?? []).slice(0, 6),
          rows: (table.filas ?? []).slice(0, ROWS_PER_TABLE_SLIDE).map((row) => row.slice(0, 6)),
        },
      });
    }
  }

  // 6. HOJAS GENÉRICAS
  for (const [name, table] of Object.entries(pandasData.genericas ?? {}).slice(0, 3)) {
    if (slides.length >= MAX_SLIDES - 2) break;
    if (hasValidTableData(table)) {
      slides.push({
        type: 'table',
        title: String(name),
        subtitle: `Detalle adicional de la hoja "${name}"`,
        content: {
          headers: (table.encabezados ?? []).slice(0, 6),
          rows: (table.filas ?? []).slice(0, ROWS_PER_TABLE_SLIDE).map((row) => row.slice(0, 6)),
        },
      });
    }
  }

  // 7. COSO
  if (slides.length < MAX_SLIDES - 2 && hasValidTableData(pandasData.coso)) {
    slides.push({
      type: 'table',
      title: 'Evaluación COSO',
      subtitle: 'Control interno detectado en el archivo',
      content: {
        headers: (pandasData.coso!.encabezados ?? []).slice(0, 5),
        rows: (pandasData.coso!.filas ?? []).slice(0, 10).map((row) => row.slice(0, 5)),
      },
    });
  }

  // 8. CONCLUSIONES
  if (slides.length < MAX_SLIDES - 1 && pandasData.conclusiones?.length) {
    slides.push({
      type: 'text_bullets',
      title: 'Conclusiones del Análisis',
      subtitle: 'Hallazgos identificados automáticamente a partir de los datos',
      content: pandasData.conclusiones.slice(0, 8),
    });
  }

  // 9. CIERRE
  slides.push({
    type: 'closing',
    title: 'Fin del Reporte',
    subtitle: 'La presentación se estructuró con base en las hojas organizadas del Excel.',
  });

  // ENFORCE: never exceed MAX_SLIDES
  return slides.slice(0, MAX_SLIDES);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE BUILDER — COMMISSIONS PATH (Specialized)
// ═══════════════════════════════════════════════════════════════════════════════

function buildCommissionSlides(pandasData: PandasData): SlidePayload[] {
  const slides: SlidePayload[] = [];

  // Helper: only push if we have room
  function pushSlide(slide: SlidePayload): boolean {
    if (slides.length >= MAX_SLIDES - 1) return false; // reserve 1 for closing
    slides.push(slide);
    return true;
  }

  // ── 1. PORTADA ───────────────────────────────────────────────────────────
  pushSlide({
    type: 'title',
    title: 'Reporte de Auditoría de Comisiones',
    subtitle: `Análisis integral: ${pandasData.metadatos?.archivo || 'Archivo Excel'}`,
  });

  // ── 2. ESTRUCTURA DEL ARCHIVO ────────────────────────────────────────────
  const budget = pandasData.presupuesto_slides ?? {};
  if (budget.estructura && pandasData.metadatos?.hojas_encontradas?.length) {
    pushSlide({
      type: 'text_bullets',
      title: 'Estructura del Archivo Detectada',
      subtitle: 'Hojas procesadas para la generación del reporte',
      content: pandasData.metadatos.hojas_encontradas.map((h) => `Hoja: ${String(h)}`),
    });
  }

  // ── 3. RESUMEN EJECUTIVO — KPIs globales ─────────────────────────────────
  if (pandasData.resumen_ejecutivo) {
    const r = pandasData.resumen_ejecutivo;
    pushSlide({
      type: 'kpi_row',
      title: 'Resumen Ejecutivo',
      content: [
        { label: 'Total Comisiones', value: String(r.total_comisiones) },
        { label: 'Valor Total', value: formatCOP(r.valor_total) },
        { label: 'Solicitantes', value: String(r.unique_solicitantes) },
        { label: 'Ciudades Destino', value: String(r.unique_ciudades) },
        { label: 'Centros de Costos', value: String(r.unique_centros) },
      ],
    });
  }

  // ── 4. DESGLOSE FINANCIERO ────────────────────────────────────────────────
  if (hasValidChartData(pandasData.grafica_valores)) {
    const kpis = (pandasData.grafica_valores!.labels ?? []).map((lbl, idx: number) => {
      const val = pandasData.grafica_valores?.valores?.[idx];
      return { label: String(lbl), value: formatCOP(val) };
    });
    if (pandasData.resumen_ejecutivo) {
      kpis.push({
        label: 'Promedio/Comisión',
        value: formatCOP(pandasData.resumen_ejecutivo.promedio_comision),
      });
    }
    pushSlide({
      type: 'kpi_row',
      title: 'Desglose Financiero por Tipo de Gasto',
      content: kpis.slice(0, 5),
    });
  }

  // ── 5. GRÁFICA DISTRIBUCIÓN POR ESTADO ────────────────────────────────────
  if (hasValidChartData(pandasData.grafica_estados)) {
    pushSlide({
      type: 'chart',
      title: 'Distribución de Comisiones por Estado',
      content: {
        name: 'Estados',
        labels: pandasData.grafica_estados!.labels,
        values: pandasData.grafica_estados!.valores,
        barDir: 'bar',
      },
    });
  }

  // ── 6. TOP CIUDADES ───────────────────────────────────────────────────────
  if (hasValidChartData(pandasData.grafica_ciudades)) {
    pushSlide({
      type: 'chart',
      title: 'Top Ciudades de Destino',
      content: {
        name: 'Comisiones',
        labels: pandasData.grafica_ciudades!.labels,
        values: pandasData.grafica_ciudades!.valores,
        barDir: 'col',
      },
    });
  }

  // ── 7. TOP SOLICITANTES ───────────────────────────────────────────────────
  if (pandasData.top_solicitantes) {
    const ts = pandasData.top_solicitantes;
    const topKpis = (ts.labels ?? []).slice(0, 5).map((name, idx: number) => ({
      label: String(name),
      value: formatCOP(ts.valores?.[idx] ?? 0),
      subtitle: `${String(ts.conteos?.[idx] ?? 0)} comisiones`,
    }));
    if (topKpis.length > 0) {
      pushSlide({
        type: 'kpi_row',
        title: 'Top 5 Solicitantes por Valor',
        content: topKpis,
      });
    }
  }

  // ── 8. CENTROS DE COSTOS ──────────────────────────────────────────────────
  if (hasValidChartData(pandasData.centros_costos)) {
    pushSlide({
      type: 'chart',
      title: 'Distribución por Centro de Costos (COP)',
      content: {
        name: 'Valor COP',
        labels: pandasData.centros_costos!.labels,
        values: pandasData.centros_costos!.valores,
        barDir: 'bar',
      },
    });
  }

  // ── 9. TABLA PRINCIPAL — PAGINADA ─────────────────────────────────────────
  if (hasValidTableData(pandasData.muestra_tabla)) {
    const allRows = pandasData.muestra_tabla!.filas!;
    const headers = pandasData.muestra_tabla!.encabezados!;
    const maxPages = Math.min(
      Math.ceil(allRows.length / ROWS_PER_TABLE_SLIDE),
      budget.tabla_principal || 2
    );

    for (let page = 0; page < maxPages; page++) {
      if (slides.length >= MAX_SLIDES - 3) break; // reserve room for closing + conclusions + hallazgos
      const startRow = page * ROWS_PER_TABLE_SLIDE;
      const endRow = Math.min(startRow + ROWS_PER_TABLE_SLIDE, allRows.length);
      const pageRows = allRows.slice(startRow, endRow);
      const pageLabel = maxPages > 1 ? ` (${page + 1}/${maxPages})` : '';

      pushSlide({
        type: 'table',
        title: `Muestra de Comisiones${pageLabel}`,
        subtitle: `Registros ${startRow + 1} - ${endRow} de ${allRows.length}`,
        content: {
          headers: headers,
          rows: pageRows,
        },
      });
    }
  }

  // ── 10. COSO ──────────────────────────────────────────────────────────────
  if (hasValidTableData(pandasData.coso)) {
    pushSlide({
      type: 'table',
      title: 'Evaluación COSO - Control Interno',
      content: {
        headers: pandasData.coso!.encabezados ?? [],
        rows: (pandasData.coso!.filas ?? []).slice(0, ROWS_PER_TABLE_SLIDE),
      },
    });
  }

  // ── 11. HALLAZGOS Y OPORTUNIDADES DE MEJORA ───────────────────────────────
  if (pandasData.otras_tablas) {
    for (const [key, obj] of Object.entries(pandasData.otras_tablas)) {
      if (slides.length >= MAX_SLIDES - 2) break;
      const tableName = key.trim();
      const isHallazgo = tableName.toLowerCase().includes('hallazgo');
      const isOpoMejora = tableName.toLowerCase().includes('opo') || tableName.toLowerCase().includes('mejora') || tableName.toLowerCase().includes('oportunidad');

      if (isHallazgo && hasValidTableData(obj)) {
        const displayHeaders = (obj.encabezados ?? []).slice(0, 5);
        const displayRows = (obj.filas ?? []).slice(0, 10).map((row) => row.slice(0, 5));

        pushSlide({
          type: 'table',
          title: 'Hallazgos de Auditoría',
          content: { headers: displayHeaders, rows: displayRows },
        });

        // Resumen de hallazgos en bullets
        if (slides.length < MAX_SLIDES - 2) {
          const bulletItems = (obj.filas ?? [])
            .map((row) => {
              const hallazgo = String(row[0] || '').replace(/_x000d_\\n/g, ' ').substring(0, 120);
              return hallazgo;
            })
            .filter((t: string) => t.length > 5 && t !== '—' && t !== 'ù');

          if (bulletItems.length > 0) {
            pushSlide({
              type: 'text_bullets',
              title: 'Resumen de Hallazgos Clave',
              content: bulletItems.slice(0, 8),
            });
          }
        }
      }

      if (isOpoMejora && hasValidTableData(obj)) {
        const displayHeaders = (obj.encabezados ?? []).slice(0, 4);
        const displayRows = (obj.filas ?? []).slice(0, ROWS_PER_TABLE_SLIDE).map((row) => row.slice(0, 4));

        pushSlide({
          type: 'table',
          title: 'Oportunidades de Mejora',
          content: { headers: displayHeaders, rows: displayRows },
        });

        // Progress KPIs
        if (obj.progress && slides.length < MAX_SLIDES - 2) {
          const totalItems = obj.progress.length;
          const completed = obj.progress.filter((p: number) => p >= 1.0).length;
          const inProgress = obj.progress.filter((p: number) => p > 0 && p < 1.0).length;
          const pending = obj.progress.filter((p: number) => p === 0).length;
          const avgProgress = obj.progress.reduce((a: number, b: number) => a + b, 0) / totalItems;

          pushSlide({
            type: 'kpi_row',
            title: 'Estado de Oportunidades de Mejora',
            content: [
              { label: 'Total Ítems', value: String(totalItems) },
              { label: 'Completados', value: String(completed) },
              { label: 'En Ejecución', value: String(inProgress) },
              { label: 'Pendientes', value: String(pending) },
              { label: 'Avance Promedio', value: Math.round(avgProgress * 100) + '%' },
            ],
          });
        }
      }
    }
  }

  // ── 12. HOJAS GENÉRICAS ───────────────────────────────────────────────────
  if (pandasData.genericas) {
    for (const [name, obj] of Object.entries(pandasData.genericas)) {
      if (slides.length >= MAX_SLIDES - 2) break;
      if (hasValidTableData(obj)) {
        pushSlide({
          type: 'table',
          title: `Detalles: ${name}`,
          subtitle: `Datos extraídos de la hoja "${name}"`,
          content: {
            headers: (obj.encabezados ?? []).slice(0, 6),
            rows: (obj.filas ?? []).slice(0, ROWS_PER_TABLE_SLIDE).map((row) => row.slice(0, 6)),
          },
        });
      }
    }
  }

  // ── 13. CONCLUSIONES ──────────────────────────────────────────────────────
  if (pandasData.conclusiones?.length && slides.length < MAX_SLIDES - 1) {
    pushSlide({
      type: 'text_bullets',
      title: 'Conclusiones del Análisis',
      subtitle: 'Hallazgos identificados automáticamente a partir de los datos reales',
      content: pandasData.conclusiones.slice(0, 8),
    });
  }

  // ── 14. CIERRE ────────────────────────────────────────────────────────────
  slides.push({
    type: 'closing',
    title: 'Fin del Reporte',
    subtitle: '¡Gracias por su atención!',
  });

  // HARD ENFORCE: never exceed MAX_SLIDES
  return slides.slice(0, MAX_SLIDES);
}

// ═══════════════════════════════════════════════════════════════════════════════
// API HANDLER
// ═══════════════════════════════════════════════════════════════════════════════

export async function POST(req: NextRequest) {
  let tempDir = '';
  let filePath = '';

  try {
    const depStatus = await getRuntimeDependencyStatus(false);
    if (!depStatus.ok) {
      return NextResponse.json({ error: getRuntimeFailureMessage(depStatus) }, { status: 503 });
    }

    const contentLength = Number(req.headers.get('content-length') ?? 0);
    if (contentLength > MAX_MULTIPART_SIZE_BYTES) {
      return NextResponse.json(
        { error: `La solicitud excede el limite permitido de ${getMaxExcelUploadSizeMb()} MB.` },
        { status: 413 }
      );
    }

    const formData = await req.formData();
    const file = formData.get('file') as File | null;

    if (!file) {
      return NextResponse.json({ error: 'No se subió ningún archivo.' }, { status: 400 });
    }

    const validationError = validateExcelUpload(file);
    if (validationError) {
      return NextResponse.json({ error: validationError }, { status: 400 });
    }

    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);

    tempDir = await fs.mkdtemp(path.join(/* turbopackIgnore: true */ os.tmpdir(), 'socya-advanced-'));
    filePath = path.join(tempDir, sanitizeUploadName(file.name));

    await fs.writeFile(/* turbopackIgnore: true */ filePath, buffer);

    try {
      const { stdout, stderr } = await execFileAsync('python', ['-X', 'utf8', ORGANIZER_SCRIPT_NAME, filePath], {
        encoding: 'utf8',
        timeout: PYTHON_TIMEOUT_MS,
        maxBuffer: 20 * 1024 * 1024,
        windowsHide: true,
        env: { ...process.env, PYTHONUTF8: '1' },
      });

      if (stderr?.trim() && !stdout?.trim()) {
        console.error('Python Stderr:', stderr);
        throw new Error('Error al procesar los datos con Python.');
      }

      const pandasData = JSON.parse(stdout) as PandasData;
      if (pandasData?.error) {
        throw new Error(String(pandasData.error));
      }

      // Determine which builder to use based on Python analysis
      const slides = isSpecializedCommissionsDataset(pandasData)
        ? buildCommissionSlides(pandasData)
        : buildGenericSlides(pandasData);

      console.log(`[advanced-generate] Generated ${slides.length}/${MAX_SLIDES} slides for: ${file.name}`);

      return NextResponse.json(
        { slides },
        {
          headers: {
            'Cache-Control': 'no-store',
          },
        }
      );
    } catch (err: unknown) {
      console.error('Error in Python execution:', err);
      if (isTimedOut(err)) {
        return NextResponse.json(
          { error: 'El analisis del Excel excedio el tiempo permitido. Intenta con un archivo mas pequeno o simplificado.' },
          { status: 504 }
        );
      }
      return NextResponse.json({ error: `Error procesando el Excel con pandas: ${getErrorMessage(err, 'Error interno del servidor.')}` }, { status: 500 });
    }
  } catch (error: unknown) {
    console.error('API Error:', error);
    return NextResponse.json({ error: getErrorMessage(error, 'Error interno del servidor.') }, { status: 500 });
  } finally {
    if (filePath) {
      await fs.unlink(/* turbopackIgnore: true */ filePath).catch(() => {});
    }
    if (tempDir) {
      await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
