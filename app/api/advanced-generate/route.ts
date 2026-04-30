import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, getMaxExcelUploadSizeMb, validateExcelUpload } from '@/utils/excel-file';
import { ORGANIZER_SCRIPT_NAME, getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';
import panelUtils from '../../../utils/excel-ai-panel.cjs';

const execFileAsync = promisify(execFile);
const { buildProcessingProfile } = panelUtils as {
  buildProcessingProfile: (payload: Record<string, unknown>) => { timeoutMs: number };
};

export const runtime = 'nodejs';
export const maxDuration = 600;

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
type KPIAutoData = {
  label: string;
  value: string;
  importancia?: number;
};
type PandasData = {
  error?: unknown;
  metadatos?: {
    archivo?: string;
    hojas_encontradas?: unknown[];
  };
  resumen_ejecutivo?: {
    valor_total?: number;
    unique_ciudades?: number;
    unique_centros?: number;
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
// ── HELPERS ─────────────────────────────────────────────────────────────────

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
// API HANDLER
// ═══════════════════════════════════════════════════════════════════════════════

export async function POST(req: NextRequest) {
  let tempDir = '';
  let filePath = '';

  try {
    const depStatus = await getRuntimeDependencyStatus(false);
    if (!depStatus.capabilities.analysis) {
      return NextResponse.json({ error: getRuntimeFailureMessage(depStatus, 'analysis') }, { status: 503 });
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
    const userPrompt = String(formData.get('userPrompt') ?? '').trim();

    if (!file) {
      return NextResponse.json({ error: 'No se subió ningún archivo.' }, { status: 400 });
    }

    const validationError = validateExcelUpload(file);
    if (validationError) {
      return NextResponse.json({ error: validationError }, { status: 400 });
    }

    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);
    const processingProfile = buildProcessingProfile({
      fileSizeBytes: file.size,
      userPrompt,
    });

    tempDir = await fs.mkdtemp(path.join(/* turbopackIgnore: true */ os.tmpdir(), 'socya-advanced-'));
    filePath = path.join(tempDir, sanitizeUploadName(file.name));

    await fs.writeFile(/* turbopackIgnore: true */ filePath, buffer);

    try {
      const args = ['-X', 'utf8', ORGANIZER_SCRIPT_NAME, filePath];
      if (userPrompt) {
        args.push(userPrompt);
      }

      const { stdout, stderr } = await execFileAsync('python', args, {
        encoding: 'utf8',
        timeout: processingProfile.timeoutMs,
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
      const slides = buildGenericSlides(pandasData);

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
          { error: 'El analisis del Excel excedio el tiempo permitido. Intenta de nuevo o usa un archivo mas pequeno si el libro es especialmente pesado.' },
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
