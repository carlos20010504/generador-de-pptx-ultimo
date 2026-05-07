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
export const maxDuration = 1800;

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
  presentation_request?: {
    prompt?: string;
    audience?: string;
    language?: string;
    theme?: {
      name?: string;
    };
  };
  presentation_design_ia?: {
    presentation_meta?: {
      title?: string;
      subtitle?: string;
    };
    slides?: Array<{
      type?: string;
      title?: string;
      insight_text?: string;
      design_notes?: string;
      cta_text?: string;
      summary_bullets?: string[];
    }>;
  };
  visual_plan_ia?: {
    storyline?: string[];
  };
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

function parseTheme(value: FormDataEntryValue | null) {
  const raw = String(value ?? '').trim();
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch {
    return null;
  }
}

function buildExtendedTimeoutMs(baseTimeoutMs: number): number {
  const aiWaitBufferMs = 12 * 60 * 1000;
  const routeBudgetMs = 28 * 60 * 1000;
  return Math.max(baseTimeoutMs, Math.min(baseTimeoutMs + aiWaitBufferMs, routeBudgetMs));
}

function cleanText(value: unknown, fallback = ''): string {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  return text || fallback;
}

function buildPromptFocusPhrase(pandasData: PandasData): string {
  const prompt = cleanText(pandasData.presentation_request?.prompt);
  const primarySheet = cleanText(pandasData.resumen_generico?.hoja_principal || pandasData.metadatos?.hojas_encontradas?.[0]);
  if (prompt && primarySheet) {
    return `${prompt} · ${primarySheet}`.slice(0, 120);
  }
  return prompt || primarySheet || cleanText(pandasData.metadatos?.archivo, 'datos reales del Excel');
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
  const focusPhrase = buildPromptFocusPhrase(pandasData);
  const planMeta = pandasData.presentation_design_ia?.presentation_meta || {};
  const planSlides = Array.isArray(pandasData.presentation_design_ia?.slides) ? pandasData.presentation_design_ia!.slides! : [];
  const titleSlide = planSlides.find((slide) => slide?.type === 'title_slide');
  const executiveSlide = planSlides.find((slide) => slide?.type === 'executive_summary');
  const closingSlide = planSlides.find((slide) => slide?.type === 'closing_slide');
  const narrativeHints = (pandasData.visual_plan_ia?.storyline || []).map((item) => cleanText(item)).filter(Boolean);

  // 1. PORTADA
  slides.push({
    type: 'title',
    title: cleanText(titleSlide?.title || planMeta.title, `Decision ejecutiva sobre ${focusPhrase}`),
    subtitle: cleanText(planMeta.subtitle, `Lectura guiada por prompt sobre ${fileLabel}`),
  });

  // 2. RESUMEN KPIs — Automáticos
  if (pandasData.kpis_automaticos?.length) {
    slides.push({
      type: 'kpi_row',
      title: cleanText(executiveSlide?.title, `Indicadores para responder al prompt: ${focusPhrase}`),
      content: pandasData.kpis_automaticos.map((kpi) => ({
        label: kpi.label,
        value: kpi.value,
      })),
    });
  } else if (pandasData.resumen_generico) {
    slides.push({
      type: 'kpi_row',
      title: `Base priorizada para ${focusPhrase}`,
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
        title: `Detalle priorizado: ${focusPhrase}${pageLabel}`,
        subtitle: `Registros ${startRow + 1} - ${endRow} de ${allRows.length} en ${mainSheetName}`,
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
        title: `Soporte alineado al prompt: ${String(name)}`,
        subtitle: `Bloque complementario de "${name}" seleccionado desde el Excel`,
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
        title: `Detalle adicional para ${focusPhrase}`,
        subtitle: `Soporte proveniente de la hoja "${name}"`,
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
      title: `Control priorizado para ${focusPhrase}`,
      subtitle: 'Vista de control interno solo si aporta al pedido del usuario',
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
      title: cleanText(closingSlide?.title, `Conclusiones sobre ${focusPhrase}`),
      subtitle: cleanText(executiveSlide?.insight_text, narrativeHints[0] || 'Mensajes clave construidos desde el prompt y los datos'),
      content: pandasData.conclusiones.slice(0, 8),
    });
  }

  // 9. CIERRE
  slides.push({
    type: 'closing',
    title: cleanText(closingSlide?.title, `Siguiente decision sobre ${focusPhrase}`),
    subtitle: cleanText(closingSlide?.cta_text, narrativeHints[1] || 'La salida cambia segun el prompt y la evidencia priorizada del Excel.'),
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
    const audience = String(formData.get('audience') ?? 'ejecutivos').trim();
    const language = String(formData.get('language') ?? 'Español').trim();
    const theme = parseTheme(formData.get('theme'));

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
    const pythonTimeoutMs = buildExtendedTimeoutMs(processingProfile.timeoutMs);

    tempDir = await fs.mkdtemp(path.join(/* turbopackIgnore: true */ os.tmpdir(), 'socya-advanced-'));
    filePath = path.join(tempDir, sanitizeUploadName(file.name));

    await fs.writeFile(/* turbopackIgnore: true */ filePath, buffer);

    try {
      const presentationRequest = JSON.stringify({
        prompt: userPrompt,
        audience,
        language,
        current_date: new Date().toLocaleDateString(),
        theme,
      });
      const args = ['-X', 'utf8', ORGANIZER_SCRIPT_NAME, filePath, presentationRequest];

      const { stdout, stderr } = await execFileAsync('python', args, {
        encoding: 'utf8',
        timeout: pythonTimeoutMs,
        maxBuffer: 20 * 1024 * 1024,
        windowsHide: true,
        env: {
          ...process.env,
          PYTHONUTF8: '1',
          SOCYA_AI_WAIT_ON_RATE_LIMIT: '0',
          SOCYA_AI_MAX_WAIT_SECONDS: '0',
          SOCYA_AI_WAIT_POLL_SECONDS: '15',
          SOCYA_AI_EXECUTION_MODE: 'best_effort',
          SOCYA_AI_HARD_DEADLINE_SECONDS: '35',
        },
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
