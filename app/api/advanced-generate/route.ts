import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, getMaxExcelUploadSizeMb, validateExcelUpload } from '@/utils/excel-file';
import { ORGANIZER_SCRIPT_PATH, PROJECT_ROOT, getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';

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
  labels?: unknown[];
  valores?: number[];
};
type TopSolicitantesData = {
  labels?: unknown[];
  valores?: number[];
  conteos?: unknown[];
};
type PandasData = {
  error?: unknown;
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
  };
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
  resumen_generico?: {
    hoja_principal?: string;
    total_filas?: number;
    total_columnas?: number;
    columnas_numericas?: unknown[];
    columnas?: unknown[];
  };
};

const MAX_MULTIPART_SIZE_BYTES = MAX_EXCEL_UPLOAD_BYTES + 1024 * 1024;
const PYTHON_TIMEOUT_MS = 90 * 1000;

function formatCOP(val?: number): string {
  const safeValue = Number(val ?? 0);
  if (safeValue >= 1000000) return '$' + (safeValue / 1000000).toFixed(1) + 'M';
  if (safeValue >= 1000) return '$' + (safeValue / 1000).toFixed(0) + 'K';
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

function isSpecializedCommissionsDataset(pandasData: PandasData): boolean {
  const headers = pandasData?.muestra_tabla?.encabezados ?? [];
  const normalized = headers.map((header) =>
    String(header).normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase()
  );

  return normalized.includes('solicitante') && normalized.some((header) => header.includes('valor total'));
}

function buildGenericSlides(pandasData: PandasData): SlidePayload[] {
  const slides: SlidePayload[] = [];
  const fileLabel = String(pandasData.metadatos?.archivo || 'Archivo Excel');
  const mainSheetName = String(pandasData.resumen_generico?.hoja_principal || pandasData.metadatos?.hojas_encontradas?.[0] || 'Hoja principal');

  slides.push({
    type: 'title',
    title: 'Reporte Ejecutivo del Archivo',
    subtitle: `Generado a partir de: ${fileLabel}`,
  });

  if (pandasData.metadatos?.hojas_encontradas?.length) {
    slides.push({
      type: 'text_bullets',
      title: 'Estructura Detectada',
      subtitle: 'Hojas identificadas durante la organización del Excel',
      content: pandasData.metadatos.hojas_encontradas.map((sheet) => `Hoja detectada: ${String(sheet)}`),
    });
  }

  if (pandasData.resumen_generico) {
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

  if (pandasData.muestra_tabla) {
    slides.push({
      type: 'table',
      title: `Vista principal: ${mainSheetName}`,
      subtitle: `Datos organizados desde la hoja ${mainSheetName}`,
      content: {
        headers: (pandasData.muestra_tabla.encabezados ?? []).slice(0, 6),
        rows: (pandasData.muestra_tabla.filas ?? []).slice(0, 12).map((row) => row.slice(0, 6)),
      },
    });
  }

  Object.entries(pandasData.otras_tablas ?? {})
    .slice(0, 3)
    .forEach(([name, table]) => {
      slides.push({
        type: 'table',
        title: String(name),
        subtitle: `Tabla organizada desde la hoja ${name}`,
        content: {
          headers: (table.encabezados ?? []).slice(0, 5),
          rows: (table.filas ?? []).slice(0, 10).map((row) => row.slice(0, 5)),
        },
      });
    });

  Object.entries(pandasData.genericas ?? {})
    .slice(0, 3)
    .forEach(([name, table]) => {
      slides.push({
        type: 'table',
        title: String(name),
        subtitle: `Detalle adicional de la hoja ${name}`,
        content: {
          headers: (table.encabezados ?? []).slice(0, 5),
          rows: (table.filas ?? []).slice(0, 10).map((row) => row.slice(0, 5)),
        },
      });
    });

  if (pandasData.coso) {
    slides.push({
      type: 'table',
      title: 'Evaluación COSO',
      subtitle: 'Control interno detectado en el archivo',
      content: {
        headers: (pandasData.coso.encabezados ?? []).slice(0, 5),
        rows: (pandasData.coso.filas ?? []).slice(0, 10).map((row) => row.slice(0, 5)),
      },
    });
  }

  if (pandasData.distribucion_mes) {
    slides.push({
      type: 'table',
      title: 'Distribución Complementaria',
      subtitle: 'Resumen adicional identificado por el organizador',
      content: {
        headers: (pandasData.distribucion_mes.encabezados ?? []).slice(0, 4),
        rows: (pandasData.distribucion_mes.filas ?? []).slice(0, 10).map((row) => row.slice(0, 4)),
      },
    });
  }

  slides.push({
    type: 'closing',
    title: 'Fin del Reporte',
    subtitle: 'La presentación se estructuró con base en las hojas organizadas del Excel.',
  });

  return slides;
}

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

    tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'socya-advanced-'));
    filePath = path.join(tempDir, sanitizeUploadName(file.name));

    await fs.writeFile(filePath, buffer);

    try {
      const { stdout, stderr } = await execFileAsync('python', ['-X', 'utf8', ORGANIZER_SCRIPT_PATH, filePath], {
        cwd: PROJECT_ROOT,
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
      if (!isSpecializedCommissionsDataset(pandasData)) {
        return NextResponse.json(
          { slides: buildGenericSlides(pandasData) },
          {
            headers: {
              'Cache-Control': 'no-store',
            },
          }
        );
      }

      const slides: SlidePayload[] = [];
      const detailLink = 'https://storage.googleapis.com/socya-mock/data.xlsx';

      // ============================================================
      // 1. PORTADA
      // ============================================================
      slides.push({
        type: 'title',
        title: 'Reporte de Auditoría de Comisiones',
        subtitle: `Análisis integral: ${pandasData.metadatos?.archivo || 'Archivo Excel'}`,
      });

      // ============================================================
      // 1.5. INVENTARIO DE DATOS ENCONTRADOS
      // ============================================================
      if (pandasData.metadatos?.hojas_encontradas) {
        slides.push({
          type: 'text_bullets',
          title: 'Estructura del Archivo Detectada',
          subtitle: 'Hojas procesadas para la generación del reporte',
          content: pandasData.metadatos.hojas_encontradas.map((h) => `Hoja detectada: ${String(h)}`)
        });
      }

      // ============================================================
      // 2. RESUMEN EJECUTIVO — KPIs globales
      // ============================================================
      if (pandasData.resumen_ejecutivo) {
        const r = pandasData.resumen_ejecutivo;
        slides.push({
          type: 'kpi_row',
          title: 'Resumen Ejecutivo',
          content: [
            { label: 'Total Comisiones', value: String(r.total_comisiones) },
            { label: 'Valor Total', value: formatCOP(r.valor_total) },
            { label: 'Solicitantes', value: String(r.unique_solicitantes) },
            { label: 'Ciudades Destino', value: String(r.unique_ciudades) },
            { label: 'Centros de Costos', value: String(r.unique_centros) },
          ]
        });
      }

      // ============================================================
      // 3. RESUMEN FINANCIERO — Desglose por tipo de gasto  
      // ============================================================
      if (pandasData.grafica_valores) {
        const kpis = (pandasData.grafica_valores.labels ?? []).map((lbl, idx: number) => {
          const val = pandasData.grafica_valores?.valores?.[idx];
          return { label: String(lbl), value: formatCOP(val) };
        });
        
        // Add promedio if available
        if (pandasData.resumen_ejecutivo) {
          kpis.push({ 
            label: 'Promedio/Comisión', 
            value: formatCOP(pandasData.resumen_ejecutivo.promedio_comision) 
          });
        }
        
        slides.push({
          type: 'kpi_row',
          title: 'Desglose Financiero por Tipo de Gasto',
          content: kpis.slice(0, 5)
        });
      }

      // ============================================================
      // 4. GRÁFICA DE DISTRIBUCIÓN POR ESTADO
      // ============================================================
      if (pandasData.grafica_estados) {
        slides.push({
          type: 'chart',
          title: 'Distribución de Comisiones por Estado',
          content: {
            name: 'Estados',
            labels: pandasData.grafica_estados.labels,
            values: pandasData.grafica_estados.valores,
            barDir: 'bar'
          }
        });
      }

      // ============================================================
      // 5. GRÁFICA TOP CIUDADES DE DESTINO
      // ============================================================
      if (pandasData.grafica_ciudades) {
        slides.push({
          type: 'chart',
          title: 'Top Ciudades de Destino',
          content: {
            name: 'Comisiones',
            labels: pandasData.grafica_ciudades.labels,
            values: pandasData.grafica_ciudades.valores,
            barDir: 'col'
          }
        });
      }

      // ============================================================
      // 6. TOP SOLICITANTES POR VALOR  
      // ============================================================
      if (pandasData.top_solicitantes) {
        const ts = pandasData.top_solicitantes;
        // Show as KPI cards for top 5
        const topKpis = (ts.labels ?? []).slice(0, 5).map((name, idx: number) => ({
          label: String(name),
          value: formatCOP(ts.valores?.[idx] ?? 0),
          subtitle: `${String(ts.conteos?.[idx] ?? 0)} comisiones`
        }));
        slides.push({
          type: 'kpi_row',
          title: 'Top 5 Solicitantes por Valor',
          content: topKpis
        });
      }

      // ============================================================
      // 7. DISTRIBUCIÓN POR CENTRO DE COSTOS
      // ============================================================
      if (pandasData.centros_costos) {
        slides.push({
          type: 'chart',
          title: 'Distribución por Centro de Costos (COP)',
          content: {
            name: 'Valor COP',
            labels: pandasData.centros_costos.labels,
            values: pandasData.centros_costos.valores,
            barDir: 'bar'
          }
        });
      }

      // ============================================================
      // 8-10. TABLAS DE DATOS PAGINADAS (Muestra Principal)
      // ============================================================
      if (pandasData.muestra_tabla) {
        const allRows = pandasData.muestra_tabla.filas ?? [];
        const headers = pandasData.muestra_tabla.encabezados ?? [];
        const ROWS_PER_SLIDE = 12;
        const maxPages = Math.min(Math.ceil(allRows.length / ROWS_PER_SLIDE), 3); // Max 3 pages
        
        for (let page = 0; page < maxPages; page++) {
          const startRow = page * ROWS_PER_SLIDE;
          const endRow = Math.min(startRow + ROWS_PER_SLIDE, allRows.length);
          const pageRows = allRows.slice(startRow, endRow);
          
          const pageLabel = maxPages > 1 ? ` (${page + 1}/${maxPages})` : '';
          slides.push({
            type: 'table',
            title: `Muestra de Comisiones${pageLabel}`,
            subtitle: `Registros ${startRow + 1} - ${endRow} de ${allRows.length}`,
            content: {
              headers: headers,
              rows: pageRows
            },
            detail_link: detailLink
          });
        }
      }

      // ============================================================
      // 11. EVALUACIÓN COSO  
      // ============================================================
      if (pandasData.coso) {
        slides.push({
          type: 'table',
          title: 'Evaluación COSO - Control Interno',
          content: {
            headers: pandasData.coso.encabezados ?? [],
            rows: pandasData.coso.filas ?? []
          }
        });
      }

      // ============================================================
      // 12. HALLAZGOS (tabla completa)
      // ============================================================
      if (pandasData.otras_tablas) {
        // Process ALL tables, not just the first one
        for (const [key, obj] of Object.entries(pandasData.otras_tablas)) {
          const tableName = key.trim();
          
          // Determine if this is "Hallazgos" or "Oportunidades de Mejora"
          const isHallazgo = tableName.toLowerCase().includes('hallazgo');
          const isOpoMejora = tableName.toLowerCase().includes('opo') || tableName.toLowerCase().includes('mejora') || tableName.toLowerCase().includes('oportunidad');
          
          if (isHallazgo) {
            // Hallazgos as a detailed table (show Hallazgo + Riesgo + Estado)
            const displayHeaders = (obj.encabezados ?? []).slice(0, 5);
            const displayRows = (obj.filas ?? []).slice(0, 10).map((row) => row.slice(0, 5));
            
            slides.push({
              type: 'table',
              title: 'Hallazgos de Auditoría',
              content: {
                headers: displayHeaders,
                rows: displayRows
              },
              detail_link: detailLink
            });
            
            // Also create bullet-point summary of key hallazgos
            const bulletItems = (obj.filas ?? [])
              .map((row) => {
                const hallazgo = String(row[0] || '').replace(/_x000d_\\n/g, ' ').substring(0, 120);
                return hallazgo;
              })
              .filter((t: string) => t.length > 5 && t !== '—' && t !== 'ù');

            if (bulletItems.length > 0) {
              slides.push({
                type: 'text_bullets',
                title: 'Resumen de Hallazgos Clave',
                content: bulletItems.slice(0, 8)
              });
            }
          }
          
          if (isOpoMejora) {
            // Oportunidades de Mejora - show with progress info
            const displayHeaders = (obj.encabezados ?? []).slice(0, 4);
            const displayRows = (obj.filas ?? []).slice(0, 12).map((row) => row.slice(0, 4));
            
            slides.push({
              type: 'table',
              title: 'Oportunidades de Mejora',
              content: {
                headers: displayHeaders,
                rows: displayRows
              },
              detail_link: detailLink
            });
            
            // Create progress KPIs from Opo Mejora data
            if (obj.progress) {
              const totalItems = obj.progress.length;
              const completed = obj.progress.filter((p: number) => p >= 1.0).length;
              const inProgress = obj.progress.filter((p: number) => p > 0 && p < 1.0).length;
              const pending = obj.progress.filter((p: number) => p === 0).length;
              const avgProgress = obj.progress.reduce((a: number, b: number) => a + b, 0) / totalItems;
              
              slides.push({
                type: 'kpi_row',
                title: 'Estado de Oportunidades de Mejora',
                content: [
                  { label: 'Total Ítems', value: String(totalItems) },
                  { label: 'Completados', value: String(completed) },
                  { label: 'En Ejecución', value: String(inProgress) },
                  { label: 'Pendientes', value: String(pending) },
                  { label: 'Avance Promedio', value: Math.round(avgProgress * 100) + '%' },
                ]
              });
            }
          }
        }
      }

      // ============================================================
      // 13. HOJAS GENÉRICAS ADICIONALES
      // ============================================================
      if (pandasData.genericas) {
        for (const [name, obj] of Object.entries(pandasData.genericas)) {
          slides.push({
            type: 'table',
            title: `Detalles: ${name}`,
            subtitle: `Datos extraídos de la hoja ${name}`,
            content: {
              headers: obj.encabezados ?? [],
              rows: (obj.filas ?? []).slice(0, 12) // Limit to 12 rows per generic slide
            }
          });
        }
      }

      // ============================================================
      // CIERRE
      // ============================================================
      slides.push({
        type: 'closing',
        title: 'Fin del Reporte',
        subtitle: '¡Gracias por su atención!'
      });

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
      await fs.unlink(filePath).catch(() => {});
    }
    if (tempDir) {
      await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
