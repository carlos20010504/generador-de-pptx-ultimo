import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, getMaxExcelUploadSizeMb, validateExcelUpload } from '@/utils/excel-file';
import { getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';
import panelUtils from '../../../utils/excel-ai-panel.cjs';

const execFileAsync = promisify(execFile);
const { buildProcessingProfile } = panelUtils as {
  buildProcessingProfile: (payload: Record<string, unknown>) => { timeoutMs: number };
};

export const runtime = 'nodejs';
export const maxDuration = 1800;

type ExecFileError = Error & { code?: string; killed?: boolean; stderr?: string };

const MAX_MULTIPART_SIZE_BYTES = MAX_EXCEL_UPLOAD_BYTES + 1024 * 1024;

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

function getExecErrorMessage(error: unknown): string {
  if (!error || typeof error !== 'object') {
    return 'No se pudo completar el analisis del Excel.';
  }

  const execError = error as ExecFileError;
  const stderrMessage = String(execError.stderr || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .pop();

  return stderrMessage || execError.message || 'No se pudo completar el analisis del Excel.';
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

export async function POST(req: NextRequest) {
  let tempDir = '';
  let inputPath = '';

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
    const file = formData.get('file');
    const userPrompt = String(formData.get('userPrompt') ?? '').trim();
    const audience = String(formData.get('audience') ?? 'ejecutivos').trim();
    const language = String(formData.get('language') ?? 'Español').trim();
    const theme = parseTheme(formData.get('theme'));

    if (!(file instanceof File)) {
      return NextResponse.json({ error: 'No se subió ningún archivo Excel válido.' }, { status: 400 });
    }

    const validationError = validateExcelUpload(file);
    if (validationError) {
      return NextResponse.json({ error: validationError }, { status: 400 });
    }

    const processingProfile = buildProcessingProfile({
      fileSizeBytes: file.size,
      userPrompt,
    });
    const pythonTimeoutMs = buildExtendedTimeoutMs(processingProfile.timeoutMs);

    tempDir = await fs.mkdtemp(path.join(/* turbopackIgnore: true */ os.tmpdir(), 'socya-intelligence-'));
    inputPath = path.join(tempDir, sanitizeUploadName(file.name));

    const bytes = await file.arrayBuffer();
    await fs.writeFile(/* turbopackIgnore: true */ inputPath, Buffer.from(bytes));

    const presentationRequest = JSON.stringify({
      prompt: userPrompt,
      audience,
      language,
      current_date: new Date().toLocaleDateString(),
      theme,
    });

    const args = ['-X', 'utf8', '-m', 'socya_pipeline', 'plan',
      '--input', inputPath, '--request', presentationRequest];

    type ExecWithStdout = Error & { stdout?: string; stderr?: string };
    let pipelineStdout = '';
    try {
      const result = await execFileAsync('python', args, {
        encoding: 'utf8',
        timeout: pythonTimeoutMs,
        maxBuffer: 30 * 1024 * 1024,
        windowsHide: true,
        env: {
          ...process.env,
          PYTHONUTF8: '1',
          SOCYA_AI_PROFILE: 'fast',
        },
      });
      pipelineStdout = result.stdout ?? '';
    } catch (pipelineErr: unknown) {
      // CLI exits with code 2 for structured PipelineErrors — stdout still has valid JSON
      const execErr = pipelineErr as ExecWithStdout;
      if (execErr?.stdout?.trim()) {
        pipelineStdout = execErr.stdout;
      } else {
        throw pipelineErr;
      }
    }

    const parsedJson = JSON.parse(pipelineStdout);

    // Structured error from pipeline (exit 2): return 422
    if (parsedJson?.error) {
      return NextResponse.json({ error: parsedJson.error }, { status: 422, headers: { 'Cache-Control': 'no-store' } });
    }

    // Adapter: translate new pipeline shape → existing AIControlPanel.tsx shape
    // (UI gets updated in Phase 4; this keeps the panel working until then)
    const adapter = {
      model: parsedJson.ai_status?.model,
      ai_status: parsedJson.ai_status,
      dataset: {
        fileName: file.name,
        sizeMb: Number((file.size / (1024 * 1024)).toFixed(1)),
        primarySheet: parsedJson.audit?.provenance_per_slide?.[0]?.sheet || null,
        sheetCount: 1,    // unknown without re-parsing; UI will show this in Phase 4
        totalRows: 0,
        totalColumns: 0,
      },
      semanticSummary: {
        topic: parsedJson.presentation_meta?.title || '',
        informationType: '',
        aboutText: parsedJson.presentation_meta?.subtitle || '',
        emphasis: [],
      },
      executiveSummary: parsedJson.presentation_meta?.subtitle || '',
      suggestions: parsedJson.prompt_suggestions || [],
      promptHints: [],
      keyFindings: [],
      trends: [],
      patterns: [],
      powerPointPlan: {
        recommendedSlides: (parsedJson.slides || []).slice(0, 3).map((s: { type?: string; title?: string; narrative?: string; subtitle?: string }) => ({
          type: s.type, title: s.title, reason: s.narrative || s.subtitle || ''
        })),
        preferredCharts: [],
        preferredTables: [],
        palette: { name: 'Analitica Moderna', colors: [] },
        preferredNarrative: [],
        kpiLabels: [],
      },
      processing: { timeoutMinutes: 16, tier: 'standard' },
      healthSignals: parsedJson.audit?.slides_dropped?.length
        ? [`${parsedJson.audit.slides_dropped.length} slides descartados por baja calidad`]
        : [],
    };

    return NextResponse.json(adapter, { headers: { 'Cache-Control': 'no-store' } });
  } catch (error: unknown) {
    console.error('[excel-intelligence] Error:', error);
    if (isTimedOut(error)) {
      return NextResponse.json(
        { error: 'El análisis inteligente del Excel excedió el tiempo permitido para el volumen actual. Intenta de nuevo: la ruta admite procesamiento extendido, pero el archivo requiere más tiempo del disponible.' },
        { status: 504 }
      );
    }

    return NextResponse.json({ error: getExecErrorMessage(error) }, { status: 500 });
  } finally {
    if (inputPath) {
      await fs.unlink(/* turbopackIgnore: true */ inputPath).catch(() => {});
    }
    if (tempDir) {
      await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
