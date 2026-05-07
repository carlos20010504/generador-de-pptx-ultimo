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
const { buildExcelIntelligenceReport, buildProcessingProfile } = panelUtils as {
  buildExcelIntelligenceReport: (payload: Record<string, unknown>) => Record<string, unknown>;
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
    const args = ['-X', 'utf8', ORGANIZER_SCRIPT_NAME, '--panel-report', inputPath, presentationRequest];

    const { stdout, stderr } = await execFileAsync('python', args, {
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

    if (stderr?.trim() && !stdout?.trim()) {
      throw new Error(stderr.trim());
    }

    const payload = JSON.parse(stdout) as {
      error?: string;
      analysis?: Record<string, unknown>;
      suggestions?: string[];
      model?: string;
    };

    if (payload?.error) {
      throw new Error(payload.error);
    }

    const report = buildExcelIntelligenceReport({
      analysis: payload.analysis,
      suggestions: payload.suggestions,
      fileName: file.name,
      fileSizeBytes: file.size,
      userPrompt,
      model: payload.model,
    });

    return NextResponse.json(report, {
      headers: {
        'Cache-Control': 'no-store',
      },
    });
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
