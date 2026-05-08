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

type ExecFileError = Error & { code?: string; killed?: boolean };

// ── HARD LIMITS ─────────────────────────────────────────────────────────────
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

      const newArgs = ['-X', 'utf8', '-m', 'socya_pipeline', 'plan',
        '--input', filePath, '--request', presentationRequest];
      type ExecWithStdout = Error & { stdout?: string; stderr?: string };
      let newStdout = '';
      try {
        const result = await execFileAsync('python', newArgs, {
          encoding: 'utf8', timeout: pythonTimeoutMs, maxBuffer: 20 * 1024 * 1024,
          windowsHide: true,
          env: { ...process.env, PYTHONUTF8: '1', SOCYA_AI_PROFILE: 'patient' },
        });
        newStdout = result.stdout ?? '';
      } catch (pipelineErr: unknown) {
        // CLI exits with code 2 for structured PipelineErrors — stdout still has valid JSON
        const execErr = pipelineErr as ExecWithStdout;
        if (execErr?.stdout?.trim()) {
          newStdout = execErr.stdout;
        } else {
          throw pipelineErr;
        }
      }
      const parsed = JSON.parse(newStdout);
      const httpStatus = parsed?.error ? 422 : 200;
      return NextResponse.json(parsed, { status: httpStatus, headers: { 'Cache-Control': 'no-store' } });
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
