import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, getMaxExcelUploadSizeMb, validateExcelUpload } from '@/utils/excel-file';
import { GENERATOR_SCRIPT_NAME, getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';
import panelUtils from '../../../utils/excel-ai-panel.cjs';

const execFileAsync = promisify(execFile);
const { buildProcessingProfile } = panelUtils as {
  buildProcessingProfile: (payload: Record<string, unknown>) => { timeoutMs: number };
};

export const runtime = 'nodejs';
export const maxDuration = 1800;

type VisualMode = 'charts' | 'tables' | 'mixed' | 'boardroom';
type ExecFileError = Error & { code?: string; killed?: boolean; stderr?: string };

const MAX_MULTIPART_SIZE_BYTES = MAX_EXCEL_UPLOAD_BYTES + 1024 * 1024;
function normalizeVisualMode(value: FormDataEntryValue | null): VisualMode {
  const raw = String(value ?? '').trim().toLowerCase();
  if (raw === 'charts' || raw === 'tables' || raw === 'mixed' || raw === 'boardroom') return raw;
  return 'boardroom';
}

function sanitizeUploadName(fileName: string): string {
  const parsed = path.parse(fileName);
  const safeBase = parsed.name.replace(/[^a-zA-Z0-9_-]+/g, '_').replace(/^_+|_+$/g, '') || 'archivo';
  const safeExt = parsed.ext.replace(/[^a-zA-Z0-9.]+/g, '').toLowerCase();
  return `${safeBase}${safeExt}`;
}

function buildOutputPath(inputPath: string): string {
  const ext = path.extname(inputPath);
  const base = path.basename(inputPath, ext).replace(/[^a-zA-Z0-9_-]+/g, '_');
  const outputName = `Presentacion_Ejecutiva_Socya_${base}.pptx`;
  return path.join(path.dirname(inputPath), outputName);
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
    return 'Error generando la presentación premium.';
  }

  const execError = error as ExecFileError;
  const stderrMessage = String(execError.stderr || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .pop();

  return stderrMessage || execError.message || 'Error generando la presentación premium.';
}

function getRelevantPythonWarnings(stderr: string): string[] {
  return stderr
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !/^INFO:\s+Error cargando OpenRouter:/i.test(line))
    .filter((line) => !/^INFO:\s+Error en fallback de OpenRouter/i.test(line));
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
  let outputPath = '';

  try {
    const depStatus = await getRuntimeDependencyStatus(false);
    if (!depStatus.capabilities.generation) {
      return NextResponse.json({ error: getRuntimeFailureMessage(depStatus, 'generation') }, { status: 503 });
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
    const visualMode = normalizeVisualMode(formData.get('visualMode'));
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

    tempDir = await fs.mkdtemp(path.join(/* turbopackIgnore: true */ os.tmpdir(), 'socya-pptx-'));
    inputPath = path.join(tempDir, sanitizeUploadName(file.name));
    outputPath = buildOutputPath(inputPath);

    const bytes = await file.arrayBuffer();
    await fs.writeFile(/* turbopackIgnore: true */ inputPath, Buffer.from(bytes));

    const presentationRequest = JSON.stringify({
      prompt: userPrompt,
      audience,
      language,
      current_date: new Date().toLocaleDateString(),
      theme,
    });
    const args = ['-X', 'utf8', GENERATOR_SCRIPT_NAME, inputPath, outputPath, presentationRequest];

    const { stderr } = await execFileAsync('python', args, {
      encoding: 'utf8',
      maxBuffer: 20 * 1024 * 1024,
      timeout: pythonTimeoutMs,
      windowsHide: true,
      env: {
        ...process.env,
        PYTHONUTF8: '1',
        SOCYA_AI_PROFILE: 'patient',
        SOCYA_PRESENTATION_MODE: visualMode,
      },
    });

    const relevantWarnings = stderr?.trim() ? getRelevantPythonWarnings(stderr) : [];
    if (relevantWarnings.length > 0) {
      console.warn('[generate-pptx] stderr:', relevantWarnings.join('\n'));
    }

    await fs.access(/* turbopackIgnore: true */ outputPath);
    const pptxBuffer = await fs.readFile(/* turbopackIgnore: true */ outputPath);
    const downloadName = path.basename(outputPath);

    return new NextResponse(pptxBuffer, {
      status: 200,
      headers: {
        'Cache-Control': 'no-store',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${downloadName}"`,
      },
    });
  } catch (error: unknown) {
    console.error('[generate-pptx] Error:', error);
    if (isTimedOut(error)) {
      return NextResponse.json(
        { error: 'La generacion del PowerPoint excedio el tiempo permitido. Intenta de nuevo o usa un Excel mas pequeno si el archivo es especialmente pesado.' },
        { status: 504 }
      );
    }
    const message = getExecErrorMessage(error);
    const status = /solo de gráficas|solo gráficas|solo de tablas|solo tablas|datos tabulares válidos|datos válidos|no se puede organizar/i.test(message) ? 422 : 500;
    return NextResponse.json(
      { error: message },
      { status }
    );
  } finally {
    await Promise.all([
      outputPath ? fs.unlink(/* turbopackIgnore: true */ outputPath).catch(() => {}) : Promise.resolve(),
      inputPath ? fs.unlink(/* turbopackIgnore: true */ inputPath).catch(() => {}) : Promise.resolve(),
    ]);

    if (tempDir) {
      await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
