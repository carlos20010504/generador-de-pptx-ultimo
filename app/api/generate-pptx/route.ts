import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, getMaxExcelUploadSizeMb, validateExcelUpload } from '@/utils/excel-file';
import { GENERATOR_SCRIPT_PATH, PROJECT_ROOT, getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';

const execFileAsync = promisify(execFile);

export const runtime = 'nodejs';

type VisualMode = 'charts' | 'tables' | 'mixed';
type ExecFileError = Error & { code?: string; killed?: boolean; stderr?: string };

const MAX_MULTIPART_SIZE_BYTES = MAX_EXCEL_UPLOAD_BYTES + 1024 * 1024;
const GENERATION_TIMEOUT_MS = 2 * 60 * 1000;

function normalizeVisualMode(value: FormDataEntryValue | null): VisualMode {
  const raw = String(value ?? '').trim().toLowerCase();
  if (raw === 'charts' || raw === 'tables') return raw;
  return 'mixed';
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

export async function POST(req: NextRequest) {
  let tempDir = '';
  let inputPath = '';
  let outputPath = '';

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
    const file = formData.get('file');
    const visualMode = normalizeVisualMode(formData.get('visualMode'));

    if (!(file instanceof File)) {
      return NextResponse.json({ error: 'No se subió ningún archivo Excel válido.' }, { status: 400 });
    }

    const validationError = validateExcelUpload(file);
    if (validationError) {
      return NextResponse.json({ error: validationError }, { status: 400 });
    }

    tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'socya-pptx-'));
    inputPath = path.join(tempDir, sanitizeUploadName(file.name));
    outputPath = buildOutputPath(inputPath);

    const bytes = await file.arrayBuffer();
    await fs.writeFile(inputPath, Buffer.from(bytes));

    const { stderr } = await execFileAsync('node', [GENERATOR_SCRIPT_PATH, inputPath], {
      cwd: PROJECT_ROOT,
      encoding: 'utf8',
      maxBuffer: 20 * 1024 * 1024,
      timeout: GENERATION_TIMEOUT_MS,
      windowsHide: true,
      env: {
        ...process.env,
        SOCYA_PRESENTATION_MODE: visualMode,
      },
    });

    if (stderr?.trim()) {
      console.warn('[generate-pptx] stderr:', stderr);
    }

    await fs.access(outputPath);
    const pptxBuffer = await fs.readFile(outputPath);
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
        { error: 'La generacion del PowerPoint excedio el tiempo permitido. Intenta con un Excel mas pequeno o simplificado.' },
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
      outputPath ? fs.unlink(outputPath).catch(() => {}) : Promise.resolve(),
      inputPath ? fs.unlink(inputPath).catch(() => {}) : Promise.resolve(),
    ]);

    if (tempDir) {
      await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
