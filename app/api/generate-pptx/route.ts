import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import { randomUUID } from 'crypto';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, getMaxExcelUploadSizeMb, validateExcelUpload } from '@/utils/excel-file';
import { getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';
import panelUtils from '../../../utils/excel-ai-panel.cjs';
import { makeSSEStream, sseHeaders } from '@/utils/sse-stream';

const execFileAsync = promisify(execFile);
const { buildProcessingProfile } = panelUtils as {
  buildProcessingProfile: (payload: Record<string, unknown>) => { timeoutMs: number };
};

export const runtime = 'nodejs';
export const maxDuration = 1800;

const MAX_MULTIPART_SIZE_BYTES = MAX_EXCEL_UPLOAD_BYTES + 1024 * 1024;

type PendingEntry = {
  outputPath: string;
  tempDir: string;
  filename: string;
  expires: number;
};

// In-memory pending downloads (token → file paths + expiry).
// Stored on globalThis so it survives HMR reloads in dev.
const PENDING: Map<string, PendingEntry> = (() => {
  const g = globalThis as Record<string, unknown>;
  if (!g.__SOCYA_PENDING__) {
    g.__SOCYA_PENDING__ = new Map<string, PendingEntry>();
  }
  return g.__SOCYA_PENDING__ as Map<string, PendingEntry>;
})();

function sanitizeUploadName(fileName: string): string {
  const parsed = path.parse(fileName);
  const safeBase = parsed.name.replace(/[^a-zA-Z0-9_-]+/g, '_').replace(/^_+|_+$/g, '') || 'archivo';
  const safeExt = parsed.ext.replace(/[^a-zA-Z0-9.]+/g, '').toLowerCase();
  return `${safeBase}${safeExt}`;
}

function buildOutputPath(inputPath: string): string {
  const ext = path.extname(inputPath);
  const base = path.basename(inputPath, ext).replace(/[^a-zA-Z0-9_-]+/g, '_');
  return path.join(path.dirname(inputPath), `Presentacion_Ejecutiva_Socya_${base}.pptx`);
}

function buildExtendedTimeoutMs(baseTimeoutMs: number): number {
  const aiWaitBufferMs = 12 * 60 * 1000;
  const routeBudgetMs = 28 * 60 * 1000;
  return Math.max(baseTimeoutMs, Math.min(baseTimeoutMs + aiWaitBufferMs, routeBudgetMs));
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

export async function POST(req: NextRequest) {
  const depStatus = await getRuntimeDependencyStatus(false);
  if (!depStatus.capabilities.generation) {
    return NextResponse.json(
      {
        error: {
          code: 'PYTHON_RUNTIME_ERROR',
          message: getRuntimeFailureMessage(depStatus, 'generation'),
          user_action: 'report_bug',
        },
      },
      { status: 503 }
    );
  }

  const contentLength = Number(req.headers.get('content-length') ?? 0);
  if (contentLength > MAX_MULTIPART_SIZE_BYTES) {
    return NextResponse.json(
      {
        error: {
          code: 'EXCEL_INVALID',
          message: `La solicitud excede ${getMaxExcelUploadSizeMb()} MB.`,
          user_action: 'upload_smaller',
        },
      },
      { status: 413 }
    );
  }

  const formData = await req.formData();
  const file = formData.get('file');
  const userPrompt = String(formData.get('userPrompt') ?? '').trim();
  const audience = String(formData.get('audience') ?? 'ejecutivos').trim();
  const language = String(formData.get('language') ?? 'Español').trim();
  const theme = parseTheme(formData.get('theme'));
  // From the PreparePanel: 0-based indices the user toggled OFF.
  // Forwarded to Python which drops them right before render.
  let excludeSlideIndices: number[] = [];
  try {
    const raw = formData.get('excludeSlideIndices');
    if (raw) {
      const parsed = JSON.parse(String(raw));
      if (Array.isArray(parsed)) {
        excludeSlideIndices = parsed.filter((n) => Number.isFinite(Number(n))).map(Number);
      }
    }
  } catch { /* ignore malformed input — generate everything */ }

  if (!(file instanceof File)) {
    return NextResponse.json(
      {
        error: {
          code: 'EXCEL_INVALID',
          message: 'No se subió ningún archivo Excel válido.',
          user_action: 'upload_again',
        },
      },
      { status: 400 }
    );
  }

  const validationError = validateExcelUpload(file);
  if (validationError) {
    return NextResponse.json(
      { error: { code: 'EXCEL_INVALID', message: validationError, user_action: 'upload_again' } },
      { status: 400 }
    );
  }

  const processingProfile = buildProcessingProfile({ fileSizeBytes: file.size, userPrompt });
  const pythonTimeoutMs = buildExtendedTimeoutMs(processingProfile.timeoutMs);

  // Read file bytes before handing off to background — the File object may not be
  // accessible once the async boundary is crossed.
  const fileBytes = await file.arrayBuffer();
  const fileName = file.name;

  const { stream, send, close } = makeSSEStream();

  // Run generation in background — fire and forget relative to the SSE Response.
  void (async () => {
    let tempDir = '';
    let inputPath = '';
    let outputPath = '';
    try {
      send({ phase: 'parsing', step: '1/5', message: 'Leyendo Excel…' });

      tempDir = await fs.mkdtemp(path.join(/* turbopackIgnore: true */ os.tmpdir(), 'socya-pptx-'));
      inputPath = path.join(tempDir, sanitizeUploadName(fileName));
      outputPath = buildOutputPath(inputPath);
      await fs.writeFile(/* turbopackIgnore: true */ inputPath, Buffer.from(fileBytes));

      send({ phase: 'inventory', step: '2/5', message: 'Construyendo inventario…' });
      send({ phase: 'planning', step: '3/5', message: 'Consultando IA…' });

      const presentationRequest = JSON.stringify({
        prompt: userPrompt,
        audience,
        language,
        current_date: new Date().toLocaleDateString(),
        theme,
        excludeSlideIndices,
      });

      const templatePath = path.join(process.cwd(), 'Plantilla_Presentacion_Socya (1) (1).pptx');
      const args = [
        '-X', 'utf8', '-m', 'socya_pipeline', 'generate',
        '--input', inputPath, '--output', outputPath,
        '--template', templatePath, '--request', presentationRequest,
      ];

      try {
        await execFileAsync('python', args, {
          encoding: 'utf8',
          timeout: pythonTimeoutMs,
          maxBuffer: 20 * 1024 * 1024,
          windowsHide: true,
          env: { ...process.env, PYTHONUTF8: '1', SOCYA_AI_PROFILE: 'patient' },
        });
      } catch (err: unknown) {
        // The CLI exits with code 2 on PipelineError, writing JSON to stdout.
        const errObj = err as { stdout?: string };
        if (errObj?.stdout && errObj.stdout.trim().startsWith('{')) {
          try {
            const parsed = JSON.parse(errObj.stdout) as { error?: { message?: string } };
            if (parsed?.error) {
              send({
                phase: 'error',
                message: parsed.error.message ?? 'Error en el pipeline',
                data: parsed.error,
              });
              return;
            }
          } catch { /* fall through */ }
        }
        throw err;
      }

      send({ phase: 'validating', step: '4/5', message: 'Validando datos…' });
      send({ phase: 'rendering', step: '5/5', message: 'Renderizando PPTX…' });

      // Read audit JSON written by the CLI alongside the pptx (best-effort).
      const auditPath = outputPath.replace(/\.pptx$/i, '.audit.json');
      let audit: unknown = null;
      try {
        const auditText = await fs.readFile(/* turbopackIgnore: true */ auditPath, 'utf-8');
        audit = JSON.parse(auditText);
      } catch { /* audit is best-effort */ }

      const token = randomUUID();
      PENDING.set(token, {
        outputPath,
        tempDir,
        filename: path.basename(outputPath),
        expires: Date.now() + 5 * 60_000,
      });

      send({
        phase: 'done',
        message: 'Listo.',
        data: {
          downloadToken: token,
          filename: path.basename(outputPath),
          audit,
        },
      });
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : 'Error inesperado';
      send({
        phase: 'error',
        message,
        data: { code: 'PYTHON_RUNTIME_ERROR', message, user_action: 'report_bug' },
      });
      // Clean up on error — tempDir cleanup is safe here since no GET will follow.
      if (tempDir) {
        await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
      }
    } finally {
      close();
    }
  })();

  return new Response(stream, { headers: sseHeaders() });
}

export async function GET(req: NextRequest) {
  const token = new URL(req.url).searchParams.get('token');
  const entry = token ? PENDING.get(token) : null;

  if (!entry || entry.expires < Date.now()) {
    return NextResponse.json(
      { error: { code: 'TIMEOUT', message: 'Token expirado o desconocido.', user_action: 'retry' } },
      { status: 404 }
    );
  }

  PENDING.delete(token!);

  try {
    const buffer = await fs.readFile(/* turbopackIgnore: true */ entry.outputPath);
    // Cleanup tempDir asynchronously — after the response is sent.
    void fs.rm(entry.tempDir, { recursive: true, force: true }).catch(() => {});
    return new NextResponse(buffer as unknown as BodyInit, {
      status: 200,
      headers: {
        'Cache-Control': 'no-store',
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': `attachment; filename="${entry.filename}"`,
      },
    });
  } catch {
    return NextResponse.json(
      {
        error: {
          code: 'PYTHON_RUNTIME_ERROR',
          message: 'Archivo no disponible.',
          user_action: 'retry',
        },
      },
      { status: 404 }
    );
  }
}
