import { NextRequest, NextResponse } from 'next/server';
import { execFile, spawn } from 'child_process';
import { promisify } from 'util';
import { randomUUID } from 'crypto';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, getMaxExcelUploadSizeMb, validateExcelUpload } from '@/utils/excel-file';
import { getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';
import panelUtils from '../../../utils/excel-ai-panel.cjs';
import { makeSSEStream, sseHeaders, coercePhase } from '@/utils/sse-stream';

const execFileAsync = promisify(execFile);

// Spawn Python pipeline with live stderr-progress streaming. Python emits
// `__PROGRESS__:{json}` lines to stderr at major phase boundaries; we parse
// them and forward each one to the SSE stream so the user sees real-time
// progress instead of a frozen "Consultando IA…" message.
//
// Resolves on exit-0 with the captured stdout (the final JSON result).
// Rejects with an Error that carries `.stdout`, `.stderr`, `.code`, `.killed`
// so the caller can replicate the old execFileAsync error contract.
function spawnWithProgress(
  command: string,
  args: string[],
  options: { cwd?: string; env?: NodeJS.ProcessEnv; timeoutMs: number; onProgress: (phase: string, message: string) => void }
): Promise<{ stdout: string; stderr: string }> {
  return new Promise((resolve, reject) => {
    const proc = spawn(command, args, {
      cwd: options.cwd,
      env: options.env,
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });

    let stdout = '';
    let stderrBuf = '';
    let stderrTail = '';  // partial last line waiting for newline
    let killedByTimeout = false;

    const timer = setTimeout(() => {
      killedByTimeout = true;
      proc.kill('SIGKILL');
    }, options.timeoutMs);

    proc.stdout.setEncoding('utf8');
    proc.stdout.on('data', (chunk: string) => {
      stdout += chunk;
    });

    proc.stderr.setEncoding('utf8');
    proc.stderr.on('data', (chunk: string) => {
      stderrBuf += chunk;
      stderrTail += chunk;
      // Process complete lines; keep partial tail for next chunk.
      let nl: number;
      while ((nl = stderrTail.indexOf('\n')) !== -1) {
        const line = stderrTail.slice(0, nl).trimEnd();
        stderrTail = stderrTail.slice(nl + 1);
        if (line.startsWith('__PROGRESS__:')) {
          try {
            const payload = JSON.parse(line.slice('__PROGRESS__:'.length));
            const phase = String(payload?.phase ?? 'progress');
            const message = String(payload?.message ?? '');
            if (message) options.onProgress(phase, message);
          } catch { /* malformed progress — ignore */ }
        }
        // Non-progress stderr lines fall through silently (Python tracebacks
        // already get surfaced via the exit-code path below).
      }
    });

    proc.on('error', (err) => {
      clearTimeout(timer);
      reject(err);
    });

    proc.on('exit', (code) => {
      clearTimeout(timer);
      if (killedByTimeout) {
        const e = new Error(`Python pipeline killed by timeout (${options.timeoutMs}ms)`) as Error & {
          stdout: string; stderr: string; killed: boolean; code: string;
        };
        e.stdout = stdout;
        e.stderr = stderrBuf;
        e.killed = true;
        e.code = 'ETIMEDOUT';
        reject(e);
        return;
      }
      if (code === 0) {
        resolve({ stdout, stderr: stderrBuf });
        return;
      }
      const e = new Error(`Python pipeline exited with code ${code}`) as Error & {
        stdout: string; stderr: string; code: number | null;
      };
      e.stdout = stdout;
      e.stderr = stderrBuf;
      e.code = code;
      reject(e);
    });
  });
}
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
  // PPTX→PNG previews generated alongside the deck. Best-effort:
  // empty array means no previews available (e.g. PowerPoint not installed).
  previewSlides?: string[];   // filenames (no path), inside `previewDir`
  previewDir?: string;        // absolute path to the dir holding the PNGs
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

// Limpieza oportunista: borra entradas PENDING ya expiradas y sus carpetas
// temporales. Sin esto, los tokens que el usuario nunca descarga quedan en
// memoria para siempre (y sus tempDirs en disco) → fuga lenta de memoria y
// disco en servers de larga vida. Lo corremos al inicio de cada POST: es
// barato (recorre un Map chico) y evita un setInterval global, que en dev con
// HMR se duplicaría en cada recompilación.
function sweepExpiredPending(): void {
  const now = Date.now();
  for (const [token, entry] of PENDING) {
    if (entry.expires < now) {
      PENDING.delete(token);
      void fs.rm(entry.tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
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
  // Barre tokens expirados/abandonados antes de empezar (ver sweepExpiredPending).
  sweepExpiredPending();

  // Retry con forceRefresh si el cached status falla — evita 30s de bloqueo
  // por un cold-start que ya se resolvió.
  let depStatus = await getRuntimeDependencyStatus(false);
  if (!depStatus.capabilities.generation) {
    depStatus = await getRuntimeDependencyStatus(true);
  }
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
  const preferredModel = String(formData.get('preferredModel') ?? '').trim() || undefined;
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

  // Map of {originalIndex: nuevoTitulo} editado por el usuario en el panel.
  let slideTitles: Record<number, string> | null = null;
  try {
    const raw = formData.get('slideTitles');
    if (raw) {
      const parsed = JSON.parse(String(raw));
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
        slideTitles = parsed as Record<number, string>;
      }
    }
  } catch { /* ignore */ }

  // Reordenamiento: array de originalIndex en el orden que quiere el usuario.
  let slideOrder: number[] | null = null;
  try {
    const raw = formData.get('slideOrder');
    if (raw) {
      const parsed = JSON.parse(String(raw));
      if (Array.isArray(parsed)) {
        slideOrder = parsed.filter((n) => Number.isFinite(Number(n))).map(Number);
      }
    }
  } catch { /* ignore */ }

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
    // Limpia la carpeta temporal de este request. Se llama en cada salida por
    // error del pipeline (los early-return del catch de abajo) — sin esto los
    // tempDirs de generaciones fallidas se acumulaban en disco (cada Excel que
    // falla dejaba una carpeta socya-pptx-XXXXX huérfana). Solo se omite en el
    // camino feliz, donde el tempDir lo necesita el GET de descarga y se borra
    // recién cuando el usuario baja el archivo o expira el token.
    const cleanupTempDir = async () => {
      if (tempDir) {
        await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
      }
    };
    try {
      send({ phase: 'parsing', step: '1/5', message: 'Subiendo Excel al servidor…' });

      tempDir = await fs.mkdtemp(path.join(/* turbopackIgnore: true */ os.tmpdir(), 'socya-pptx-'));
      inputPath = path.join(tempDir, sanitizeUploadName(fileName));
      outputPath = buildOutputPath(inputPath);
      await fs.writeFile(/* turbopackIgnore: true */ inputPath, Buffer.from(fileBytes));

      // No mandamos updates falsos de planning/validating/rendering aquí —
      // el pipeline Python emite progreso real vía __PROGRESS__:... en
      // stderr y el spawnWithProgress lo reenvía vía SSE. Antes el usuario
      // veía "Consultando IA…" estático durante minutos sin saber qué pasaba.

      const presentationRequest = JSON.stringify({
        prompt: userPrompt,
        audience,
        language,
        current_date: new Date().toLocaleDateString(),
        theme,
        excludeSlideIndices,
        slideTitles,
        slideOrder,
        preferredModel,
      });

      const templatePath = path.join(process.cwd(), 'Plantilla_Presentacion_Socya (1) (1).pptx');
      const args = [
        '-X', 'utf8', '-m', 'socya_pipeline', 'generate',
        '--input', inputPath, '--output', outputPath,
        '--template', templatePath, '--request', presentationRequest,
      ];

      try {
        await spawnWithProgress('python', args, {
          timeoutMs: pythonTimeoutMs,
          env: { ...process.env, PYTHONUTF8: '1', SOCYA_AI_PROFILE: 'patient' },
          onProgress: (phase, message) => {
            // Map Python phases to the 5-step UI in components/GenerationProgress.tsx.
            // 'extracting' (Python step 5) folds into 'validating' so the UI
            // doesn't need to grow a 6th step for a sub-second phase.
            const uiPhase = phase === 'extracting' ? 'validating' : coercePhase(phase);
            send({ phase: uiPhase, message });
          },
        });
      } catch (err: unknown) {
        // The CLI exits with code 2 on PipelineError, writing JSON to stdout.
        const errObj = err as { stdout?: string; stderr?: string; code?: string | number; killed?: boolean };
        if (errObj?.stdout && errObj.stdout.trim().startsWith('{')) {
          try {
            const parsed = JSON.parse(errObj.stdout) as { error?: { message?: string } };
            if (parsed?.error) {
              send({
                phase: 'error',
                message: parsed.error.message ?? 'Error en el pipeline',
                data: parsed.error,
              });
              await cleanupTempDir();
              return;
            }
          } catch { /* fall through */ }
        }
        if (errObj?.code === 'ETIMEDOUT' || errObj?.killed) {
          send({
            phase: 'error',
            message: 'La generación tardó demasiado y fue cancelada. Intenta nuevamente o reduce el tamaño del Excel.',
            data: { code: 'TIMEOUT', user_action: 'retry' },
          });
          await cleanupTempDir();
          return;
        }
        // Surface tail of stderr for non-pipeline errors (e.g. Python crashes).
        const stderrTail = (errObj?.stderr || '').split('\n')
          .filter((l) => l && !l.startsWith('__PROGRESS__:'))
          .slice(-3).join(' | ').slice(0, 500);
        send({
          phase: 'error',
          message: stderrTail
            ? `Error en el pipeline Python: ${stderrTail}`
            : 'Error inesperado en el pipeline Python.',
          data: { code: 'PYTHON_RUNTIME_ERROR', user_action: 'retry' },
        });
        await cleanupTempDir();
        return;
      }

      // Read audit JSON written by the CLI alongside the pptx (best-effort).
      const auditPath = outputPath.replace(/\.pptx$/i, '.audit.json');
      let audit: unknown = null;
      try {
        const auditText = await fs.readFile(/* turbopackIgnore: true */ auditPath, 'utf-8');
        audit = JSON.parse(auditText);
      } catch { /* audit is best-effort */ }

      // Best-effort PPTX → PNG previews. Fails soft on Linux / non-Office
      // boxes. Done synchronously so the SSE 'done' carries the count.
      send({ phase: 'rendering', message: 'Generando vista previa…' });
      const previewDir = path.join(tempDir, 'previews');
      let previewSlides: string[] = [];
      try {
        const previewArgs = [
          '-X', 'utf8', '-m', 'socya_pipeline', 'preview-pptx',
          '--input', outputPath,
          '--output-dir', previewDir,
        ];
        const { stdout } = await execFileAsync('python', previewArgs, {
          encoding: 'utf8',
          timeout: 120 * 1000,
          maxBuffer: 4 * 1024 * 1024,
          windowsHide: true,
          env: { ...process.env, PYTHONUTF8: '1' },
        });
        const parsed = JSON.parse(stdout) as {
          ok: boolean;
          slides?: { index: number; filename: string }[];
        };
        if (parsed.ok && Array.isArray(parsed.slides)) {
          previewSlides = parsed.slides.map(s => s.filename);
        }
      } catch {
        // Preview failed — that's OK, the user can still download blind.
        previewSlides = [];
      }

      const token = randomUUID();
      // TTL inicial: 30 min. El usuario suele revisar el preview, navegar
      // entre slides, abrir/cerrar el lightbox, hablar con alguien... 5 min
      // era muy ajustado y producía "Slide N sin imagen" silencioso. Además
      // /api/pptx-preview hace touch en cada acceso para extender la vida
      // mientras se esté navegando.
      PENDING.set(token, {
        outputPath,
        tempDir,
        filename: path.basename(outputPath),
        expires: Date.now() + 30 * 60_000,
        previewSlides,
        previewDir,
      });

      send({
        phase: 'done',
        message: 'Listo.',
        data: {
          downloadToken: token,
          filename: path.basename(outputPath),
          audit,
          previewCount: previewSlides.length,
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
