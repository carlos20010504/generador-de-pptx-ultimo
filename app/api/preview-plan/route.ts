import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { MAX_EXCEL_UPLOAD_BYTES, validateExcelUpload } from '@/utils/excel-file';
import { getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';

const execFileAsync = promisify(execFile);

export const runtime = 'nodejs';
export const maxDuration = 300;

// Preview-plan endpoint: runs the AI planner ONCE (cached by file hash) and
// returns the slide list so the UI can render checkboxes BEFORE the user
// commits to rendering the .pptx. Subsequent /api/generate-pptx calls with
// the same file + prompt will hit the plan cache → zero extra AI tokens.
export async function POST(req: NextRequest) {
  // Pre-warm: ver comentario en /api/quick-summary/route.ts.
  const url = new URL(req.url);
  if (url.searchParams.get('warmup') === '1') {
    return NextResponse.json({ ok: true, warmup: true }, { status: 200 });
  }

  const contentLength = req.headers.get('content-length');
  if (contentLength && Number(contentLength) > MAX_EXCEL_UPLOAD_BYTES * 1.5) {
    return NextResponse.json(
      { error: 'El archivo excede el límite permitido.' },
      { status: 413 }
    );
  }

  // Si el cached status dice "fail", reintentamos UNA vez con forceRefresh —
  // así no nos comemos 30s de error por un cold-start que ya se resolvió.
  let runtime = await getRuntimeDependencyStatus();
  if (!runtime.ok) {
    runtime = await getRuntimeDependencyStatus(true);
  }
  if (!runtime.ok) {
    return NextResponse.json(
      { error: getRuntimeFailureMessage(runtime) },
      { status: 503 }
    );
  }

  let tempDir: string | null = null;
  try {
    const formData = await req.formData();
    const file = formData.get('file');
    if (!(file instanceof File)) {
      return NextResponse.json({ error: 'Archivo Excel requerido.' }, { status: 400 });
    }
    const validationError = validateExcelUpload(file);
    if (validationError) {
      return NextResponse.json({ error: validationError }, { status: 400 });
    }

    tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'socya-preview-'));
    const safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 100) || 'upload.xlsx';
    const inputPath = path.join(tempDir, safeName);
    const buffer = Buffer.from(await file.arrayBuffer());
    await fs.writeFile(inputPath, buffer);

    const userPrompt = String(formData.get('userPrompt') ?? '').trim();
    const audience = String(formData.get('audience') ?? 'ejecutivos').trim();
    const language = String(formData.get('language') ?? 'es').trim();
    const preferredModel = String(formData.get('preferredModel') ?? '').trim() || undefined;

    const requestPayload = JSON.stringify({
      prompt: userPrompt, audience, language, preferredModel,
    });

    const { stdout } = await execFileAsync(
      'python',
      ['-m', 'socya_pipeline', 'plan', '--input', inputPath, '--request', requestPayload],
      {
        maxBuffer: 30 * 1024 * 1024,
        timeout: 280_000,
        // Force UTF-8 stdio so non-ASCII output (em-dashes, narrow-spaces, ñ)
        // doesn't crash on Windows where the default code page is cp1252.
        env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' },
        windowsHide: true,
      }
    );

    const plan = JSON.parse(stdout);
    if (plan.error) {
      // Python emits { error: { code, message, details, user_action } } —
      // surface a clean string for the UI plus keep the typed payload nested.
      const errObj = plan.error;
      const message = typeof errObj === 'string'
        ? errObj
        : (errObj?.message || 'No se pudo generar el plan.');
      return NextResponse.json({ error: message, details: errObj }, { status: 400 });
    }

    // Slim the response to what the UI needs (slide previews + meta)
    const slides = (plan.slides || []).map((s: Record<string, unknown>, i: number) => ({
      index: i,
      type: s.type ?? 'unknown',
      title: s.title ?? (s.data && (s.data as Record<string, unknown>).title) ?? '',
      subtitle: s.subtitle ?? '',
      narrative: s.narrative ?? '',
      provenance: s.provenance ?? null,
      mandatory: s.type === 'title',  // title slide is always kept
    }));

    return NextResponse.json({
      meta: plan.presentation_meta ?? {},
      slides,
      ai_status: plan.ai_status ?? {},
      audit: plan.audit ?? {},
      prompt_suggestions: plan.prompt_suggestions ?? [],
      intent_report: plan.intent_report ?? null,
    });
  } catch (err: unknown) {
    const e = err as { stdout?: string; stderr?: string; code?: string; killed?: boolean; message?: string };
    if (e.code === 'ETIMEDOUT' || e.killed) {
      return NextResponse.json(
        { error: 'La generación del plan demoró demasiado. Intenta con un Excel más pequeño.' },
        { status: 408 }
      );
    }
    if (e.stdout) {
      try {
        const parsed = JSON.parse(e.stdout);
        if (parsed.error) {
          return NextResponse.json({ error: parsed.error }, { status: 400 });
        }
      } catch { /* not JSON */ }
    }
    return NextResponse.json(
      { error: 'No se pudo generar el plan del PowerPoint.' },
      { status: 500 }
    );
  } finally {
    if (tempDir) {
      fs.rm(tempDir, { recursive: true, force: true }).catch(() => undefined);
    }
  }
}
