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
export const maxDuration = 60;

// Onboarding endpoint: deterministic summary of an uploaded Excel.
// ZERO AI calls — fast (<2s on 10K rows). Used to power the upload preview UI
// before the user commits to a full PowerPoint generation.
export async function POST(req: NextRequest) {
  // Content-Length guard — reject before parsing the formData
  const contentLength = req.headers.get('content-length');
  if (contentLength && Number(contentLength) > MAX_EXCEL_UPLOAD_BYTES * 1.5) {
    return NextResponse.json(
      { error: 'El archivo excede el límite permitido.' },
      { status: 413 }
    );
  }

  // Runtime dependency check — fail fast if Python pipeline isn't reachable
  const runtime = await getRuntimeDependencyStatus();
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
      return NextResponse.json(
        { error: 'Archivo Excel requerido.' },
        { status: 400 }
      );
    }

    const validationError = validateExcelUpload(file);
    if (validationError) {
      return NextResponse.json({ error: validationError }, { status: 400 });
    }

    tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'socya-summary-'));
    const safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_').slice(0, 100) || 'upload.xlsx';
    const inputPath = path.join(tempDir, safeName);
    const buffer = Buffer.from(await file.arrayBuffer());
    await fs.writeFile(inputPath, buffer);

    const { stdout } = await execFileAsync(
      'python',
      ['-m', 'socya_pipeline', 'quick-summary', '--input', inputPath],
      {
        maxBuffer: 10 * 1024 * 1024,
        timeout: 50_000,
        env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' },
        windowsHide: true,
      }
    );

    const summary = JSON.parse(stdout);
    if (summary.error) {
      const errObj = summary.error;
      const message = typeof errObj === 'string'
        ? errObj
        : (errObj?.message || 'No se pudo generar el resumen.');
      return NextResponse.json({ error: message }, { status: 400 });
    }
    return NextResponse.json(summary);
  } catch (err: unknown) {
    const e = err as { stdout?: string; stderr?: string; message?: string };
    // Python may emit a JSON error on stdout with exit code 2
    if (e.stdout) {
      try {
        const parsed = JSON.parse(e.stdout);
        if (parsed.error) {
          return NextResponse.json({ error: parsed.error }, { status: 400 });
        }
      } catch { /* not JSON */ }
    }
    return NextResponse.json(
      { error: 'No se pudo generar el resumen del Excel.' },
      { status: 500 }
    );
  } finally {
    if (tempDir) {
      fs.rm(tempDir, { recursive: true, force: true }).catch(() => undefined);
    }
  }
}
