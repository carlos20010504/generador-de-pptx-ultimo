import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';

const execFileAsync = promisify(execFile);

type PendingEntry = {
  outputPath: string;
  tempDir: string;
  filename: string;
  expires: number;
  previewSlides?: string[];
  previewDir?: string;
};

const PENDING: Map<string, PendingEntry> = (() => {
  const g = globalThis as Record<string, unknown>;
  if (!g.__SOCYA_PENDING__) {
    g.__SOCYA_PENDING__ = new Map<string, PendingEntry>();
  }
  return g.__SOCYA_PENDING__ as Map<string, PendingEntry>;
})();

export const runtime = 'nodejs';
export const maxDuration = 300;

/**
 * GET /api/pptx-pdf?token=<uuid>
 *
 * On-demand PPTX → PDF conversion (Windows + PowerPoint only). The token is
 * the same one returned by /api/generate-pptx. We don't pre-generate the
 * PDF in the main flow because:
 *   1. PDF export takes ~5–15s and the user may never want it
 *   2. Keeps the SSE done event arriving sooner
 * If the token is expired or PDF rendering fails, returns a JSON error
 * (NOT a partial file) so the client can show a useful message.
 */
export async function GET(req: NextRequest) {
  const url = new URL(req.url);
  const token = url.searchParams.get('token');

  const entry = token ? PENDING.get(token) : null;
  if (!entry || entry.expires < Date.now()) {
    return NextResponse.json(
      { error: 'Token expirado o desconocido.' },
      { status: 404 },
    );
  }

  const pdfPath = entry.outputPath.replace(/\.pptx$/i, '.pdf');

  try {
    const { stdout } = await execFileAsync('python', [
      '-X', 'utf8', '-m', 'socya_pipeline', 'export-pdf',
      '--input', entry.outputPath,
      '--output', pdfPath,
    ], {
      encoding: 'utf8',
      timeout: 5 * 60 * 1000,
      maxBuffer: 4 * 1024 * 1024,
      windowsHide: true,
      env: { ...process.env, PYTHONUTF8: '1' },
    });
    let parsed: { ok?: boolean; path?: string; error?: string } = {};
    try { parsed = JSON.parse(stdout); } catch { /* ignore */ }
    if (!parsed.ok || !parsed.path) {
      return NextResponse.json(
        {
          error: parsed.error
            || 'No se pudo exportar a PDF (PowerPoint no disponible).',
        },
        { status: 503 },
      );
    }
  } catch (err: unknown) {
    return NextResponse.json(
      {
        error: err instanceof Error
          ? `Falló la exportación a PDF: ${err.message}`
          : 'Falló la exportación a PDF.',
      },
      { status: 503 },
    );
  }

  try {
    const buf = await fs.readFile(/* turbopackIgnore: true */ pdfPath);
    const pdfFilename = path.basename(pdfPath);
    return new NextResponse(buf as unknown as BodyInit, {
      status: 200,
      headers: {
        'Cache-Control': 'no-store',
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${pdfFilename}"`,
      },
    });
  } catch {
    return NextResponse.json(
      { error: 'PDF generado pero no se pudo leer del disco.' },
      { status: 500 },
    );
  }
}
