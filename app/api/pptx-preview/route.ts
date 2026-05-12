import { NextRequest, NextResponse } from 'next/server';
import fs from 'fs/promises';
import path from 'path';

// Reuse the same in-memory PENDING map shared with /api/generate-pptx.
// We can't import from the route file directly (would create a runtime
// import cycle in Next), but the route stores the map on globalThis so
// we just reach for the same key here.

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

/**
 * GET /api/pptx-preview?token=<uuid>&slide=<0-based-index>
 *
 * Serves a PNG preview of slide N for the deck identified by `token`. The
 * token is the same one returned by /api/generate-pptx; previews live in
 * `tempDir/previews/` until the deck is downloaded (which deletes tempDir)
 * or 5 minutes pass (whichever comes first).
 */
export async function GET(req: NextRequest) {
  const url = new URL(req.url);
  const token = url.searchParams.get('token');
  const slideRaw = url.searchParams.get('slide');
  const slide = slideRaw ? Number(slideRaw) : NaN;

  const entry = token ? PENDING.get(token) : null;
  if (!entry || entry.expires < Date.now()) {
    return NextResponse.json(
      { error: 'Token expirado o desconocido.' },
      { status: 404 },
    );
  }

  const slides = entry.previewSlides || [];
  if (!entry.previewDir || slides.length === 0) {
    return NextResponse.json(
      { error: 'Vista previa no disponible para este deck.' },
      { status: 404 },
    );
  }
  if (!Number.isInteger(slide) || slide < 0 || slide >= slides.length) {
    return NextResponse.json(
      { error: `Índice de slide fuera de rango (0..${slides.length - 1}).` },
      { status: 400 },
    );
  }

  // Defense: only serve filenames that belong to the previewSlides list,
  // and resolve them inside the previewDir to block any path traversal.
  const fileName = slides[slide];
  const safeName = path.basename(fileName);
  if (safeName !== fileName) {
    return NextResponse.json({ error: 'Nombre de archivo inválido.' }, { status: 400 });
  }
  const fullPath = path.join(entry.previewDir, safeName);
  const resolved = path.resolve(fullPath);
  const baseResolved = path.resolve(entry.previewDir);
  if (!resolved.startsWith(baseResolved)) {
    return NextResponse.json({ error: 'Ruta inválida.' }, { status: 400 });
  }

  try {
    const buf = await fs.readFile(/* turbopackIgnore: true */ resolved);
    return new NextResponse(buf as unknown as BodyInit, {
      status: 200,
      headers: {
        'Content-Type': 'image/png',
        'Cache-Control': 'private, max-age=300',
        'Content-Disposition': `inline; filename="${safeName}"`,
      },
    });
  } catch {
    return NextResponse.json(
      { error: 'No se pudo leer la imagen de preview.' },
      { status: 404 },
    );
  }
}
