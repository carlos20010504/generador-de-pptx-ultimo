// Node-runtime-only side of instrumentation. Lives in its own file so
// Turbopack (Next.js 16) doesn't try to compile child_process for the
// Edge runtime. `instrumentation.ts` lazy-requires this file only when
// NEXT_RUNTIME === 'nodejs'.
//
// We use this to PRE-WARM the Python pipeline. On Windows (especially with
// Defender's real-time scanner active), the first import of socya_pipeline
// can take 5-20 minutes because every .pyd in pandas / matplotlib / openpyxl
// gets scanned on first read. Subsequent imports hit the OS file cache and
// are sub-second.
//
// Without warm-up, the FIRST user request after `npm run dev` pays the full
// cold-start cost and almost certainly times out. With warm-up, the user
// sees "Analizando archivo…" while Python warms up in the background, and
// by the time they upload anything Python is ready.

import { spawn } from 'child_process';

export function preWarmPython(): void {
  if (process.env.NODE_ENV === 'test') return;
  if (process.env.SOCYA_SKIP_WARMUP === '1') return;

  const startedAt = Date.now();
  console.log('[socya] pre-warming Python pipeline (esto puede tardar varios minutos en el primer arranque)…');

  const child = spawn('python', [
    '-c',
    // Importing cli transitively pulls parser (pandas), planner, renderer
    // (matplotlib + python-pptx), ai_chain (requests), etc. — exactly the
    // surface that the actual /api/preview-plan and /api/generate-pptx
    // calls will need. After this completes, the OS file cache is hot.
    'import socya_pipeline.cli',
  ], {
    windowsHide: true,
    stdio: 'ignore',  // we don't need its output, just want to warm caches
    env: { ...process.env, PYTHONUTF8: '1' },
    detached: false,
  });

  child.on('error', (err) => {
      console.warn('[socya] pre-warm spawn failed (Python no disponible?):', err.message);
  });

  child.on('exit', (code) => {
    const elapsedSec = ((Date.now() - startedAt) / 1000).toFixed(1);
    if (code === 0) {
          console.log(`[socya] pre-warm completado en ${elapsedSec}s — Python listo`);
    } else {
          console.warn(`[socya] pre-warm salió con código ${code} en ${elapsedSec}s — los primeros requests pueden tardar`);
    }
  });
}
