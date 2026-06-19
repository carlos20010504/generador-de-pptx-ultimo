// Starts `next dev` and opens the browser on 127.0.0.1 (not localhost) so
// Chrome's HSTS for "localhost" can't upgrade to HTTPS. Polls /api/health
// until the server is actually accepting requests before opening.
import { spawn, exec } from 'node:child_process';
import http from 'node:http';

const PORT = process.env.PORT || '3001';
const SERVER_URL = `http://127.0.0.1:${PORT}`;

const nextBinUrl = new URL('../node_modules/next/dist/bin/next', import.meta.url);
const nextBin = decodeURIComponent(nextBinUrl.pathname).replace(/^\/([A-Za-z]:)/, '$1');
const next = spawn(
  process.execPath,
  [nextBin, 'dev', '--webpack', '-p', PORT],
  { stdio: 'inherit', shell: false }
);

next.on('exit', (code) => process.exit(code ?? 0));
process.on('SIGINT', () => next.kill('SIGINT'));
process.on('SIGTERM', () => next.kill('SIGTERM'));

function ping() {
  return new Promise((resolve) => {
    const req = http.get(`${SERVER_URL}/api/health`, { timeout: 1500 }, (res) => {
      res.resume();
      resolve(res.statusCode === 200);
    });
    req.on('error', () => resolve(false));
    req.on('timeout', () => { req.destroy(); resolve(false); });
  });
}

function printBanner(url) {
  // Banner grande, URL en su propia línea para que el terminal la haga clickable.
  const line = '═'.repeat(64);
  console.log(`\n${line}`);
  console.log('  ✓ Server listo. URLs disponibles:');
  console.log('');
  console.log(`     ${url}`);
  console.log(`     http://localhost:${PORT}`);
  console.log('');
  console.log('  → Ctrl+click sobre la URL (o copiala a la barra del browser)');
  console.log('');
  console.log('  Si Chrome no abre y dice "chrome-error://chromewebdata/":');
  console.log('    1) Abrí chrome://net-internals/#hsts');
  console.log('    2) Pegá "localhost" en "Delete domain security policies"');
  console.log('    3) Click Delete, recargá → ya anda');
  console.log('  Alternativa: usá Edge o Firefox (no tienen ese cache).');
  console.log(`${line}\n`);
}

function tryOpen(url) {
  const cmd = process.platform === 'win32'
    ? `start "" "${url}"`
    : process.platform === 'darwin'
      ? `open "${url}"`
      : `xdg-open "${url}"`;
  exec(cmd, () => {});
}

(async () => {
  const deadline = Date.now() + 120_000;
  while (Date.now() < deadline) {
    if (await ping()) {
      printBanner(SERVER_URL);
      tryOpen(SERVER_URL);
      return;
    }
    await new Promise((r) => setTimeout(r, 1000));
  }
  console.error(`\n  ⚠  Server no respondió en 120s — abrí manualmente: ${SERVER_URL}\n`);
})();
