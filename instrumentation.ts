// Next.js instrumentation hook — runs once per server instance on boot.
//
// IMPORTANTE (Next.js 16 + Turbopack): este archivo se compila para AMBOS
// runtimes (Node y Edge). Cualquier import directo de un módulo Node-only
// (child_process, fs, etc.) causa un error de compilación del lado Edge,
// AUNQUE el código esté detrás de un guard de runtime. La doc recomienda
// usar require() condicional para que Turbopack solo analice la rama
// del runtime correspondiente:
//   node_modules/next/dist/docs/01-app/03-api-reference/03-file-conventions/instrumentation.md
//
// Por eso el código real vive en `instrumentation-node.ts` (lazy-require).

export function register() {
  if (process.env.NEXT_RUNTIME === 'nodejs') {
    // require condicional → Turbopack solo evalúa este import en el bundle Node.
    // Usamos require() (no `import await`) porque la doc oficial lo exige para
    // que el static analysis del Edge runtime no toque este path.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const { preWarmPython } = require('./instrumentation-node');
    preWarmPython();
  }
  // Edge runtime: no-op. No tenemos nada que pre-warmar en Edge.
}
