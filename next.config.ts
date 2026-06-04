import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  reactCompiler: true,
  // El launcher (scripts/dev-with-browser.mjs) abre el navegador en
  // http://127.0.0.1:<port> a propósito (evita que Chrome fuerce HTTPS por
  // HSTS sobre "localhost"). Pero Next 16 inicializa el dev server con
  // "localhost" y BLOQUEA por defecto los recursos dev (incluido el websocket
  // de HMR /_next/webpack-hmr) pedidos desde cualquier otro origen. Sin esta
  // línea el HMR queda roto cuando entrás por 127.0.0.1 y el cliente puede
  // forzar un location.reload() al recompilar una ruta on-demand — lo que el
  // usuario percibía como "el programa se cierra mientras la IA hace el plan".
  // Ver node_modules/next/dist/docs/.../allowedDevOrigins.md
  allowedDevOrigins: ['127.0.0.1'],
  outputFileTracingIncludes: {
    // NOTA: se removió './generate_template_presentation.py' — ese archivo no
    // existe en el repo y hacía que `next build` emitiera un warning de
    // tracing por un path inexistente. El route real usa el paquete
    // socya_pipeline + la plantilla .pptx.
    '/api/generate-pptx': [
      './organizer.py',
      './Plantilla_Presentacion_Socya (1) (1).pptx',
    ],
    '/api/advanced-generate': [
      './organizer.py',
    ],
    '/api/excel-intelligence': [
      './organizer.py',
    ],
  },
};

export default nextConfig;
