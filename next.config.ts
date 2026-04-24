import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  reactCompiler: true,
  outputFileTracingIncludes: {
    '/api/generate-pptx': [
      './generate_template_presentation.py',
      './organizer.py',
      './Plantilla_Presentacion_Socya (1) (1).pptx',
    ],
    '/api/advanced-generate': [
      './organizer.py',
    ],
  },
};

export default nextConfig;
