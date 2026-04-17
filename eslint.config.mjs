import { defineConfig, globalIgnores } from "eslint/config";
import nextVitals from "eslint-config-next/core-web-vitals";
import nextTs from "eslint-config-next/typescript";

const eslintConfig = defineConfig([
  ...nextVitals,
  ...nextTs,
  {
    files: [
      "utils/advanced-pptx-generator.ts",
      "utils/excel-organizer.ts",
      "utils/excel-parser.ts",
      "utils/pptx-renderer.ts",
      "utils/socya-renderer.ts",
      "__tests__/advanced-pptx-generator.test.ts",
      "__tests__/safe-mode.test.js",
      "utils/presentation-integrity.cjs",
    ],
    rules: {
      "@typescript-eslint/ban-ts-comment": "off",
      "@typescript-eslint/no-explicit-any": "off",
      "@typescript-eslint/no-require-imports": "off",
      "prefer-const": "off",
    },
  },
  // Override default ignores of eslint-config-next.
  globalIgnores([
    // Default ignores of eslint-config-next:
    ".next/**",
    "out/**",
    "build/**",
    "next-env.d.ts",
    "convert.js",
    "generate_excel_org_presentation.js",
    "validate-pipeline.js",
    "validate-quick.js",
    "test-parser-v2.ts",
    "test_render.ts",
  ]),
]);

export default eslintConfig;
