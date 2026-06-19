import { execFile } from 'child_process';
import fs from 'fs/promises';
import { promisify } from 'util';

const execFileAsync = promisify(execFile);
const RUNTIME_STATUS_TTL_MS = 60 * 1000;

export const ORGANIZER_SCRIPT_NAME = 'organizer.py';
export const PIPELINE_PACKAGE_DIR = 'socya_pipeline';
export const PIPELINE_PACKAGE_INIT = 'socya_pipeline/__init__.py';
export const PPTX_ANALYZER_SCRIPT_NAME = 'analyze_presentation.py';
export const TEMPLATE_PRESENTATION_NAME = 'Plantilla_Presentacion_Socya (1) (1).pptx';
export type RuntimeCapability = 'analysis' | 'generation';

export interface RuntimeDependencyStatus {
  ok: boolean;
  checkedAt: string;
  capabilities: {
    analysis: boolean;
    generation: boolean;
  };
  python: {
    ok: boolean;
    version: string | null;
    error: string | null;
  };
  scripts: {
    organizer: boolean;
    pipeline: boolean;
    analyzer: boolean;
    template: boolean;
  };
}

let cachedRuntimeStatus: RuntimeDependencyStatus | null = null;
let cachedRuntimeStatusAt = 0;

function nowIso(): string {
  return new Date().toISOString();
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(/* turbopackIgnore: true */ filePath);
    return true;
  } catch {
    return false;
  }
}

async function detectPython(): Promise<RuntimeDependencyStatus['python']> {
  try {
    const result = await execFileAsync('python', ['--version'], {
      encoding: 'utf8',
      timeout: 10 * 1000,
      windowsHide: true,
      env: { ...process.env, PYTHONUTF8: '1' },
    });
    const version = String(result.stdout || result.stderr || '').trim() || null;
    return {
      ok: true,
      version,
      error: null,
    };
  } catch (error: unknown) {
    return {
      ok: false,
      version: null,
      error: error instanceof Error ? error.message : 'No se pudo ejecutar Python.',
    };
  }
}

// NOTE: Antes había un `detectPipeline()` que ejecutaba `python -c "import
// socya_pipeline; from ...; print('ok')"` como pre-flight. Producía dos
// problemas grandes:
//   1. El cold-start de Python en Windows con todos los imports (pandas,
//      openpyxl, python-pptx, matplotlib, requests) podía exceder 30s →
//      timeout → el cache se envenenaba con `pipeline:false` por 30s más,
//      bloqueando TODOS los requests con un mensaje genérico
//      "El backend no tiene todas sus dependencias de analisis operativas".
//   2. Es trabajo duplicado: el comando real (`python -m socya_pipeline ...`)
//      tiene que importar exactamente los mismos módulos. Cualquier error de
//      importación va a aparecer ahí, donde el route ya tiene manejo de
//      errores (parsea stdout JSON, lee stderr, etc.).
//
// Ahora `pipelineOk` se deriva de la presencia del directorio del paquete
// (`socya_pipeline/__init__.py`) — chequeo de filesystem barato (~5ms).
// Si el paquete está en disco pero su importación falla en runtime, el
// comando real fallará con el traceback Python real y el route lo surface.
export async function getRuntimeDependencyStatus(forceRefresh = false): Promise<RuntimeDependencyStatus> {
  const now = Date.now();
  if (!forceRefresh && cachedRuntimeStatus && now - cachedRuntimeStatusAt < RUNTIME_STATUS_TTL_MS) {
    return cachedRuntimeStatus;
  }

  const [python, organizerExists, pipelineExists, analyzerExists, templateExists] = await Promise.all([
    detectPython(),
    fileExists(ORGANIZER_SCRIPT_NAME),
    fileExists(PIPELINE_PACKAGE_INIT),
    fileExists(PPTX_ANALYZER_SCRIPT_NAME),
    fileExists(TEMPLATE_PRESENTATION_NAME),
  ]);

  const analysisReady = python.ok && organizerExists;
  const generationReady = analysisReady && pipelineExists && templateExists;

  const status: RuntimeDependencyStatus = {
    ok: generationReady,
    checkedAt: nowIso(),
    capabilities: {
      analysis: analysisReady,
      generation: generationReady,
    },
    python,
    scripts: {
      organizer: organizerExists,
      pipeline: pipelineExists,
      analyzer: analyzerExists,
      template: templateExists,
    },
  };

  cachedRuntimeStatus = status;
  cachedRuntimeStatusAt = now;
  return status;
}

export function getRuntimeFailureMessage(
  status: RuntimeDependencyStatus,
  capability: RuntimeCapability = 'generation'
): string {
  if (!status.python.ok) {
    return `El runtime de Python no esta disponible: ${status.python.error || 'error desconocido'}`;
  }

  if (!status.scripts.organizer) {
    return 'No se encontro el script organizer.py requerido por el backend.';
  }

  if (capability === 'generation') {
    if (!status.scripts.pipeline) {
      return 'No se encontro el paquete socya_pipeline en el backend.';
    }
    if (!status.scripts.template) {
      return 'No se encontro la plantilla Plantilla_Presentacion_Socya (1) (1).pptx requerida por el backend.';
    }
  }

  return 'El backend no tiene todas sus dependencias operativas disponibles.';
}
