import { execFile } from 'child_process';
import fs from 'fs/promises';
import { promisify } from 'util';

const execFileAsync = promisify(execFile);
const RUNTIME_STATUS_TTL_MS = 30 * 1000;

export const ORGANIZER_SCRIPT_NAME = 'organizer.py';
export const PIPELINE_MODULE = 'socya_pipeline';
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

async function detectPipeline(): Promise<boolean> {
  try {
    await execFileAsync(
      'python',
      ['-c', 'import socya_pipeline; from socya_pipeline import cli, parser, planner, renderer; print("ok")'],
      {
        encoding: 'utf8',
        // 30s — Python cold-start en Windows puede tardar >15s con todos los
        // imports (pandas, openpyxl, prompt_intent stdlib re/unicodedata/difflib).
        // Si el primer detectPipeline timeoutea, el `pipeline:false` se cachea
        // 30s y bloquea TODOS los requests posteriores con el error genérico
        // "El backend no tiene todas sus dependencias de analisis operativas".
        timeout: 30 * 1000,
        windowsHide: true,
        env: { ...process.env, PYTHONUTF8: '1' },
      }
    );
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

export async function getRuntimeDependencyStatus(forceRefresh = false): Promise<RuntimeDependencyStatus> {
  const now = Date.now();
  if (!forceRefresh && cachedRuntimeStatus && now - cachedRuntimeStatusAt < RUNTIME_STATUS_TTL_MS) {
    return cachedRuntimeStatus;
  }

  const [python, organizerExists, pipelineOk, analyzerExists, templateExists] = await Promise.all([
    detectPython(),
    fileExists(ORGANIZER_SCRIPT_NAME),
    detectPipeline(),
    fileExists(PPTX_ANALYZER_SCRIPT_NAME),
    fileExists(TEMPLATE_PRESENTATION_NAME),
  ]);

  const analysisReady = python.ok && organizerExists;
  const generationReady = analysisReady && pipelineOk && templateExists;

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
      pipeline: pipelineOk,
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
  capability: RuntimeCapability = 'analysis'
): string {
  if (!status.python.ok) {
    return `El runtime de Python no esta disponible: ${status.python.error || 'error desconocido'}`;
  }

  if (!status.scripts.organizer) {
    return 'No se encontro el script organizer.py requerido por el backend.';
  }

  if (capability === 'analysis') {
    return 'El backend no tiene todas sus dependencias de analisis operativas disponibles.';
  }

  if (!status.scripts.pipeline) {
    return 'El modulo socya_pipeline no esta disponible en el entorno Python del backend.';
  }

  if (!status.scripts.template) {
    return 'No se encontro la plantilla Plantilla_Presentacion_Socya (1) (1).pptx requerida por el backend.';
  }

  return 'El backend no tiene todas sus dependencias operativas disponibles.';
}
