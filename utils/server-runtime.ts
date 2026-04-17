import { execFile } from 'child_process';
import fs from 'fs/promises';
import path from 'path';
import { promisify } from 'util';

const execFileAsync = promisify(execFile);
const RUNTIME_STATUS_TTL_MS = 30 * 1000;

export const PROJECT_ROOT = /* turbopackIgnore: true */ process.cwd();
export const ORGANIZER_SCRIPT_PATH = path.join(PROJECT_ROOT, 'organizer.py');
export const GENERATOR_SCRIPT_PATH = path.join(PROJECT_ROOT, 'generate_excel_org_presentation.js');

export interface RuntimeDependencyStatus {
  ok: boolean;
  checkedAt: string;
  python: {
    ok: boolean;
    version: string | null;
    error: string | null;
  };
  scripts: {
    organizer: boolean;
    generator: boolean;
  };
}

let cachedRuntimeStatus: RuntimeDependencyStatus | null = null;
let cachedRuntimeStatusAt = 0;

function nowIso(): string {
  return new Date().toISOString();
}

async function fileExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function detectPython(): Promise<RuntimeDependencyStatus['python']> {
  try {
    const result = await execFileAsync('python', ['--version'], {
      cwd: PROJECT_ROOT,
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

  const [python, organizerExists, generatorExists] = await Promise.all([
    detectPython(),
    fileExists(ORGANIZER_SCRIPT_PATH),
    fileExists(GENERATOR_SCRIPT_PATH),
  ]);

  const status: RuntimeDependencyStatus = {
    ok: python.ok && organizerExists && generatorExists,
    checkedAt: nowIso(),
    python,
    scripts: {
      organizer: organizerExists,
      generator: generatorExists,
    },
  };

  cachedRuntimeStatus = status;
  cachedRuntimeStatusAt = now;
  return status;
}

export function getRuntimeFailureMessage(status: RuntimeDependencyStatus): string {
  if (!status.python.ok) {
    return `El runtime de Python no esta disponible: ${status.python.error || 'error desconocido'}`;
  }

  if (!status.scripts.organizer) {
    return 'No se encontro el script organizer.py requerido por el backend.';
  }

  if (!status.scripts.generator) {
    return 'No se encontro el script generate_excel_org_presentation.js requerido por el backend.';
  }

  return 'El backend no tiene todas sus dependencias operativas disponibles.';
}
