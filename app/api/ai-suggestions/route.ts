import { NextRequest, NextResponse } from 'next/server';
import { execFile } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import { ORGANIZER_SCRIPT_NAME, getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';

const execFileAsync = promisify(execFile);

export async function POST(req: NextRequest) {
  let tempDir = '';
  let inputPath = '';

  try {
    const depStatus = await getRuntimeDependencyStatus(false);
    if (!depStatus.capabilities.analysis) {
      return NextResponse.json({ error: getRuntimeFailureMessage(depStatus, 'analysis') }, { status: 503 });
    }

    const formData = await req.formData();
    const file = formData.get('file');

    if (!(file instanceof File)) {
      return NextResponse.json({ error: 'No se subió ningún archivo Excel válido.' }, { status: 400 });
    }

    tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'socya-suggestions-'));
    inputPath = path.join(tempDir, file.name.replace(/[^a-zA-Z0-9.-]+/g, '_'));

    const bytes = await file.arrayBuffer();
    await fs.writeFile(inputPath, Buffer.from(bytes));

    const { stdout } = await execFileAsync('python', ['-X', 'utf8', ORGANIZER_SCRIPT_NAME, '--suggestions', inputPath], {
      encoding: 'utf8',
      env: { ...process.env, PYTHONUTF8: '1' },
    });

    const suggestions = JSON.parse(stdout);
    return NextResponse.json(suggestions);

  } catch (error: unknown) {
    console.error('[ai-suggestions] Error:', error);
    return NextResponse.json({ error: 'Error obteniendo sugerencias de la IA.' }, { status: 500 });
  } finally {
    if (tempDir) {
      await fs.rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}
