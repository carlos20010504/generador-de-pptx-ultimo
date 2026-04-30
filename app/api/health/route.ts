import { NextResponse } from 'next/server';
import { getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';

export const runtime = 'nodejs';

export async function GET() {
  const status = await getRuntimeDependencyStatus(true);
  const generationReady = status.capabilities.generation;

  return NextResponse.json(
    {
      ok: generationReady,
      message: generationReady ? 'Backend operativo.' : getRuntimeFailureMessage(status, 'generation'),
      dependencies: status,
    },
    {
      status: generationReady ? 200 : 503,
      headers: {
        'Cache-Control': 'no-store',
      },
    }
  );
}
