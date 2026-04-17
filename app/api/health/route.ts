import { NextResponse } from 'next/server';
import { getRuntimeDependencyStatus, getRuntimeFailureMessage } from '@/utils/server-runtime';

export const runtime = 'nodejs';

export async function GET() {
  const status = await getRuntimeDependencyStatus(true);

  return NextResponse.json(
    {
      ok: status.ok,
      message: status.ok ? 'Backend operativo.' : getRuntimeFailureMessage(status),
      dependencies: status,
    },
    {
      status: status.ok ? 200 : 503,
      headers: {
        'Cache-Control': 'no-store',
      },
    }
  );
}
