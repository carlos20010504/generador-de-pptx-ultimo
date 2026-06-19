export type SSEPhase =
  | 'parsing' | 'inventory' | 'planning' | 'extracting'
  | 'validating' | 'rendering' | 'done' | 'error' | 'progress';

export type SSEEvent = {
  phase: SSEPhase;
  step?: string;
  message?: string;
  data?: unknown;
};

const KNOWN_PHASES = new Set<SSEPhase>([
  'parsing', 'inventory', 'planning', 'extracting',
  'validating', 'rendering', 'done', 'error', 'progress',
]);

// Coerce an arbitrary string (from Python progress events) to a valid phase.
// Unknown phases collapse to the generic 'progress' bucket so the UI still
// renders the message instead of crashing on a typo.
export function coercePhase(raw: string): SSEPhase {
  return KNOWN_PHASES.has(raw as SSEPhase) ? (raw as SSEPhase) : 'progress';
}

export function makeSSEStream(): {
  stream: ReadableStream<Uint8Array>;
  send: (event: SSEEvent) => void;
  close: () => void;
} {
  const encoder = new TextEncoder();
  let controller: ReadableStreamDefaultController<Uint8Array>;
  let closed = false;
  const stream = new ReadableStream<Uint8Array>({
    start(c) { controller = c; },
  });
  const send = (event: SSEEvent) => {
    if (closed) return;
    try {
      const payload = `data: ${JSON.stringify(event)}\n\n`;
      controller.enqueue(encoder.encode(payload));
    } catch {
      // Controller may have been closed by client disconnect — swallow so
      // the producer doesn't crash mid-stream.
      closed = true;
    }
  };
  // Heartbeat: SSE comments (lines starting with ':') are ignored by the
  // EventSource client but keep the TCP connection alive. Without these,
  // long Python phases (cold-start parse_workbook, AI calls) can leave the
  // connection silent for >30s and Windows/the browser resets it with
  // ERR_CONNECTION_RESET, surfacing as "Failed to fetch" in the UI.
  const HEARTBEAT_INTERVAL_MS = 8000;
  const heartbeatTimer = setInterval(() => {
    if (closed) return;
    try {
      controller.enqueue(encoder.encode(`: heartbeat ${Date.now()}\n\n`));
    } catch {
      closed = true;
      clearInterval(heartbeatTimer);
    }
  }, HEARTBEAT_INTERVAL_MS);

  const close = () => {
    closed = true;
    clearInterval(heartbeatTimer);
    try { controller.close(); } catch { /* already closed */ }
  };
  return { stream, send, close };
}

export function sseHeaders() {
  return {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    'Connection': 'keep-alive',
    'X-Accel-Buffering': 'no',
  };
}
