export type SSEEvent = {
  phase: 'parsing' | 'inventory' | 'planning' | 'validating' | 'rendering' | 'done' | 'error';
  step?: string;
  message?: string;
  data?: unknown;
};

export function makeSSEStream(): {
  stream: ReadableStream<Uint8Array>;
  send: (event: SSEEvent) => void;
  close: () => void;
} {
  const encoder = new TextEncoder();
  let controller: ReadableStreamDefaultController<Uint8Array>;
  const stream = new ReadableStream<Uint8Array>({
    start(c) { controller = c; },
  });
  const send = (event: SSEEvent) => {
    const payload = `data: ${JSON.stringify(event)}\n\n`;
    controller.enqueue(encoder.encode(payload));
  };
  const close = () => { try { controller.close(); } catch { /* ignore if already closed */ } };
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
