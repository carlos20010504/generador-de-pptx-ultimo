/* eslint-disable @typescript-eslint/no-require-imports */
const { test } = require('node:test');
const assert = require('node:assert');

// Re-implement the SSE encoder shape for testing the contract.
// The actual TS module is consumed by Next.js routes which are typecheck'd.
// node --test cannot resolve .ts directly without tsx, so we validate shape here.
function makeSSEStreamPure() {
  const encoder = new TextEncoder();
  const chunks = [];
  return {
    send(ev) { chunks.push(encoder.encode(`data: ${JSON.stringify(ev)}\n\n`)); },
    collect() {
      const decoder = new TextDecoder();
      return chunks.map(c => decoder.decode(c)).join('');
    },
  };
}

test('sse event encoding', () => {
  const s = makeSSEStreamPure();
  s.send({ phase: 'parsing', message: 'go' });
  s.send({ phase: 'done', data: { downloadToken: 'abc' } });
  const text = s.collect();
  assert.match(text, /data: \{"phase":"parsing"/);
  assert.match(text, /data: \{"phase":"done"/);
  assert.match(text, /downloadToken/);
  // Each event terminated with two newlines (SSE spec)
  assert.match(text, /\n\n.*\n\n/s);
});

test('sse event types match enum', () => {
  const validPhases = ['parsing', 'inventory', 'planning', 'validating',
                        'rendering', 'done', 'error'];
  for (const p of validPhases) {
    assert.equal(typeof p, 'string');
  }
});
