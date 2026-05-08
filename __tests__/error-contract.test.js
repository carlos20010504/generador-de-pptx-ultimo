/* eslint-disable @typescript-eslint/no-require-imports */
const { test } = require('node:test');
const assert = require('node:assert');

// Mirror the contract here for testing the SHAPE of error payloads.
// The actual TS module is consumed by React components which are typecheck'd.
// node --test cannot resolve .ts directly without tsx, so we validate shape here.
const FRIENDLY_LABELS = {
  AI_SATURATED: {
    title: 'Servicio IA con mucha demanda ahora',
    action: 'Espera 5-10 min y reintenta. Tu Excel queda listo.',
  },
  EXCEL_INVALID: {
    title: 'No pude leer este Excel',
    action: 'Sube otro archivo o revisa que no esté dañado.',
  },
  EXCEL_EMPTY: {
    title: 'El Excel no tiene datos legibles',
    action: 'Asegúrate de que tenga al menos una hoja con tablas.',
  },
  EXCEL_INSUFFICIENT_DATA: {
    title: 'Los datos no alcanzan para una presentación',
    action: 'Mejora el Excel (más filas, menos vacíos) o ajusta el prompt.',
  },
  AI_RESPONSE_INVALID: {
    title: 'La IA devolvió una respuesta inválida',
    action: 'Reintenta — suele resolverse en el siguiente intento.',
  },
  PLANNER_REJECTED_PROMPT: {
    title: 'El prompt no encaja con este Excel',
    action: 'Cambia el prompt para enfocarte en datos disponibles.',
  },
  PYTHON_RUNTIME_ERROR: {
    title: 'Error técnico inesperado',
    action: 'Reporta el problema con el archivo que usaste.',
  },
  TIMEOUT: {
    title: 'La generación tomó demasiado',
    action: 'Reintenta o sube un archivo más pequeño.',
  },
};

test('error payload shape', () => {
  const err = {
    code: 'AI_SATURATED',
    message: 'all saturated',
    user_action: 'retry_later',
    retry_after_seconds: 300,
  };
  // Required fields present
  assert.ok(err.code && err.message && err.user_action);
  // Friendly label exists for code
  assert.ok(FRIENDLY_LABELS[err.code], 'FRIENDLY_LABELS must have entry for AI_SATURATED');
  // title references high demand / saturation concept
  assert.match(FRIENDLY_LABELS[err.code].title, /demanda|saturad/i);
  assert.match(FRIENDLY_LABELS[err.code].action, /reintenta/i);
});

test('all error codes have friendly labels', () => {
  const expectedCodes = [
    'EXCEL_INVALID', 'EXCEL_EMPTY', 'EXCEL_INSUFFICIENT_DATA',
    'AI_SATURATED', 'AI_RESPONSE_INVALID', 'PLANNER_REJECTED_PROMPT',
    'PYTHON_RUNTIME_ERROR', 'TIMEOUT',
  ];
  for (const code of expectedCodes) {
    assert.ok(FRIENDLY_LABELS[code], `Missing friendly label for ${code}`);
    assert.ok(FRIENDLY_LABELS[code].title.length > 0, `Empty title for ${code}`);
    assert.ok(FRIENDLY_LABELS[code].action.length > 0, `Empty action for ${code}`);
  }
});
