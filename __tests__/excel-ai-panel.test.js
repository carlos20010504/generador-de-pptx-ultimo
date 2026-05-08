/* eslint-disable @typescript-eslint/no-require-imports */
const test = require('node:test');
const assert = require('node:assert/strict');

const {
  buildProcessingProfile,
  constants,
} = require('../utils/excel-ai-panel.cjs');

test('activa perfil extendido para archivos grandes y prompts largos', () => {
  const standard = buildProcessingProfile({ fileSizeBytes: 2 * 1024 * 1024, userPrompt: '' });
  const extended = buildProcessingProfile({
    fileSizeBytes: constants.SHALLOW_VALIDATION_THRESHOLD_BYTES,
    userPrompt: 'Quiero una narrativa con tendencias, comparativos y recomendaciones ejecutivas.',
  });
  const highVolume = buildProcessingProfile({
    fileSizeBytes: constants.EXTENDED_PROCESSING_THRESHOLD_BYTES + 1024,
    userPrompt: 'x'.repeat(900),
  });

  assert.equal(standard.tier, 'estandar');
  assert.equal(standard.skipClientDeepValidation, false);
  assert.equal(extended.tier, 'extendido');
  assert.equal(extended.skipClientDeepValidation, true);
  assert.equal(highVolume.tier, 'alto-volumen');
  assert.equal(highVolume.timeoutMs, constants.MAX_TIMEOUT_MS);
});
