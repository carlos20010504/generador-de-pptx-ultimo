const STANDARD_TIMEOUT_MS = 8 * 60 * 1000;
const EXTENDED_TIMEOUT_MS = 16 * 60 * 1000;
const MAX_TIMEOUT_MS = 24 * 60 * 1000;
const SHALLOW_VALIDATION_THRESHOLD_BYTES = 10 * 1024 * 1024;
const EXTENDED_PROCESSING_THRESHOLD_BYTES = 16 * 1024 * 1024;

function buildProcessingProfile({ fileSizeBytes = 0, userPrompt = '' } = {}) {
  const promptLength = String(userPrompt || '').trim().length;
  let timeoutMs = STANDARD_TIMEOUT_MS;
  let tier = 'estandar';

  if (fileSizeBytes >= SHALLOW_VALIDATION_THRESHOLD_BYTES || promptLength >= 180) {
    timeoutMs = EXTENDED_TIMEOUT_MS;
    tier = 'extendido';
  }

  if (fileSizeBytes >= EXTENDED_PROCESSING_THRESHOLD_BYTES || promptLength >= 420) {
    timeoutMs = MAX_TIMEOUT_MS;
    tier = 'alto-volumen';
  }

  return {
    tier,
    timeoutMs,
    timeoutMinutes: Math.round(timeoutMs / 60000),
    skipClientDeepValidation: fileSizeBytes >= SHALLOW_VALIDATION_THRESHOLD_BYTES,
    allowsExtendedProcessing: timeoutMs > STANDARD_TIMEOUT_MS,
  };
}

module.exports = {
  buildProcessingProfile,
  constants: {
    STANDARD_TIMEOUT_MS,
    EXTENDED_TIMEOUT_MS,
    MAX_TIMEOUT_MS,
    SHALLOW_VALIDATION_THRESHOLD_BYTES,
    EXTENDED_PROCESSING_THRESHOLD_BYTES,
  },
};
