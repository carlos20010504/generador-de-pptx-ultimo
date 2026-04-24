const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

const {
  assertNoUnauthorizedDataDependencies,
  createProcessingContext,
  decideSafeMode,
  ensureExplicitInputDependencies,
  validateOrganizerConsistency,
} = require('../utils/presentation-integrity.cjs');

const ROOT_DIR = path.resolve(__dirname, '..');

test('bloquea el fallback implicito cuando no se envía archivo de entrada', () => {
  assert.throws(
    () => createProcessingContext('', ROOT_DIR),
    (error) => /fallback/i.test(String(error?.message || ''))
  );
});

test('rechaza el uso del excel legado como dependencia de entrada', () => {
  const context = createProcessingContext(path.join(ROOT_DIR, 'excel org.xlsx'), ROOT_DIR);
  assert.throws(
    () => ensureExplicitInputDependencies(context),
    /excel org\.xlsx/i
  );
});

test('valida que el organizador responda para el mismo archivo suministrado', () => {
  const input = path.join(ROOT_DIR, 'entrada-prueba.xlsx');
  assert.throws(
    () =>
      validateOrganizerConsistency(
        {
          metadatos: {
            archivo: 'otro_archivo.xlsx',
            hojas_encontradas: ['Hoja1'],
          },
        },
        input
      ),
    /dependencia no autorizada detectada/i
  );
});

test('activa modo seguro para datasets no especializados y con estructura limitada', () => {
  const safeMode = decideSafeMode({
    requestedVisualMode: 'mixed',
    specializedCommissions: false,
    sourceData: {
      metadatos: { hojas_encontradas: ['PLAN'] },
      resumen_generico: { hoja_principal: 'PLAN' },
      genericas: {},
      otras_tablas: {},
    },
  });

  assert.equal(safeMode.enabled, true);
  assert.match(safeMode.reasons.join(' '), /flujo especializado de comisiones/i);
});

test('rechaza rutas de datos no autorizadas distintas al archivo actual', () => {
  const context = createProcessingContext(path.join(ROOT_DIR, 'entrada-prueba.xlsx'), ROOT_DIR);

  assert.throws(
    () =>
      assertNoUnauthorizedDataDependencies(context, [
        context.inputFile,
        path.join(ROOT_DIR, 'organizer.py'),
        path.join(ROOT_DIR, 'origen-no-autorizado.xlsx'),
      ]),
    /dependencia no autorizada detectada durante la generación/i
  );
});
