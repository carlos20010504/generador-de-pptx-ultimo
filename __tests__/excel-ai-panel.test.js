/* eslint-disable @typescript-eslint/no-require-imports */
const test = require('node:test');
const assert = require('node:assert/strict');

const {
  buildExcelIntelligenceReport,
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

test('normaliza un reporte rico de IA para el panel derecho', () => {
  const report = buildExcelIntelligenceReport({
    fileName: 'auditoria.xlsx',
    fileSizeBytes: 24 * 1024 * 1024,
    userPrompt: 'Usa tonos sobrios y enfatiza riesgos.',
    suggestions: [
      'Abre con un semáforo de riesgos y una conclusión ejecutiva.',
      'Incluye una gráfica comparativa de hallazgos por categoría.',
      'Cierra con recomendaciones priorizadas por impacto.',
    ],
    analysis: {
      metadatos: {
        archivo: 'auditoria.xlsx',
        hojas_encontradas: ['Resumen', 'Hallazgos', 'COSO'],
        tipo_libro: 'auditoria_control',
        familias_detectadas: ['auditoria', 'hallazgos', 'coso'],
      },
      resumen_generico: {
        hoja_principal: 'Resumen',
        total_filas: 18540,
        total_columnas: 24,
      },
      resumen_ejecutivo_ia: {
        vision_general: 'El archivo evidencia focos de control interno con concentración en pocos frentes críticos.',
        alerta_principal: 'La concentración de hallazgos en un grupo reducido exige priorización inmediata.',
        recomendacion: 'Ordena las acciones por impacto y madurez del control.',
      },
      briefing_ejecutivo_ia: {
        de_que_trata: 'Revisión ejecutiva de riesgos, hallazgos y acciones de mejora.',
        breve_resumen: ['El comportamiento del libro favorece una narrativa de comité y no una exposición tabular extensa.'],
        objetivos: ['Visualizar riesgos críticos y su plan de respuesta.'],
      },
      conclusiones: [
        'Se detectan 18.540 registros utilizables y múltiples hojas con enfoque de auditoría.',
        'La información requiere priorizar hallazgos, no replicar tablas completas.',
      ],
      graficas_automaticas: [
        { titulo: 'Hallazgos por categoría' },
        { titulo: 'Estado de acciones correctivas' },
      ],
      kpis_automaticos: [
        { label: 'Hallazgos críticos', value: '12' },
        { label: 'Áreas impactadas', value: '5' },
      ],
      muestra_tabla: {
        hoja_origen: 'Resumen',
      },
      otras_tablas: {
        Hallazgos: {},
      },
      bloques_textuales: [{ title: 'Hallazgos relevantes' }],
      analisis_avanzado: {
        tendencia: {
          col_temporal: 'Mes',
          col_valor: 'Hallazgos',
          tendencia: 'ascendente',
        },
        pareto: [
          {
            columna: 'Proceso',
            top_valores: [{ valor: 'Tesorería', pct: 68 }],
          },
        ],
        insights: [
          { texto: 'El patrón temporal muestra crecimiento sostenido de hallazgos abiertos.' },
        ],
      },
    },
  });

  assert.equal(report.dataset.workbookType, 'auditoria_control');
  assert.equal(report.dataset.sheetCount, 3);
  assert.equal(report.processing.tier, 'alto-volumen');
  assert.equal(report.powerPointPlan.palette.name, 'Comité Ejecutivo');
  assert.ok(report.executiveSummary.includes('Revisión ejecutiva'));
  assert.ok(report.trends.some((item) => /tendencia ascendente/i.test(item)));
  assert.ok(report.patterns.some((item) => /Tesorería/i.test(item)));
  assert.ok(report.powerPointPlan.recommendedSlides.length >= 4);
  assert.equal(report.suggestions.length, 3);
});

test('tolera payloads incompletos o malformados sin romper el panel', () => {
  const report = buildExcelIntelligenceReport({
    analysis: {
      metadatos: {
        archivo: 'malformado.xlsx',
        hojas_encontradas: [],
      },
      conclusiones: [null, '', 'Solo se detectó una tabla parcial.'],
    },
    suggestions: [null, ''],
  });

  assert.equal(report.dataset.fileName, 'malformado.xlsx');
  assert.equal(report.dataset.sheetCount, 0);
  assert.equal(report.suggestions.length, 0);
  assert.ok(Array.isArray(report.keyFindings));
  assert.ok(report.executiveSummary.length > 0);
});
