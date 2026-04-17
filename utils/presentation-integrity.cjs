const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

function normalizePath(filePath) {
  return path.resolve(filePath).toLowerCase();
}

function sanitizeBaseName(filePath) {
  return path.basename(filePath, path.extname(filePath)).replace(/[^a-zA-Z0-9_-]+/g, '_');
}

function createProcessingContext(cliInput, rootDir) {
  if (!cliInput || !String(cliInput).trim()) {
    throw new Error(
      'No se proporcionó un archivo Excel de entrada. El modo seguro bloquea cualquier fallback automático y exige un archivo explícito por ejecución.'
    );
  }

  const inputFile = path.resolve(cliInput);
  const safeBaseName = sanitizeBaseName(inputFile);
  const outputFile = path.join(path.dirname(inputFile), `Presentacion_Ejecutiva_Socya_${safeBaseName}.pptx`);
  const auditLogFile = path.join(path.dirname(inputFile), `Presentacion_Ejecutiva_Socya_${safeBaseName}.audit.json`);
  const allowedSupportFiles = [
    path.join(rootDir, 'organizer.py'),
    path.join(rootDir, 'generate_excel_org_presentation.js'),
  ];

  return {
    rootDir,
    inputFile,
    inputLabel: path.basename(inputFile),
    outputFile,
    auditLogFile,
    allowedSupportFiles,
  };
}

function ensureExplicitInputDependencies(context) {
  if (normalizePath(context.inputFile).endsWith(`${path.sep}excel org.xlsx`)) {
    throw new Error(
      'Dependencia no autorizada detectada: se intentó usar el archivo legado `excel org.xlsx`. Cada generación debe usar únicamente el Excel suministrado en la solicitud actual.'
    );
  }
}

function computeFileHash(filePath) {
  const content = fs.readFileSync(filePath);
  return crypto.createHash('sha256').update(content).digest('hex');
}

function validateOrganizerConsistency(sourceData, inputFile) {
  const fileLabel = path.basename(inputFile);
  const organizerFile = String(sourceData?.metadatos?.archivo || '').trim();

  if (!sourceData || typeof sourceData !== 'object') {
    throw new Error('El organizador no devolvió una estructura válida para el Excel suministrado.');
  }

  if (sourceData.error) {
    throw new Error(`El organizador reportó un error: ${String(sourceData.error)}`);
  }

  if (organizerFile && organizerFile !== fileLabel) {
    throw new Error(
      `Dependencia no autorizada detectada: el organizador devolvió datos para \`${organizerFile}\`, pero la ejecución actual corresponde a \`${fileLabel}\`.`
    );
  }

  const detectedSheets = Array.isArray(sourceData?.metadatos?.hojas_encontradas) ? sourceData.metadatos.hojas_encontradas.length : 0;
  if (detectedSheets === 0) {
    throw new Error('No se detectaron hojas utilizables en el Excel suministrado. El modo seguro detiene la generación para evitar una presentación infiel.');
  }
}

function countOrganizerTables(sourceData) {
  const counts = [
    sourceData?.muestra_tabla ? 1 : 0,
    Object.keys(sourceData?.otras_tablas || {}).length,
    Object.keys(sourceData?.genericas || {}).length,
    sourceData?.coso ? 1 : 0,
    sourceData?.distribucion_mes ? 1 : 0,
  ];
  return counts.reduce((acc, value) => acc + value, 0);
}

function decideSafeMode(options) {
  const {
    requestedVisualMode = 'mixed',
    sourceData = null,
    specializedCommissions = false,
  } = options || {};

  const reasons = [];
  const organizerTableCount = countOrganizerTables(sourceData || {});
  const hasGenericSummary = !!sourceData?.resumen_generico;

  if (!sourceData) reasons.push('No fue posible leer la organización estructurada del Excel.');
  if (!specializedCommissions) reasons.push('El archivo no coincide con el flujo especializado de comisiones.');
  if (hasGenericSummary) reasons.push('La narrativa debe ser conservadora y pegada a la estructura organizada del archivo.');
  if (requestedVisualMode !== 'mixed') reasons.push(`El usuario solicitó una salida controlada en modo ${requestedVisualMode}.`);
  if (organizerTableCount <= 2) reasons.push('La estructura detectada es limitada y requiere una generación literal.');

  return {
    enabled: reasons.length > 0,
    reasons,
    organizerTableCount,
  };
}

function assertNoUnauthorizedDataDependencies(context, referencedPaths) {
  const allowed = new Set([
    normalizePath(context.inputFile),
    ...context.allowedSupportFiles.map(normalizePath),
  ]);

  for (const filePath of referencedPaths || []) {
    if (!filePath) continue;
    const normalized = normalizePath(filePath);
    if (!allowed.has(normalized)) {
      throw new Error(
        `Dependencia no autorizada detectada durante la generación: \`${filePath}\`. El proceso solo puede usar el Excel solicitado y los scripts internos autorizados.`
      );
    }
  }
}

function buildAuditRecord(options) {
  const {
    context,
    presentationMode,
    safeMode,
    sourceData,
    outputFile,
    warnings = [],
  } = options;

  return {
    generatedAt: new Date().toISOString(),
    inputFile: context.inputFile,
    inputLabel: context.inputLabel,
    inputSha256: computeFileHash(context.inputFile),
    outputFile,
    outputExists: fs.existsSync(outputFile),
    outputSizeBytes: fs.existsSync(outputFile) ? fs.statSync(outputFile).size : 0,
    presentationMode,
    safeMode: {
      enabled: !!safeMode?.enabled,
      reasons: safeMode?.reasons || [],
      organizerTableCount: safeMode?.organizerTableCount || 0,
    },
    sourceSummary: {
      organizerFile: String(sourceData?.metadatos?.archivo || ''),
      detectedSheets: Array.isArray(sourceData?.metadatos?.hojas_encontradas) ? sourceData.metadatos.hojas_encontradas : [],
      hasSpecializedCommissionsShape: !!sourceData?.resumen_ejecutivo,
      hasGenericSummary: !!sourceData?.resumen_generico,
      hasMainTable: !!sourceData?.muestra_tabla,
      otherTables: Object.keys(sourceData?.otras_tablas || {}),
      genericTables: Object.keys(sourceData?.genericas || {}),
    },
    warnings,
  };
}

function writeAuditRecord(auditLogFile, record) {
  fs.writeFileSync(auditLogFile, JSON.stringify(record, null, 2), 'utf8');
}

module.exports = {
  assertNoUnauthorizedDataDependencies,
  buildAuditRecord,
  createProcessingContext,
  decideSafeMode,
  ensureExplicitInputDependencies,
  validateOrganizerConsistency,
  writeAuditRecord,
};
