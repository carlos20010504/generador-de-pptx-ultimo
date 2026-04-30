import * as XLSX from 'xlsx';

export interface ExcelLikeFile {
  name: string;
  size: number;
  type?: string;
}

export const MAX_EXCEL_UPLOAD_BYTES = 60 * 1024 * 1024;
export const CLIENT_DEEP_VALIDATION_THRESHOLD_BYTES = 12 * 1024 * 1024;
export const ACCEPTED_EXCEL_EXTENSIONS = ['.xlsx', '.xls', '.xlsm'] as const;
export const ACCEPTED_EXCEL_MIME_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
  'application/vnd.ms-excel.sheet.macroenabled.12',
  'application/octet-stream',
]);

export function getExcelExtension(fileName: string): string {
  const lastDot = fileName.lastIndexOf('.');
  return lastDot >= 0 ? fileName.slice(lastDot).toLowerCase() : '';
}

export function getMaxExcelUploadSizeMb(): number {
  return Math.round(MAX_EXCEL_UPLOAD_BYTES / (1024 * 1024));
}

export function shouldSkipClientContentValidation(file: ExcelLikeFile): boolean {
  return file.size >= CLIENT_DEEP_VALIDATION_THRESHOLD_BYTES;
}

export function validateExcelUpload(file: ExcelLikeFile): string | null {
  const ext = getExcelExtension(file.name);
  const mime = String(file.type || '').toLowerCase();

  if (!ACCEPTED_EXCEL_EXTENSIONS.includes(ext as typeof ACCEPTED_EXCEL_EXTENSIONS[number])) {
    return 'El archivo debe ser un Excel valido (.xlsx, .xls o .xlsm).';
  }

  if (mime && !ACCEPTED_EXCEL_MIME_TYPES.has(mime)) {
    return 'El archivo subido no tiene un tipo MIME de Excel permitido.';
  }

  if (file.size <= 0) {
    return 'El archivo Excel esta vacio.';
  }

  if (file.size > MAX_EXCEL_UPLOAD_BYTES) {
    return `El archivo excede el limite permitido de ${getMaxExcelUploadSizeMb()} MB.`;
  }

  return null;
}

function isMeaningfulCell(value: unknown): boolean {
  if (value === null || value === undefined) return false;
  const text = String(value).trim();
  if (!text) return false;
  return !['-', '--', 'n/a', 'na', 'null', 'undefined'].includes(text.toLowerCase());
}

function countMeaningfulCells(row: unknown[]): number {
  return row.reduce<number>((count, cell) => count + (isMeaningfulCell(cell) ? 1 : 0), 0);
}

function looksLikeHeaderCell(value: unknown): boolean {
  if (!isMeaningfulCell(value)) return false;
  const text = String(value).trim();
  const numeric = Number(String(text).replace(/[$,%\s,]/g, ''));
  if (!Number.isNaN(numeric) && text.length < 8) return false;
  return /[a-zA-ZáéíóúÁÉÍÓÚñÑ]/.test(text);
}

function findHeaderRow(rows: unknown[][]): number {
  let bestIndex = 0;
  let bestScore = -1;

  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const row = rows[i] || [];
    const meaningful = countMeaningfulCells(row);
    const headerLike = row.reduce<number>((count, cell) => count + (looksLikeHeaderCell(cell) ? 1 : 0), 0);
    const score = headerLike * 3 + meaningful;
    if (score > bestScore && headerLike >= 2) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex;
}

function sanitizeHeaders(row: unknown[]): string[] {
  return row
    .map((cell) => String(cell ?? '').trim())
    .map((text, index) => (looksLikeHeaderCell(text) ? text : `Columna ${index + 1}`));
}

function isDenseDataRow(row: unknown[], columnCount: number): boolean {
  if (!row?.length || columnCount <= 0) return false;
  const meaningful = countMeaningfulCells(row.slice(0, columnCount));
  return meaningful >= Math.max(2, Math.ceil(columnCount * 0.55));
}

export async function validateExcelContents(file: File): Promise<string | null> {
  if (shouldSkipClientContentValidation(file)) {
    return null;
  }

  try {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    if (!wb.SheetNames.length) return 'El archivo Excel no tiene hojas.';

    let foundUsableSheet = false;

    // Check the first few sheets
    for (const sheetName of wb.SheetNames.slice(0, 3)) {
      const ws = wb.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as unknown[][];
      const rows = data.filter((row) => Array.isArray(row) && countMeaningfulCells(row) > 0);

      if (rows.length < 2) continue;

      const headerIndex = findHeaderRow(rows);
      const headers = sanitizeHeaders(rows[headerIndex] || []);
      const bodyRows = rows
        .slice(headerIndex + 1)
        .filter((row) => isDenseDataRow(row, headers.length))
        .slice(0, 100);

      if (headers.length < 2 || bodyRows.length === 0) continue;
      foundUsableSheet = true;

      const numericColumns: number[] = [];
      const headerNames: string[] = [];

      for (let i = 0; i < headers.length; i++) {
        const header = String(headers[i] || '').toLowerCase();
        const isNumericKeyword = ['costo', 'valor', 'presupuesto', 'monto', 'total', 'precio', 'saldo', 'cantidad'].some(k => header.includes(k));
        const isExcludedKeyword = ['centro', 'id', 'codigo', 'código', 'cuenta'].some(k => header.includes(k));

        if (isNumericKeyword && !isExcludedKeyword) {
          numericColumns.push(i);
          headerNames[i] = String(headers[i]);
        }
      }

      if (numericColumns.length > 0) {
        for (let rowIdx = 0; rowIdx < bodyRows.length; rowIdx++) {
          const row = bodyRows[rowIdx];
          if (!row || !row.length) continue;

          for (const colIdx of numericColumns) {
            const val = row[colIdx];
            if (isMeaningfulCell(val)) {
              if (typeof val === 'string') {
                const cleaned = val.replace(/[$,\s%]/g, '');
                if (Number.isNaN(Number(cleaned))) {
                  const excelRowNumber = headerIndex + rowIdx + 2;
                  return `Error de validacion en la hoja '${sheetName}', fila ${excelRowNumber}: la columna '${headerNames[colIdx]}' debe contener valores numericos, pero se encontro texto ("${val}"). Por favor corrige el Excel y vuelve a intentar.`;
                }
              }
            }
          }
        }
      }
    }

    if (!foundUsableSheet) {
      return 'El Excel no contiene una hoja con encabezados y filas de datos consistentes para generar una presentacion confiable.';
    }
  } catch (error) {
    console.error('Error reading excel in validateExcelContents:', error);
    return 'Error interno al leer el contenido del Excel para validarlo.';
  }
  
  return null;
}
