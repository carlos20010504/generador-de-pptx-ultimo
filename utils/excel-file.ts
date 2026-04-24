import * as XLSX from 'xlsx';

export interface ExcelLikeFile {
  name: string;
  size: number;
  type?: string;
}

export const MAX_EXCEL_UPLOAD_BYTES = 20 * 1024 * 1024;
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

export async function validateExcelContents(file: File): Promise<string | null> {
  try {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    if (!wb.SheetNames.length) return 'El archivo Excel no tiene hojas.';

    // Check the first few sheets
    for (const sheetName of wb.SheetNames.slice(0, 3)) {
      const ws = wb.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
      
      if (data.length < 2) continue;
      
      // Look for a header row (first row with enough strings)
      const headers = data[0] || [];
      const numericColumns: number[] = [];
      const headerNames: string[] = [];

      for (let i = 0; i < headers.length; i++) {
        const header = String(headers[i] || '').toLowerCase();
        if (['costo', 'valor', 'presupuesto', 'monto', 'total', 'precio', 'saldo', 'cantidad'].some(k => header.includes(k))) {
          numericColumns.push(i);
          headerNames[i] = String(headers[i]);
        }
      }

      if (numericColumns.length > 0) {
        for (let rowIdx = 1; rowIdx < Math.min(data.length, 100); rowIdx++) {
          const row = data[rowIdx];
          if (!row || !row.length) continue;
          
          for (const colIdx of numericColumns) {
            const val = row[colIdx];
            if (val !== undefined && val !== null && val !== '') {
              // Si es string, verifiquemos si se puede convertir a numero
              if (typeof val === 'string') {
                const cleaned = val.replace(/[\$,\s]/g, '');
                if (isNaN(Number(cleaned)) && cleaned.toLowerCase() !== 'na' && cleaned !== '-') {
                  return `Error de validación en la hoja '${sheetName}', Fila ${rowIdx + 1}: La columna '${headerNames[colIdx]}' debe contener valores numéricos, pero se encontró texto ("${val}"). Por favor corrige el Excel y vuelve a intentar.`;
                }
              }
            }
          }
        }
      }
    }
  } catch (error) {
    console.error("Error reading excel in validateExcelContents:", error);
    return 'Error interno al leer el contenido del Excel para validarlo.';
  }
  
  return null;
}
