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
