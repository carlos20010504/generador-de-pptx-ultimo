export interface PptxLikeFile {
  name: string;
  size: number;
  type?: string;
}

export const MAX_PPTX_UPLOAD_BYTES = 25 * 1024 * 1024;
export const ACCEPTED_PPTX_EXTENSIONS = ['.pptx'] as const;
export const ACCEPTED_PPTX_MIME_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  'application/octet-stream',
  'application/zip',
]);

export function getPptxExtension(fileName: string): string {
  const lastDot = fileName.lastIndexOf('.');
  return lastDot >= 0 ? fileName.slice(lastDot).toLowerCase() : '';
}

export function getMaxPptxUploadSizeMb(): number {
  return Math.round(MAX_PPTX_UPLOAD_BYTES / (1024 * 1024));
}

export function validatePptxUpload(file: PptxLikeFile): string | null {
  const ext = getPptxExtension(file.name);
  const mime = String(file.type || '').toLowerCase();

  if (!ACCEPTED_PPTX_EXTENSIONS.includes(ext as typeof ACCEPTED_PPTX_EXTENSIONS[number])) {
    return 'El archivo debe ser una presentación PowerPoint válida (.pptx).';
  }

  if (mime && !ACCEPTED_PPTX_MIME_TYPES.has(mime)) {
    return 'El archivo subido no tiene un tipo MIME de PowerPoint permitido.';
  }

  if (file.size <= 0) {
    return 'El archivo PowerPoint está vacío.';
  }

  if (file.size > MAX_PPTX_UPLOAD_BYTES) {
    return `El archivo excede el límite permitido de ${getMaxPptxUploadSizeMb()} MB.`;
  }

  return null;
}

export async function validatePptxContents(file: File): Promise<string | null> {
  try {
    const buffer = await file.arrayBuffer();
    if (buffer.byteLength < 4) {
      return 'El archivo PowerPoint está incompleto.';
    }

    const bytes = new Uint8Array(buffer.slice(0, 4));
    const isZip =
      bytes[0] === 0x50 &&
      bytes[1] === 0x4b &&
      (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07) &&
      (bytes[3] === 0x04 || bytes[3] === 0x06 || bytes[3] === 0x08);

    if (!isZip) {
      return 'El archivo no parece ser un PowerPoint Open XML válido.';
    }
  } catch (error) {
    console.error('Error validating pptx contents:', error);
    return 'Error interno al leer el contenido del PowerPoint para validarlo.';
  }

  return null;
}
