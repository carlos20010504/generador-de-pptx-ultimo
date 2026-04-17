import type { OrganizerMode } from '@/utils/excel-organizer';

const DOWNLOAD_TIMEOUT_MS = 2 * 60 * 1000;

function sanitizeDownloadName(fileName: string): string {
  return fileName.replace(/[<>:"/\\|?*\x00-\x1F]+/g, '_').trim() || `Reporte_Socya_${Date.now()}.pptx`;
}

function getErrorMessage(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  return fallback;
}

export async function generatePowerPointFromExcel(file: File, visualMode: OrganizerMode = 'mixed'): Promise<void> {
  try {
    console.log('[1/3] 📤 Enviando archivo al backend premium...');

    const formData = new FormData();
    formData.append('file', file);
    formData.append('visualMode', visualMode === 'auto' ? 'mixed' : visualMode);

    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), DOWNLOAD_TIMEOUT_MS);

    // Call the premium PPTX API that uses the same spectacular generator validated offline.
    const response = await fetch('/api/generate-pptx', {
      method: 'POST',
      body: formData,
      signal: controller.signal,
    }).finally(() => {
      window.clearTimeout(timeoutId);
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => null);
      throw new Error(errorData?.error || 'Error al generar el PowerPoint en el servidor');
    }

    console.log('[2/3] 🖼 Recibiendo presentación final validada...');
    const blob = await response.blob();
    if (!blob.size) {
      throw new Error('El servidor devolvió un archivo vacío.');
    }

    const disposition = response.headers.get('content-disposition') || '';
    const match = disposition.match(/filename="([^"]+)"/i);
    const downloadName = sanitizeDownloadName(match?.[1] || `Reporte_Socya_${Date.now()}.pptx`);

    console.log('[3/3] 💾 Descargando presentación...');
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

  } catch (err: unknown) {
    console.error('[PPTX Helper] Error:', err);
    if (err instanceof DOMException && err.name === 'AbortError') {
      throw new Error('La generación tardó demasiado tiempo. Intenta con un archivo más pequeño o simplificado.');
    }
    throw new Error(getErrorMessage(err, 'Error al generar la presentación.'));
  }
}
