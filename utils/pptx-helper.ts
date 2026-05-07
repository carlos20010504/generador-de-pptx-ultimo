import type { OrganizerMode } from '@/utils/excel-organizer';

const DOWNLOAD_TIMEOUT_MS = 30 * 60 * 1000;

function sanitizeDownloadName(fileName: string): string {
  return fileName.replace(/[<>:"/\\|?*\x00-\x1F]+/g, '_').trim() || `Reporte_Socya_${Date.now()}.pptx`;
}

function getErrorMessage(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message) {
    return error.message;
  }

  return fallback;
}

export async function generatePowerPointFromExcel(
  file: File, 
  visualMode: OrganizerMode = 'mixed',
  userPrompt: string = '',
  presentationContext?: {
    audience: string;
    language: string;
    theme: {
      key: string;
      name: string;
      primary_hex: string;
      accent_hex: string;
      text_hex: string;
      bg_hex: string;
    };
  }
): Promise<void> {
  try {
    console.log('[1/3] 📤 Enviando archivo al backend premium...');

    const formData = new FormData();
    formData.append('file', file);
    formData.append('visualMode', visualMode === 'auto' ? 'mixed' : visualMode);
    if (userPrompt) {
      formData.append('userPrompt', userPrompt);
    }
    if (presentationContext) {
      formData.append('audience', presentationContext.audience);
      formData.append('language', presentationContext.language);
      formData.append('theme', JSON.stringify(presentationContext.theme));
    }

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
      throw new Error('La generación tardó demasiado tiempo. Intenta de nuevo o usa un archivo más pequeño si el Excel es especialmente pesado.');
    }
    throw new Error(getErrorMessage(err, 'Error al generar la presentación.'));
  }
}
