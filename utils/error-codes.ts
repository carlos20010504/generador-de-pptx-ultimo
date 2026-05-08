export type ErrorCode =
  | 'EXCEL_INVALID' | 'EXCEL_EMPTY' | 'EXCEL_INSUFFICIENT_DATA'
  | 'AI_SATURATED' | 'AI_RESPONSE_INVALID' | 'PLANNER_REJECTED_PROMPT'
  | 'PYTHON_RUNTIME_ERROR' | 'TIMEOUT';

export type UserAction =
  | 'retry' | 'retry_later' | 'upload_again' | 'change_prompt'
  | 'improve_excel_or_change_prompt' | 'upload_smaller' | 'report_bug';

export interface PipelineErrorPayload {
  code: ErrorCode;
  message: string;
  details?: string;
  user_action: UserAction;
  retry_after_seconds?: number;
}

export function isPipelineError(x: unknown): x is PipelineErrorPayload {
  return Boolean(x && typeof x === 'object' && 'code' in x && 'message' in x
                  && 'user_action' in x);
}

const FRIENDLY: Record<ErrorCode, { title: string; action: string }> = {
  EXCEL_INVALID: {
    title: 'No pude leer este Excel',
    action: 'Sube otro archivo o revisa que no esté dañado.',
  },
  EXCEL_EMPTY: {
    title: 'El Excel no tiene datos legibles',
    action: 'Asegúrate de que tenga al menos una hoja con tablas.',
  },
  EXCEL_INSUFFICIENT_DATA: {
    title: 'Los datos no alcanzan para una presentación',
    action: 'Mejora el Excel (más filas, menos vacíos) o ajusta el prompt.',
  },
  AI_SATURATED: {
    title: 'Servicio IA con mucha demanda ahora',
    action: 'Espera 5-10 min y reintenta. Tu Excel queda listo.',
  },
  AI_RESPONSE_INVALID: {
    title: 'La IA devolvió una respuesta inválida',
    action: 'Reintenta — suele resolverse en el siguiente intento.',
  },
  PLANNER_REJECTED_PROMPT: {
    title: 'El prompt no encaja con este Excel',
    action: 'Cambia el prompt para enfocarte en datos disponibles.',
  },
  PYTHON_RUNTIME_ERROR: {
    title: 'Error técnico inesperado',
    action: 'Reporta el problema con el archivo que usaste.',
  },
  TIMEOUT: {
    title: 'La generación tomó demasiado',
    action: 'Reintenta o sube un archivo más pequeño.',
  },
};

export function formatErrorForUser(err: PipelineErrorPayload) {
  return FRIENDLY[err.code] ?? {
    title: 'Algo salió mal',
    action: err.message || 'Reintenta más tarde.',
  };
}
