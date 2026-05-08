export interface AIStatus {
  model?: string;
  cache_hit?: boolean;
  fallback_steps?: { from: string; reason: string; message?: string }[];
}

const MODEL_LABELS: Record<string, string> = {
  'nousresearch/hermes-3-llama-3.1-405b:free': 'Hermes 3 405B',
  'meta-llama/llama-3.3-70b-instruct:free': 'Llama 3.3 70B',
  'mistralai/mistral-small-3.1-24b-instruct:free': 'Mistral Small 24B',
  'qwen/qwen-2.5-72b-instruct:free': 'Qwen 2.5 72B',
};

export function aiStatusBadge(status?: AIStatus): {
  label: string; tone: 'ok' | 'warn' | 'cache';
} {
  if (!status) return { label: 'Esperando IA…', tone: 'warn' };
  if (status.cache_hit) return {
    label: 'Resultado en caché (sin consumo IA)', tone: 'cache',
  };
  const label = MODEL_LABELS[status.model || ''] || status.model || 'IA';
  if ((status.fallback_steps?.length || 0) > 0) {
    const skipped = status.fallback_steps?.[0]?.from || 'modelo primario';
    const skippedLabel = MODEL_LABELS[skipped] || skipped;
    return { label: `${skippedLabel} saturado → usando ${label}`, tone: 'warn' };
  }
  return { label: `${label} · listo`, tone: 'ok' };
}
