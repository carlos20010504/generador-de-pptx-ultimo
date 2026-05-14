// lib/model-presets.ts
export interface ModelPreset {
  id: string;
  label: string;
  sub: string;
  provider: string;
  model: string;
}

export const MODEL_PRESETS: ModelPreset[] = [
  { id: "fast",         label: "🚀 Rápido",       sub: "Cerebras GPT-OSS 120B",      provider: "cerebras", model: "gpt-oss-120b" },
  { id: "experimental", label: "🧪 Experimental", sub: "Cerebras GLM 4.7 (preview)", provider: "cerebras", model: "zai-glm-4.7" },
  { id: "economico",    label: "⚖️ Económico",     sub: "Cerebras Llama 3.1 8B",      provider: "cerebras", model: "llama3.1-8b" },
  { id: "backup",       label: "🎯 Backup Groq",  sub: "Llama 3.3 70B",               provider: "groq",     model: "llama-3.3-70b-versatile" },
];

export const DEFAULT_PRESET_ID = "fast";
export const PREFERRED_MODEL_LS_KEY = "socya:preferred_model_id";

export function presetToModelString(presetId: string): string | null {
  const p = MODEL_PRESETS.find(x => x.id === presetId);
  return p ? `${p.provider}/${p.model}` : null;
}
