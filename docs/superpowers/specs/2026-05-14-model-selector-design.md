# Model Selector + Cerebras Primary — Design

**Date:** 2026-05-14
**Status:** Approved (verbal)
**Branch base:** `feature/hermes-driven-pipeline`

## Problema

Usuario quiere:
1. Cerebras como provider primario para el plan (key paga, ~2000 tok/s).
2. Selector en la UI para elegir entre 4 modelos curados.
3. Quitar Gemini (no lo usa, simplifica el código).

## Decisiones

| Decisión | Opción | Por qué |
|---|---|---|
| Granularidad selector | **Curado** (4 picks con label-promesa) | User no sabe model IDs, sí entiende "rápido vs balanceado" |
| Provider quitado | **Gemini** | API distinta (no openai-compat), 15 RPM, no se usa |
| Default | **Cerebras GPT-OSS 120B** | Paga, súper rápido, calidad alta |
| Persistencia | **localStorage** | Selección sticky entre sesiones |
| Fallback si elegido falla | **Sí, al chain default** | Resilience > obediencia estricta — el plan es prioridad |

## Catálogo de modelos curados

```ts
export const MODEL_PRESETS = [
  { id: "fast",        label: "🚀 Rápido",        sub: "Cerebras GPT-OSS 120B",     provider: "cerebras", model: "gpt-oss-120b" },
  { id: "experimental", label: "🧪 Experimental",  sub: "Cerebras GLM 4.7 (preview)", provider: "cerebras", model: "zai-glm-4.7" },
  { id: "economico",   label: "⚖️ Económico",      sub: "Cerebras Llama 3.1 8B",      provider: "cerebras", model: "llama3.1-8b" },
  { id: "backup",      label: "🎯 Backup Groq",   sub: "Llama 3.3 70B",              provider: "groq",     model: "llama-3.3-70b-versatile" },
] as const;
```

## Arquitectura

### Backend — `socya_pipeline/ai_chain.py`

1. **Eliminar Gemini**: borra `_ProviderDef` de gemini, `_call_gemini`, rama `style="gemini"` en dispatcher.
2. **Reordenar PROVIDER_DEFS**: `cerebras > groq > openrouter`.
3. **Cerebras models**: `("gpt-oss-120b", "zai-glm-4.7", "llama3.1-8b")` en ese orden.
4. **Nuevo parámetro `preferred_model`** en `AIChain.__init__`:
   - Si está seteado (formato `"provider/model"`), `_enumerate_attempts` lo emite PRIMERO.
   - El chain default sigue como fallback si falla (configurable per-profile).
5. **Nuevo parámetro `single_model_mode`** (bool):
   - Si True: solo prueba el `preferred_model`, no chain. Si falla → AI_SATURATED.
   - Si False (default): preferred_model primero + chain como fallback.

### Plumbing — `socya_pipeline/cli.py`, `socya_pipeline/planner.py`

- `cmd_plan` y `cmd_generate` aceptan `request.get("preferred_model")` y lo pasan a `plan_presentation(...)`.
- `plan_presentation` lo pasa a `AIChain(preferred_model=...)`.
- Cache key incluye `preferred_model` (distinto modelo = distinto cache).

### API — `app/api/preview-plan/route.ts` y `generate-pptx`

- formData lee `preferredModel` y lo añade al request payload pasado a Python.

### UI — `components/PreparePanel.tsx`

- Nuevo selector arriba del prompt (encima de "Tu instrucción para la IA"):
  ```
  Modelo IA: [🚀 Rápido — Cerebras GPT-OSS 120B ▼]
  ```
- Dropdown con los 4 presets.
- Persiste selección en `localStorage["socya:preferred_model"]`.
- Pasa el `provider/model` al `onConfirm` y a las llamadas `/api/preview-plan` + `/api/generate-pptx`.

### UI — `components/ExcelUploader.tsx`

- State `preferredModel` lifted (igual patrón que `userPrompt`).
- Pasado a `PreparePanel` y a las llamadas SSE de generación.

## Tests

- `__tests__/python/test_ai_chain.py`: ajustar tests que esperaban gemini en la chain. Añadir test para `preferred_model` (orden + fallback).
- E2E manual: subir Excel, cambiar selector entre los 4, verificar que el `ai_status.model` del response refleja el elegido.

## Riesgos

- **GLM 4.7 es preview**: Cerebras puede cambiarlo o quitarlo sin aviso. Mitigación: si falla, fallback al chain default lo cubre.
- **Tests viejos**: cualquier test que aserte gemini en `PROVIDER_DEFS` rompe → ajustar.
- **localStorage SSR**: PreparePanel es `"use client"` ya, no hay problema con SSR. Lazy init del state.
- **Cache invalidation**: bumpear `PLANNER_VERSION` para que decks viejos cacheados no se sirvan con un modelo distinto.
