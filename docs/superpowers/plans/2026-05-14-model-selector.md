# Model Selector + Cerebras Primary — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development.

**Goal:** Reordenar AI chain con Cerebras primero (GPT-OSS, GLM 4.7, Llama 3.1 8B), eliminar Gemini, y exponer un selector curado de 4 modelos en la UI.

**Architecture:** `ai_chain.py` acepta `preferred_model` que se prueba primero antes del chain default. UI dropdown en PreparePanel persiste en localStorage. Backend cli/route plumbing.

**Tech Stack:** Python (ai_chain), Next.js 16 + React 19 (UI), localStorage para persistencia.

**Spec:** `docs/superpowers/specs/2026-05-14-model-selector-design.md`

---

## File Structure

| Archivo | Tipo | Responsabilidad |
|---|---|---|
| `socya_pipeline/ai_chain.py` | MODIFY | Cerebras primario, eliminar Gemini, añadir `preferred_model` param |
| `socya_pipeline/planner.py` | MODIFY | Pasar `preferred_model` a AIChain, incluir en cache key |
| `socya_pipeline/cli.py` | MODIFY | Leer `preferredModel` del request, pasar a planner |
| `app/api/preview-plan/route.ts` | MODIFY | Leer `preferredModel` del formData, pasar al Python |
| `app/api/generate-pptx/route.ts` | MODIFY | Igual que arriba |
| `components/PreparePanel.tsx` | MODIFY | Selector dropdown + estado, pasar al onConfirm |
| `components/ExcelUploader.tsx` | MODIFY | Lift `preferredModel` state, pasar a PreparePanel y a SSE |
| `__tests__/python/test_ai_chain.py` | MODIFY | Ajustar tests que asumen Gemini en chain |

---

## Task 1: Reordenar AI chain + eliminar Gemini + `preferred_model`

**Files:** Modify `socya_pipeline/ai_chain.py`, `__tests__/python/test_ai_chain.py`

- [ ] **Step 1: Read `socya_pipeline/ai_chain.py` y `test_ai_chain.py`** para entender la estructura y qué tests existen.

- [ ] **Step 2: Modificar PROVIDER_DEFS.** Reemplazar el tuple completo (líneas ~59-91) con:

```python
PROVIDER_DEFS: tuple = (
    _ProviderDef(
        name="cerebras",
        base_url="https://api.cerebras.ai/v1/chat/completions",
        api_key_env="CEREBRAS_API_KEY",
        models=("gpt-oss-120b", "zai-glm-4.7", "llama3.1-8b"),
        style="openai_compat",
    ),
    _ProviderDef(
        name="groq",
        base_url="https://api.groq.com/openai/v1/chat/completions",
        api_key_env="GROQ_API_KEY",
        models=("llama-3.3-70b-versatile", "llama-3.1-8b-instant",
                "mixtral-8x7b-32768"),
        style="openai_compat",
    ),
    _ProviderDef(
        name="openrouter",
        base_url=OPENROUTER_URL,
        api_key_env="OPENROUTER_API_KEY",
        models=tuple(MODEL_CHAIN),
        style="openai_compat",
    ),
)
```

- [ ] **Step 3: Eliminar el método `_call_gemini`** (líneas ~345-408) completamente.

- [ ] **Step 4: Eliminar la rama `style="gemini"`** del dispatcher `_call_provider` (~líneas 287-289). Después del cambio el método queda:

```python
def _call_provider(self, provider: _ProviderRuntime, model: str,
                    prompt: str, system_msg: str, temperature: float) -> str:
    if provider.definition.style == "openai_compat":
        return self._call_openai_compat(provider, model, prompt, system_msg,
                                          temperature)
    raise _Fatal(f"unknown provider style: {provider.definition.style}")
```

- [ ] **Step 5: Actualizar el docstring del módulo** (líneas 1-18). Reemplazar la línea sobre Gemini y reordenar la lista de providers:

```python
"""Multi-provider AI call layer.

Soporta múltiples proveedores con detección automática:
  - Cerebras (paga, súper rápido — ~2000 tok/s, gpt-oss-120b primario)
  - Groq (free tier, ~300 tok/s, fallback rápido)
  - OpenRouter (free tier, fallback final)

Cada provider se activa SOLO si su API key está en env. Si solo hay
OPENROUTER_API_KEY (caso histórico), el comportamiento es idéntico al de
antes. El usuario amplía la red simplemente seteando otras keys.

El orden por defecto privilegia velocidad y calidad:
  Cerebras > Groq > OpenRouter
"""
```

También actualizar el mensaje en `call()` cuando no hay providers (línea ~196):

Find:
```python
                "GROQ_API_KEY, CEREBRAS_API_KEY, GEMINI_API_KEY, OPENROUTER_API_KEY.",
```
Replace with:
```python
                "CEREBRAS_API_KEY, GROQ_API_KEY, OPENROUTER_API_KEY.",
```

- [ ] **Step 6: Añadir `preferred_model` param a `__init__`.** Reemplazar:

```python
def __init__(self, api_key: str, profile: AIProfile = AIProfile.FAST):
```

con:

```python
def __init__(self, api_key: str, profile: AIProfile = AIProfile.FAST,
              preferred_model: Optional[str] = None):
    """preferred_model formato: 'provider/model' (ej. 'cerebras/gpt-oss-120b').
    Si está seteado, _enumerate_attempts lo emite PRIMERO antes del chain
    default. Si falla, cae al chain por resiliencia.
    """
```

Y al final del cuerpo del `__init__` (después de `self._providers = ...`):

```python
        self.preferred_model = (preferred_model or "").strip() or None
```

- [ ] **Step 7: Modificar `_enumerate_attempts`** para emitir el preferred_model primero. Reemplazar el método actual con:

```python
def _enumerate_attempts(self):
    """Yield (provider, model) en orden: preferred_model primero (si configurado
    y disponible), después el chain default. Capeado por max_models_to_try."""
    cap = self.settings["max_models_to_try"]
    n = 0
    yielded: set = set()

    if self.preferred_model and "/" in self.preferred_model:
        prov_name, model_name = self.preferred_model.split("/", 1)
        prov = next((p for p in self._providers
                     if p.definition.name == prov_name), None)
        if prov and model_name in prov.definition.models:
            if n < cap:
                yield prov, model_name
                yielded.add((prov_name, model_name))
                n += 1

    for provider in self._providers:
        for model in provider.definition.models:
            if (provider.definition.name, model) in yielded:
                continue
            if n >= cap:
                return
            yield provider, model
            n += 1
```

- [ ] **Step 8: Ajustar tests existentes.** En `__tests__/python/test_ai_chain.py`, buscar referencias a "gemini" o `GEMINI_API_KEY` y eliminarlas. Buscar tests que asuman provider order (groq primero) y ajustar a nuevo orden (cerebras primero). Si algún test instancia `AIChain(...)` con expectativas específicas del orden, actualizar.

- [ ] **Step 9: Añadir tests nuevos** al final de `test_ai_chain.py`:

```python
def test_preferred_model_yielded_first(monkeypatch):
    """Cuando preferred_model está set, _enumerate_attempts lo emite primero."""
    from socya_pipeline.ai_chain import AIChain, AIProfile
    monkeypatch.setenv("CEREBRAS_API_KEY", "test-cerebras")
    monkeypatch.setenv("GROQ_API_KEY", "test-groq")
    chain = AIChain(api_key="", profile=AIProfile.PATIENT,
                     preferred_model="groq/llama-3.3-70b-versatile")
    attempts = list(chain._enumerate_attempts())
    assert attempts[0][0].definition.name == "groq"
    assert attempts[0][1] == "llama-3.3-70b-versatile"


def test_preferred_model_falls_back_to_chain(monkeypatch):
    """Después del preferred, el resto del chain queda como fallback."""
    from socya_pipeline.ai_chain import AIChain, AIProfile
    monkeypatch.setenv("CEREBRAS_API_KEY", "test-cerebras")
    monkeypatch.setenv("GROQ_API_KEY", "test-groq")
    chain = AIChain(api_key="", profile=AIProfile.PATIENT,
                     preferred_model="groq/llama-3.3-70b-versatile")
    attempts = list(chain._enumerate_attempts())
    # First is preferred. Then chain order (cerebras > groq remaining > ...)
    assert attempts[0] == (next(p for p in chain._providers if p.definition.name == "groq"),
                            "llama-3.3-70b-versatile")
    # Cerebras models should come after preferred (next in chain order)
    cerebras_in_attempts = [a for a in attempts if a[0].definition.name == "cerebras"]
    assert len(cerebras_in_attempts) >= 1


def test_preferred_model_unknown_ignored(monkeypatch):
    """preferred_model con provider/modelo inexistente se ignora silenciosamente."""
    from socya_pipeline.ai_chain import AIChain, AIProfile
    monkeypatch.setenv("CEREBRAS_API_KEY", "test")
    chain = AIChain(api_key="", profile=AIProfile.FAST,
                     preferred_model="nonexistent/foo")
    attempts = list(chain._enumerate_attempts())
    # Only the FAST cap=1 default attempt; no crash
    assert len(attempts) == 1


def test_no_gemini_in_provider_defs():
    """Sanity: gemini fue eliminado de PROVIDER_DEFS."""
    from socya_pipeline.ai_chain import PROVIDER_DEFS
    names = {p.name for p in PROVIDER_DEFS}
    assert "gemini" not in names
    assert "cerebras" in names
    assert "groq" in names
    assert "openrouter" in names


def test_cerebras_is_first_provider():
    """Cerebras debe estar primero en PROVIDER_DEFS para tener prioridad."""
    from socya_pipeline.ai_chain import PROVIDER_DEFS
    assert PROVIDER_DEFS[0].name == "cerebras"
    assert "gpt-oss-120b" in PROVIDER_DEFS[0].models
```

- [ ] **Step 10: Run tests.**

```bash
python -m pytest __tests__/python/test_ai_chain.py -v --tb=short
```

Expected: all pass. Si algún viejo test falla por orden distinto, ajustar.

- [ ] **Step 11: Run full suite.**

```bash
python -m pytest __tests__/python/ -q --tb=short
```

Expected: 160 + 5 nuevos = 165 tests, todos verdes.

- [ ] **Step 12: Commit.**

```bash
git add socya_pipeline/ai_chain.py __tests__/python/test_ai_chain.py
git commit -m "feat(ai_chain): Cerebras primario con gpt-oss-120b + preferred_model param + drop Gemini"
```

---

## Task 2: Plumb preferred_model through planner + CLI

**Files:** Modify `socya_pipeline/planner.py`, `socya_pipeline/cli.py`

- [ ] **Step 1: Modificar `plan_presentation`.** En `socya_pipeline/planner.py`, busca la firma de `plan_presentation` (alrededor de línea 327). Añadir `preferred_model: Optional[str] = None` como último parámetro, y pasarlo al `AIChain(...)`:

Find:
```python
    chain = AIChain(api_key=api_key, profile=profile)
```
Replace with:
```python
    chain = AIChain(api_key=api_key, profile=profile,
                     preferred_model=preferred_model)
```

Y modificar la firma:
```python
def plan_presentation(wb: WorkbookData, blocks, user_prompt: str, audience: str,
                       language: str, api_key: str,
                       profile: AIProfile = AIProfile.PATIENT,
                       cache_dir: Optional[Path] = None,
                       file_path: Optional[Path] = None,
                       intent=None,
                       preferred_model: Optional[str] = None) -> dict:
```

- [ ] **Step 2: Incluir preferred_model en cache_key.** En el bloque que computa `intent_signature`, añadir el preferred_model al final:

```python
            intent_signature = (f"|sc={intent.requested_slide_count}"
                                 f"|rs={req_sheets}"
                                 f"|pm={preferred_model or ''}")
```

- [ ] **Step 3: Bump PLANNER_VERSION** de `"p4"` → `"p5"` (línea 12) para invalidar caches de runs previos.

- [ ] **Step 4: Modificar cmd_plan en cli.py.** Buscar la llamada a `plan_presentation(...)` en `cmd_plan` y añadir `preferred_model=request.get("preferredModel")`:

```python
        plan = plan_presentation(
            wb, inv,
            user_prompt=request.get("prompt", ""),
            audience=request.get("audience", "ejecutivos"),
            language=request.get("language", "es"),
            api_key=api_key,
            profile=profile,
            file_path=Path(args.input),
            intent=intent,
            preferred_model=request.get("preferredModel"),
        )
```

- [ ] **Step 5: Igual en cmd_generate.** Mismo cambio en la otra llamada a `plan_presentation` dentro de `cmd_generate`.

- [ ] **Step 6: Run tests.**

```bash
python -m pytest __tests__/python/ -q --tb=short
```

Expected: 165 verdes.

- [ ] **Step 7: Commit.**

```bash
git add socya_pipeline/planner.py socya_pipeline/cli.py
git commit -m "feat(planner): plumb preferred_model through plan_presentation + CLI + cache key"
```

---

## Task 3: API routes pasan preferredModel al Python

**Files:** Modify `app/api/preview-plan/route.ts`, `app/api/generate-pptx/route.ts`

- [ ] **Step 1: preview-plan/route.ts.** Buscar el bloque que arma `requestPayload` (alrededor de línea 64). Reemplazar:

```ts
    const userPrompt = String(formData.get('userPrompt') ?? '').trim();
    const audience = String(formData.get('audience') ?? 'ejecutivos').trim();
    const language = String(formData.get('language') ?? 'es').trim();

    const requestPayload = JSON.stringify({
      prompt: userPrompt, audience, language,
    });
```

con:

```ts
    const userPrompt = String(formData.get('userPrompt') ?? '').trim();
    const audience = String(formData.get('audience') ?? 'ejecutivos').trim();
    const language = String(formData.get('language') ?? 'es').trim();
    const preferredModel = String(formData.get('preferredModel') ?? '').trim() || undefined;

    const requestPayload = JSON.stringify({
      prompt: userPrompt, audience, language, preferredModel,
    });
```

- [ ] **Step 2: generate-pptx/route.ts.** Mismo patrón — leer `preferredModel` del formData y añadirlo al payload que va a Python. Buscar el spot donde se construye el JSON request al subprocess Python.

Si el route es SSE-based con stream de chunks, busca donde se serializa el request inicial (probablemente similar a preview-plan).

- [ ] **Step 3: Verificar tsc compile.**

```bash
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 4: Commit.**

```bash
git add app/api/preview-plan/route.ts app/api/generate-pptx/route.ts
git commit -m "feat(api): forward preferredModel from formData to Python pipeline"
```

---

## Task 4: UI selector + state lift

**Files:** Modify `components/PreparePanel.tsx`, `components/ExcelUploader.tsx`

- [ ] **Step 1: PreparePanel — definir presets y dropdown.** Cerca del top de `components/PreparePanel.tsx` (después de los imports, antes del componente), añadir:

```tsx
interface ModelPreset {
  id: string;
  label: string;
  sub: string;
  provider: string;
  model: string;
}

const MODEL_PRESETS: ModelPreset[] = [
  { id: "fast",         label: "🚀 Rápido",       sub: "Cerebras GPT-OSS 120B",      provider: "cerebras", model: "gpt-oss-120b" },
  { id: "experimental", label: "🧪 Experimental", sub: "Cerebras GLM 4.7 (preview)", provider: "cerebras", model: "zai-glm-4.7" },
  { id: "economico",    label: "⚖️ Económico",     sub: "Cerebras Llama 3.1 8B",      provider: "cerebras", model: "llama3.1-8b" },
  { id: "backup",       label: "🎯 Backup Groq",  sub: "Llama 3.3 70B",               provider: "groq",     model: "llama-3.3-70b-versatile" },
];

const DEFAULT_PRESET_ID = "fast";
const PREFERRED_MODEL_LS_KEY = "socya:preferred_model_id";
```

- [ ] **Step 2: Añadir props al componente.** En la `interface Props` (alrededor de línea 62), añadir:

```ts
  preferredModelId: string;
  onPreferredModelChange: (id: string) => void;
```

Y en la firma del componente (alrededor de línea 109):

```tsx
export default function PreparePanel({
  file, userPrompt, onPromptChange,
  audience, language, theme, onThemeChange, mode, onModeChange,
  onConfirm, onOpenAdvanced,
  preferredModelId, onPreferredModelChange,
}: Props) {
```

- [ ] **Step 3: Pasar `preferredModel` en los fetch calls** dentro de los useEffects de quick-summary y preview-plan. Encontrar la línea `if (userPrompt.trim()) fd.append('userPrompt', userPrompt.trim());` en el effect de preview-plan y añadir DESPUÉS:

```tsx
        const preset = MODEL_PRESETS.find(p => p.id === preferredModelId);
        if (preset) {
          fd.append('preferredModel', `${preset.provider}/${preset.model}`);
        }
```

También añadir `preferredModelId` al deps array de ese useEffect (junto con `promptKey`):

```tsx
  }, [promptKey, audience, language, preferredModelId]);
```

- [ ] **Step 4: Renderizar el selector.** En el JSX, encontrar la sección "Refinar prompt y contexto" (la que arrancamos por default). DENTRO del `prep-refine-body` div, ANTES del `<textarea value={userPrompt}...>`, insertar:

```tsx
            <div className="prep-model-row">
              <label className="prep-model-label">Modelo IA</label>
              <select
                value={preferredModelId}
                onChange={(e) => onPreferredModelChange(e.target.value)}
                className="prep-model-select"
              >
                {MODEL_PRESETS.map(p => (
                  <option key={p.id} value={p.id}>
                    {p.label} — {p.sub}
                  </option>
                ))}
              </select>
            </div>
```

- [ ] **Step 5: CSS para el selector.** Al final de `PREP_STYLES` (antes del closing backtick), añadir:

```css
/* ── Model selector ── */
.prep-model-row {
  display: flex; align-items: center; gap: 0.6rem;
  margin-bottom: 0.5rem;
}
.prep-model-label {
  color: var(--c-text-secondary);
  font-family: var(--font-heading);
  font-size: 0.66rem; font-weight: 700;
  text-transform: uppercase; letter-spacing: 0.06em;
  flex-shrink: 0;
}
.prep-model-select {
  flex: 1;
  padding: 0.45rem 0.6rem;
  background: var(--c-bg-elevated);
  border: 1px solid var(--c-border-strong);
  border-radius: var(--r-md);
  color: var(--c-text-primary);
  font-family: var(--font-sans);
  font-size: 0.78rem;
  cursor: pointer;
}
.prep-model-select:focus {
  outline: none;
  border-color: var(--c-primary);
  box-shadow: 0 0 0 3px rgba(8, 112, 98, 0.15);
}
```

- [ ] **Step 6: ExcelUploader lift state.** En `components/ExcelUploader.tsx` cerca del state de `userPrompt` (línea ~95), añadir:

```tsx
  const [preferredModelId, setPreferredModelId] = useState<string>(() => {
    if (typeof window === 'undefined') return 'fast';
    return window.localStorage.getItem('socya:preferred_model_id') || 'fast';
  });

  useEffect(() => {
    if (typeof window === 'undefined') return;
    window.localStorage.setItem('socya:preferred_model_id', preferredModelId);
  }, [preferredModelId]);
```

- [ ] **Step 7: Pasar a PreparePanel.** Encontrar el `<PreparePanel ... />` JSX y añadir las 2 props nuevas:

```tsx
              preferredModelId={preferredModelId}
              onPreferredModelChange={setPreferredModelId}
```

(Hay 2 instancias de `<PreparePanel>` por flujo dual; añadir a ambas.)

- [ ] **Step 8: Pasar `preferredModel` a la SSE de generate.** Buscar donde se hace `formData.append(...)` para `/api/generate-pptx` (en ExcelUploader) y añadir:

```tsx
    const preset = MODEL_PRESETS_REF.find(p => p.id === preferredModelId);
    if (preset) formData.append('preferredModel', `${preset.provider}/${preset.model}`);
```

Necesitas duplicar `MODEL_PRESETS` (o exportarlo de PreparePanel y reusarlo). La opción cleaner: crear un helper file `lib/model-presets.ts`:

```ts
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
```

Y reusarlo en ambos: PreparePanel y ExcelUploader. Importar con `import { MODEL_PRESETS, DEFAULT_PRESET_ID, PREFERRED_MODEL_LS_KEY, presetToModelString } from '@/lib/model-presets'`.

- [ ] **Step 9: tsc compile check.**

```bash
npx tsc --noEmit
```

Expected: no errors.

- [ ] **Step 10: Commit.**

```bash
git add components/PreparePanel.tsx components/ExcelUploader.tsx lib/model-presets.ts
git commit -m "feat(ui): model selector dropdown with localStorage persistence"
```

---

## Task 5: E2E verification

**Files:** none (testing only)

- [ ] **Step 1: Smoke test backend con preferred_model.**

```bash
SOCYA_AI_PROFILE=fast python -X utf8 -m socya_pipeline plan \
  --input "tests/fixtures/comisiones.xlsx" \
  --request '{"prompt":"9 slides","preferredModel":"cerebras/gpt-oss-120b"}'
```

Inspeccionar `ai_status.model` del JSON response — debe ser `cerebras/gpt-oss-120b`.

- [ ] **Step 2: Smoke con backup Groq.**

Cambiar a `"preferredModel":"groq/llama-3.3-70b-versatile"`. Verificar `ai_status.model = groq/...`.

- [ ] **Step 3: Smoke con preferred inexistente.** `"preferredModel":"foo/bar"`. Debe ignorar y caer al default (cerebras).

- [ ] **Step 4: UI smoke.** `npm run dev`, abrir, verificar que el dropdown aparece y persiste entre reloads (localStorage).

- [ ] **Step 5: Commit findings.** Si todo verde, append una sección "Findings" al plan file y commitear.

```bash
git add docs/superpowers/plans/2026-05-14-model-selector.md
git commit -m "test: e2e verification of model selector — Cerebras gpt-oss-120b primary"
```
