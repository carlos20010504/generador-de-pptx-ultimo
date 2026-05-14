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
import os
import re
import time
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional
import requests
from socya_pipeline.errors import PipelineError, ErrorCode

# OpenRouter free chain (tier histórico) — last-resort si los providers más
# rápidos no están configurados o saturaron todos.
MODEL_CHAIN = [
    "openai/gpt-oss-120b:free",
    "google/gemma-4-31b-it:free",
    "z-ai/glm-4.5-air:free",
    "nvidia/nemotron-3-super-120b-a12b:free",
    "openai/gpt-oss-20b:free",
]

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
APP_NAME = os.environ.get("OPENROUTER_APP_NAME", "Socya PPTX Generator")
SITE_URL = os.environ.get("OPENROUTER_SITE_URL", "http://localhost")


# ─────────────────────────────────────────────────────────────────────
# Provider definitions
# ─────────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class _ProviderDef:
    name: str
    base_url: str             # endpoint URL
    api_key_env: str          # env var con la key
    models: tuple             # modelos a probar en orden
    style: str                # "openai_compat"


# IMPORTANTE: el orden importa. Más rápido y estable arriba.
# Para activar un provider, el usuario solo setea su env var.
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


@dataclass
class _ProviderRuntime:
    definition: _ProviderDef
    api_key: str


# ─────────────────────────────────────────────────────────────────────
# Profile settings
# ─────────────────────────────────────────────────────────────────────

class AIProfile(str, Enum):
    FAST = "fast"      # 25s timeout, 1 model only, raise if it fails
    PATIENT = "patient"  # full chain multi-provider, hasta ~3 min cap duro

# `total_budget_seconds` es el techo absoluto del tiempo que la cadena puede
# bloquear el request. La API route tiene maxDuration=280s; superarlo solo
# desperdicia recursos del server y frustra al usuario con un freeze más
# largo del que ya iba a romperse.
#
# `max_models_to_try` ahora es el cap GLOBAL across providers. Con todos los
# providers configurados podemos tener hasta 11 modelos en la cadena
# (3 Cerebras + 3 Groq + 5 OpenRouter), pero rara vez probamos tantos
# antes de conseguir respuesta.
PROFILE_SETTINGS = {
    AIProfile.FAST: {
        "timeout_per_call": 25,
        "max_models_to_try": 1,
        "max_cycles": 1,
        "retry_within_model": 0,
        "total_budget_seconds": 60,
    },
    AIProfile.PATIENT: {
        "timeout_per_call": 60,
        "max_models_to_try": 13,
        "max_cycles": 4,
        "retry_within_model": 1,
        # 200s deja ~80s de margen antes del maxDuration=280s del API route.
        "total_budget_seconds": 200,
    },
}

RATE_LIMIT_TOKENS = ("rate limit", "high demand", "limit_rpm",
                     "429", "limited to", "temporarily rate-limited",
                     "quota exceeded", "resource_exhausted")
TRANSIENT_TOKENS = ("upstream error", "timed out", "timeout",
                    "service unavailable", "overloaded", "bad gateway")


# ─────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────

@dataclass
class AIChainResult:
    model: str                # "{provider}/{model}" cuando hay multi-provider
    content: str
    fallback_steps: List[dict] = field(default_factory=list)
    cache_hit: bool = False


class AIChain:
    def __init__(self, api_key: str, profile: AIProfile = AIProfile.FAST,
                  preferred_model: Optional[str] = None):
        """preferred_model formato: 'provider/model' (ej. 'cerebras/gpt-oss-120b').
        Si está seteado, _enumerate_attempts lo emite PRIMERO antes del chain
        default. Si falla, cae al chain por resiliencia.
        """
        # api_key se trata como override de OPENROUTER_API_KEY (back-compat
        # con el código que existía cuando solo hablábamos con OpenRouter).
        # Los otros providers leen su env var dedicada.
        self._openrouter_key = (api_key or "").strip().strip('"').strip("'")
        self.profile = profile
        self.settings = PROFILE_SETTINGS[profile]
        self._providers: List[_ProviderRuntime] = self._discover_providers()
        self.preferred_model = (preferred_model or "").strip() or None

    def _discover_providers(self) -> List[_ProviderRuntime]:
        """Enumera providers con key configurada. OpenRouter usa el key del
        constructor; los demás leen su env var dedicada. Si solo hay
        OPENROUTER_API_KEY (caso histórico), el resultado es la lista vieja."""
        out: List[_ProviderRuntime] = []
        for pdef in PROVIDER_DEFS:
            if pdef.name == "openrouter":
                key = self._openrouter_key
            else:
                key = (os.environ.get(pdef.api_key_env, "") or "").strip()
                key = key.strip('"').strip("'")
            if key:
                out.append(_ProviderRuntime(definition=pdef, api_key=key))
        return out

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

    def call(self, prompt: str, system_msg: str = "You must output strictly valid JSON.",
             temperature: float = 0.2) -> AIChainResult:
        if not self._providers:
            raise PipelineError(
                ErrorCode.AI_SATURATED,
                "No hay ningún proveedor de IA configurado. Setea al menos una de: "
                "CEREBRAS_API_KEY, GROQ_API_KEY, OPENROUTER_API_KEY.",
                user_action="report_bug",
            )

        fallback_steps: List[dict] = []
        last_error = "unknown"

        # Techo duro: si superamos este budget, abortamos con AI_SATURATED.
        budget = self.settings.get("total_budget_seconds", 200)
        start = time.monotonic()

        def _budget_left() -> float:
            return max(0.0, budget - (time.monotonic() - start))

        for cycle in range(self.settings["max_cycles"]):
            if _budget_left() <= 0:
                last_error = "budget_exhausted"
                break
            for provider, model in self._enumerate_attempts():
                if _budget_left() <= 0:
                    last_error = "budget_exhausted"
                    break
                model_id = f"{provider.definition.name}/{model}"
                try:
                    content = self._call_provider(
                        provider, model, prompt, system_msg, temperature)
                    return AIChainResult(
                        model=model_id, content=content,
                        fallback_steps=fallback_steps)
                except _Retryable as exc:
                    fallback_steps.append({
                        "from": model_id, "reason": exc.reason,
                        "message": str(exc)[:200],
                    })
                    last_error = exc.reason
                    if exc.retry_after and self.profile == AIProfile.PATIENT:
                        # Respeta el budget — nunca dormir más de lo que queda.
                        nap = min(exc.retry_after, 90, _budget_left())
                        if nap > 0:
                            time.sleep(nap)
                except _Fatal as exc:
                    raise PipelineError(
                        ErrorCode.AI_RESPONSE_INVALID,
                        f"El modelo {model_id} devolvió una respuesta inválida.",
                        details=str(exc)[:300],
                        user_action="retry",
                    )
            if cycle < self.settings["max_cycles"] - 1:
                cycle_nap = min(30 * (cycle + 1), 60, _budget_left())
                if cycle_nap > 0:
                    time.sleep(cycle_nap)

        # Build a richer details payload so the user (and our logs) can see
        # WHICH providers/models tried, WHICH reason dominated, y si vale la
        # pena reintentar o ya está fuera de servicio.
        n_models = sum(1 for _ in self._enumerate_attempts())
        n_attempts = len(fallback_steps)
        n_cycles = self.settings["max_cycles"]
        n_providers = len(self._providers)
        reason_counts: dict = {}
        for step in fallback_steps:
            r = step.get("reason", "unknown")
            reason_counts[r] = reason_counts.get(r, 0) + 1
        reasons_str = ", ".join(
            f"{r}:{n}" for r, n in sorted(reason_counts.items(), key=lambda kv: -kv[1])
        ) or "unknown"
        provider_names = ", ".join(p.definition.name for p in self._providers)
        details = (
            f"Último error: {last_error}. "
            f"Providers activos: {provider_names} ({n_providers}). "
            f"Modelos cap: {n_models}, ciclos: {n_cycles}, intentos totales: {n_attempts}. "
            f"Razones: {reasons_str}."
        )
        raise PipelineError(
            ErrorCode.AI_SATURATED,
            "Todos los modelos IA disponibles están saturados ahora mismo.",
            details=details,
            user_action="retry_later",
            retry_after_seconds=300,
        )

    # ─────────────────────────────────────────────────────────────────
    # Provider dispatch
    # ─────────────────────────────────────────────────────────────────

    def _call_provider(self, provider: _ProviderRuntime, model: str,
                        prompt: str, system_msg: str, temperature: float) -> str:
        if provider.definition.style == "openai_compat":
            return self._call_openai_compat(provider, model, prompt, system_msg,
                                              temperature)
        raise _Fatal(f"unknown provider style: {provider.definition.style}")

    def _call_openai_compat(self, provider: _ProviderRuntime, model: str,
                              prompt: str, system_msg: str, temperature: float) -> str:
        """Funciona para Groq, Cerebras, OpenRouter, y cualquier endpoint
        compatible con OpenAI Chat Completions API."""
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_msg},
                {"role": "user", "content": prompt},
            ],
            "temperature": temperature,
        }
        headers = {
            "Authorization": f"Bearer {provider.api_key}",
            "Content-Type": "application/json",
        }
        # OpenRouter pide attribution headers (no obligatorios pero ayudan
        # con cuotas + ranking). Otros providers los ignoran.
        if provider.definition.name == "openrouter":
            headers["HTTP-Referer"] = SITE_URL
            headers["X-Title"] = APP_NAME

        try:
            resp = requests.post(provider.definition.base_url, headers=headers,
                                  json=payload,
                                  timeout=self.settings["timeout_per_call"])
        except requests.Timeout:
            raise _Retryable("timeout", "request timed out")
        except requests.RequestException as e:
            raise _Retryable("network_error", str(e))

        text = resp.text or ""
        lower = text.lower()

        if resp.status_code == 429:
            retry_after = _parse_retry_after(resp.headers, lower)
            raise _Retryable("rate_limited", text[:200], retry_after=retry_after)
        if not resp.ok:
            if any(t in lower for t in RATE_LIMIT_TOKENS):
                raise _Retryable("rate_limited", text[:200])
            if any(t in lower for t in TRANSIENT_TOKENS):
                raise _Retryable("transient", text[:200])
            raise _Retryable(f"http_{resp.status_code}", text[:200])

        try:
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            if not content:
                raise _Fatal("empty content")
            return content
        except (KeyError, IndexError, ValueError) as e:
            raise _Fatal(f"malformed response: {e}")

    # Back-compat shim: tests viejos parchean `requests.post` y esperan que
    # `_call_one(model, ...)` use OpenRouter directamente. Lo dejamos como
    # wrapper sobre el nuevo dispatcher para no romper tests existentes.
    def _call_one(self, model: str, prompt: str, system_msg: str,
                   temperature: float) -> str:
        # Encuentra el runtime de OpenRouter (el constructor lo registra
        # cuando se pasa api_key).
        openrouter = next(
            (p for p in self._providers if p.definition.name == "openrouter"),
            None)
        if openrouter is None:
            raise _Fatal("openrouter not configured (legacy _call_one path)")
        return self._call_openai_compat(openrouter, model, prompt, system_msg,
                                          temperature)


# ─────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────

def _parse_retry_after(headers, body_lower: str) -> int:
    h = headers.get("Retry-After") or headers.get("retry-after")
    if h:
        try:
            return int(float(h))
        except (TypeError, ValueError):
            pass
    m = re.search(r"retry in\s+([0-9]+(?:\.[0-9]+)?)\s*s", body_lower)
    if m:
        return int(float(m.group(1)))
    return 0


class _Retryable(Exception):
    def __init__(self, reason: str, message: str = "", retry_after: int = 0):
        self.reason = reason
        self.retry_after = retry_after
        super().__init__(message)


class _Fatal(Exception):
    pass
