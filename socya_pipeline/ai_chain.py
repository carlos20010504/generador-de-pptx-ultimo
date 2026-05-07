"""OpenRouter call layer with 4-model fallback chain and profile-aware retry."""
import os
import time
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional
import requests
from socya_pipeline.errors import PipelineError, ErrorCode

MODEL_CHAIN = [
    "nousresearch/hermes-3-llama-3.1-405b:free",
    "meta-llama/llama-3.3-70b-instruct:free",
    "google/gemma-2-9b-it:free",
    "qwen/qwen-2.5-72b-instruct:free",
]

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
APP_NAME = os.environ.get("OPENROUTER_APP_NAME", "Socya PPTX Generator")
SITE_URL = os.environ.get("OPENROUTER_SITE_URL", "http://localhost")

class AIProfile(str, Enum):
    FAST = "fast"      # 25s timeout, 1 model only, raise if it fails
    PATIENT = "patient"  # up to 4 min, full chain, multiple cycles

PROFILE_SETTINGS = {
    AIProfile.FAST: {
        "timeout_per_call": 25,
        "max_models_to_try": 1,
        "max_cycles": 1,
        "retry_within_model": 0,
    },
    AIProfile.PATIENT: {
        "timeout_per_call": 60,
        "max_models_to_try": len(MODEL_CHAIN),
        "max_cycles": 4,
        "retry_within_model": 1,
    },
}

RATE_LIMIT_TOKENS = ("rate limit", "high demand", "limit_rpm",
                     "429", "limited to", "temporarily rate-limited")
TRANSIENT_TOKENS = ("upstream error", "timed out", "timeout",
                    "service unavailable", "overloaded", "bad gateway")

@dataclass
class AIChainResult:
    model: str
    content: str
    fallback_steps: List[dict] = field(default_factory=list)
    cache_hit: bool = False

class AIChain:
    def __init__(self, api_key: str, profile: AIProfile = AIProfile.FAST):
        self.api_key = (api_key or "").strip().strip('"').strip("'")
        self.profile = profile
        self.settings = PROFILE_SETTINGS[profile]

    def call(self, prompt: str, system_msg: str = "You must output strictly valid JSON.",
             temperature: float = 0.2) -> AIChainResult:
        if not self.api_key:
            raise PipelineError(
                ErrorCode.AI_SATURATED,
                "Falta OPENROUTER_API_KEY en el entorno.",
                user_action="report_bug",
            )

        fallback_steps: List[dict] = []
        last_error = "unknown"
        models_to_try = MODEL_CHAIN[: self.settings["max_models_to_try"]]

        for cycle in range(self.settings["max_cycles"]):
            for model in models_to_try:
                try:
                    content = self._call_one(model, prompt, system_msg, temperature)
                    return AIChainResult(model=model, content=content,
                                         fallback_steps=fallback_steps)
                except _Retryable as exc:
                    fallback_steps.append({
                        "from": model, "reason": exc.reason,
                        "message": str(exc)[:200],
                    })
                    last_error = exc.reason
                    if exc.retry_after and self.profile == AIProfile.PATIENT:
                        time.sleep(min(exc.retry_after, 90))
                except _Fatal as exc:
                    raise PipelineError(
                        ErrorCode.AI_RESPONSE_INVALID,
                        f"El modelo {model} devolvió una respuesta inválida.",
                        details=str(exc)[:300],
                        user_action="retry",
                    )
            if cycle < self.settings["max_cycles"] - 1:
                time.sleep(min(30 * (cycle + 1), 60))

        raise PipelineError(
            ErrorCode.AI_SATURATED,
            "Todos los modelos IA disponibles están saturados ahora mismo.",
            details=f"Último error: {last_error}. Intentos: {len(fallback_steps)}.",
            user_action="retry_later",
            retry_after_seconds=300,
        )

    def _call_one(self, model: str, prompt: str, system_msg: str,
                   temperature: float) -> str:
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_msg},
                {"role": "user", "content": prompt},
            ],
            "temperature": temperature,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": SITE_URL,
            "X-Title": APP_NAME,
        }
        try:
            resp = requests.post(OPENROUTER_URL, headers=headers, json=payload,
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
            raise _Fatal(f"http_{resp.status_code}: {text[:200]}")

        try:
            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            if not content:
                raise _Fatal("empty content")
            return content
        except (KeyError, IndexError, ValueError) as e:
            raise _Fatal(f"malformed response: {e}")


def _parse_retry_after(headers, body_lower: str) -> int:
    h = headers.get("Retry-After") or headers.get("retry-after")
    if h:
        try:
            return int(float(h))
        except (TypeError, ValueError):
            pass
    import re
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
