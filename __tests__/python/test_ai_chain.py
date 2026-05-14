import pytest
from unittest.mock import MagicMock, patch
from socya_pipeline.ai_chain import AIChain, AIProfile, AIChainResult, MODEL_CHAIN
from socya_pipeline.errors import PipelineError, ErrorCode

@pytest.fixture
def fast_chain():
    return AIChain(api_key="test-key", profile=AIProfile.FAST)

@pytest.fixture
def patient_chain():
    return AIChain(api_key="test-key", profile=AIProfile.PATIENT)

def _mock_response(status_code, json_body=None, headers=None):
    resp = MagicMock()
    resp.status_code = status_code
    resp.ok = 200 <= status_code < 300
    resp.text = ""
    resp.headers = headers or {}
    resp.json.return_value = json_body or {}
    return resp

def test_first_model_success(fast_chain):
    happy = _mock_response(200, {
        "choices": [{"message": {"content": '{"plan":"ok"}'}}]
    })
    with patch("socya_pipeline.ai_chain.requests.post", return_value=happy):
        result = fast_chain.call("prompt")
    # Multi-provider: el modelo ahora va prefijado con el provider name.
    assert result.model == f"openrouter/{MODEL_CHAIN[0]}"
    assert result.content == '{"plan":"ok"}'
    assert result.fallback_steps == []

def test_falls_back_on_429(patient_chain):
    rate_limited = _mock_response(429, {"error": {"message": "rate limit"}})
    happy = _mock_response(200, {"choices":[{"message":{"content":"ok"}}]})
    with patch("socya_pipeline.ai_chain.requests.post",
               side_effect=[rate_limited, happy]), \
         patch("socya_pipeline.ai_chain.time.sleep"):
        result = patient_chain.call("prompt")
    assert result.model == f"openrouter/{MODEL_CHAIN[1]}"
    assert len(result.fallback_steps) == 1
    assert result.fallback_steps[0]["from"] == f"openrouter/{MODEL_CHAIN[0]}"
    assert result.fallback_steps[0]["reason"] == "rate_limited"

def test_all_saturated_raises_ai_saturated(fast_chain):
    rate_limited = _mock_response(429, {"error": {"message": "rate limit"}})
    with patch("socya_pipeline.ai_chain.requests.post", return_value=rate_limited):
        with pytest.raises(PipelineError) as exc:
            fast_chain.call("prompt")
    assert exc.value.code == ErrorCode.AI_SATURATED

def test_no_api_key_raises():
    chain = AIChain(api_key="", profile=AIProfile.FAST)
    with pytest.raises(PipelineError) as exc:
        chain.call("prompt")
    assert exc.value.code == ErrorCode.AI_SATURATED

def test_falls_back_on_404(patient_chain):
    """A retired/missing model (404) should not abort the chain — try next model."""
    not_found = _mock_response(404, {"error": {"message": "Model not found"}})
    happy = _mock_response(200, {"choices":[{"message":{"content":"ok"}}]})
    with patch("socya_pipeline.ai_chain.requests.post",
               side_effect=[not_found, happy]), \
         patch("socya_pipeline.ai_chain.time.sleep"):
        result = patient_chain.call("prompt")
    assert result.model == f"openrouter/{MODEL_CHAIN[1]}"
    assert len(result.fallback_steps) == 1
    assert result.fallback_steps[0]["from"] == f"openrouter/{MODEL_CHAIN[0]}"
    assert "404" in result.fallback_steps[0]["reason"] or "http_404" in result.fallback_steps[0]["reason"]


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
