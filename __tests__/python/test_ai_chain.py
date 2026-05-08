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
    assert result.model == MODEL_CHAIN[0]
    assert result.content == '{"plan":"ok"}'
    assert result.fallback_steps == []

def test_falls_back_on_429(patient_chain):
    rate_limited = _mock_response(429, {"error": {"message": "rate limit"}})
    happy = _mock_response(200, {"choices":[{"message":{"content":"ok"}}]})
    with patch("socya_pipeline.ai_chain.requests.post",
               side_effect=[rate_limited, happy]), \
         patch("socya_pipeline.ai_chain.time.sleep"):
        result = patient_chain.call("prompt")
    assert result.model == MODEL_CHAIN[1]
    assert len(result.fallback_steps) == 1
    assert result.fallback_steps[0]["from"] == MODEL_CHAIN[0]
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
    assert result.model == MODEL_CHAIN[1]
    assert len(result.fallback_steps) == 1
    assert result.fallback_steps[0]["from"] == MODEL_CHAIN[0]
    assert "404" in result.fallback_steps[0]["reason"] or "http_404" in result.fallback_steps[0]["reason"]
