import json
import pytest
from unittest.mock import patch, MagicMock
from socya_pipeline.parser import parse_workbook
from socya_pipeline.inventory import build_inventory
from socya_pipeline.planner import build_payload, plan_presentation, PLANNER_VERSION

def test_build_payload_under_token_budget(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    blocks = build_inventory(wb)
    payload = build_payload(wb, blocks, user_prompt="resume", audience="ejecutivos",
                              language="es")
    serialized = json.dumps(payload, ensure_ascii=False, default=str)
    # Heuristic: ~4 chars per token. 6K tokens ≈ 24K chars (with some headroom).
    assert len(serialized) < 30_000

def test_payload_includes_block_ids(ventas_simple_path):
    wb = parse_workbook(ventas_simple_path)
    blocks = build_inventory(wb)
    payload = build_payload(wb, blocks, "p", "a", "es")
    payload_block_ids = {b["id"] for b in payload["workbook"]["blocks"]}
    inventory_ids = {b.id for b in blocks}
    assert payload_block_ids == inventory_ids

def test_plan_calls_ai_chain(ventas_simple_path, stub_planner_response):
    wb = parse_workbook(ventas_simple_path)
    blocks = build_inventory(wb)
    fake_result = MagicMock(model="hermes", content=json.dumps(stub_planner_response),
                             fallback_steps=[], cache_hit=False)
    with patch("socya_pipeline.planner.AIChain") as MockChain:
        MockChain.return_value.call.return_value = fake_result
        plan = plan_presentation(wb, blocks, user_prompt="p", audience="a",
                                  language="es", api_key="key")
    assert plan["slides"][0]["type"] == "title"
    assert plan["_meta"]["model"] == "hermes"

def test_plan_uses_cache_on_second_call(ventas_simple_path, stub_planner_response,
                                          tmp_path):
    wb = parse_workbook(ventas_simple_path)
    blocks = build_inventory(wb)
    fake_result = MagicMock(model="hermes", content=json.dumps(stub_planner_response),
                             fallback_steps=[], cache_hit=False)
    with patch("socya_pipeline.planner.AIChain") as MockChain:
        MockChain.return_value.call.return_value = fake_result
        # First call hits AI
        plan_presentation(wb, blocks, "p", "a", "es", "key", cache_dir=tmp_path,
                            file_path=ventas_simple_path)
        # Second call should hit cache
        plan2 = plan_presentation(wb, blocks, "p", "a", "es", "key", cache_dir=tmp_path,
                                    file_path=ventas_simple_path)
    assert plan2["_meta"]["cache_hit"] is True
    # Only one AI call across both invocations
    assert MockChain.return_value.call.call_count == 1
